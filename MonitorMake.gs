/**
 * MonitorMake.gs
 * ARYS / Reparto 2026
 *
 * Vigila los escenarios de Make directamente desde Apps Script.
 * No depende de Make para alertar sobre Make.
 *
 * Requiere Script Properties:
 *   MAKE_API_TOKEN
 *   MAKE_ORGANIZATION_ID
 *
 * Opcional:
 *   MAKE_API_BASE_URL        (default: https://us2.make.com/api/v2)
 *   MAKE_MONITOR_INTERVAL_MIN (default: 5)
 *
 * Comportamiento:
 * - Consulta todos los escenarios de la organización.
 * - Ignora conceptos/borradores que nunca fueron activados.
 * - Considera problema si isActive=false o isPaused=true.
 * - En el primer chequeo avisa una sola vez si ya existen escenarios detenidos.
 * - Después solo avisa cambios de estado.
 * - Cuando un escenario vuelve a estar activo, manda recuperación una sola vez.
 */

const MONITOR_MAKE_CONFIG = {
  DEFAULT_API_BASE: 'https://us2.make.com/api/v2',
  DEFAULT_INTERVAL_MIN: 5,
  PROP_ULTIMO_CHECK: 'MAKE_MONITOR_ULTIMO_CHECK_MS',
  PROP_ESTADO: 'MAKE_MONITOR_ESTADO_V1',
  PAGE_LIMIT: 100,
  MAX_PAGES: 20
};

function monitorMakeEscenarios(forzar) {
  const props = PropertiesService.getScriptProperties();
  const ahoraMs = Date.now();
  const intervaloMin = Number(props.getProperty('MAKE_MONITOR_INTERVAL_MIN') || MONITOR_MAKE_CONFIG.DEFAULT_INTERVAL_MIN);
  const ultimoCheck = Number(props.getProperty(MONITOR_MAKE_CONFIG.PROP_ULTIMO_CHECK) || 0);

  if (!forzar && ultimoCheck && (ahoraMs - ultimoCheck) < intervaloMin * 60000) {
    return { omitido: true, motivo: 'INTERVALO', intervaloMin };
  }

  const token = String(props.getProperty('MAKE_API_TOKEN') || '').trim();
  const organizationId = String(props.getProperty('MAKE_ORGANIZATION_ID') || '').trim();
  const apiBase = String(props.getProperty('MAKE_API_BASE_URL') || MONITOR_MAKE_CONFIG.DEFAULT_API_BASE)
    .trim()
    .replace(/\/$/, '');

  if (!token) throw new Error('Falta Script Property MAKE_API_TOKEN.');
  if (!organizationId) throw new Error('Falta Script Property MAKE_ORGANIZATION_ID.');

  // Se marca el momento del intento antes de consultar para evitar una tormenta
  // de reintentos cada minuto si Make presenta una falla temporal.
  props.setProperty(MONITOR_MAKE_CONFIG.PROP_ULTIMO_CHECK, String(ahoraMs));

  const escenarios = monitorMakeListarEscenarios_(apiBase, token, organizationId)
    .filter(s => s && s.type !== 'tool' && s.concept !== true);

  const estadoAnterior = monitorMakeLeerEstado_(props);
  const primerChequeo = !estadoAnterior.__inicializado;
  const estadoNuevo = { __inicializado: true, __fecha: new Date().toISOString() };

  const detenidosActuales = [];
  const nuevosDetenidos = [];
  const recuperados = [];

  escenarios.forEach(s => {
    const id = String(s.id);
    const problema = s.isPaused === true || s.isActive === false;
    const razon = s.isPaused === true ? 'PAUSADO' : (s.isActive === false ? 'INACTIVO' : 'ACTIVO');

    estadoNuevo[id] = {
      id: s.id,
      name: s.name || ('Escenario ' + id),
      teamId: s.teamId || '',
      problema,
      razon
    };

    if (problema) {
      detenidosActuales.push(estadoNuevo[id]);
    }

    const previo = estadoAnterior[id];
    if (!primerChequeo && problema && (!previo || previo.problema !== true)) {
      nuevosDetenidos.push(estadoNuevo[id]);
    }

    if (!primerChequeo && !problema && previo && previo.problema === true) {
      recuperados.push(estadoNuevo[id]);
    }
  });

  // Si un escenario desaparece del listado no lo tratamos como recuperación:
  // pudo haberse eliminado o el usuario pudo perder acceso.

  if (primerChequeo && detenidosActuales.length) {
    monitorMakeEnviarProblemas_(
      '⚠️ Make: escenarios detenidos',
      detenidosActuales,
      'Se detectaron escenarios que actualmente no están activos.'
    );
  } else if (nuevosDetenidos.length) {
    monitorMakeEnviarProblemas_(
      '🚨 Make: escenario detenido',
      nuevosDetenidos,
      'Uno o más escenarios cambiaron a estado inactivo/pausado.'
    );
  }

  if (recuperados.length) {
    monitorMakeEnviarRecuperados_(recuperados);
  }

  props.setProperty(MONITOR_MAKE_CONFIG.PROP_ESTADO, JSON.stringify(estadoNuevo));

  const resultado = {
    ok: true,
    escenarios: escenarios.length,
    detenidos: detenidosActuales.length,
    nuevosDetenidos: nuevosDetenidos.length,
    recuperados: recuperados.length,
    primerChequeo
  };

  console.log('MONITOR MAKE: ' + JSON.stringify(resultado));
  return resultado;
}

function monitorMakeListarEscenarios_(apiBase, token, organizationId) {
  const todos = [];
  let offset = 0;

  for (let pagina = 0; pagina < MONITOR_MAKE_CONFIG.MAX_PAGES; pagina++) {
    const params = [
      'organizationId=' + encodeURIComponent(organizationId),
      'pg[limit]=' + MONITOR_MAKE_CONFIG.PAGE_LIMIT,
      'pg[offset]=' + offset,
      'cols[]=id',
      'cols[]=name',
      'cols[]=teamId',
      'cols[]=isActive',
      'cols[]=isPaused',
      'cols[]=concept',
      'cols[]=type'
    ];

    const url = apiBase + '/scenarios?' + params.join('&');
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Token ' + token,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    const body = resp.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error('Make API HTTP ' + code + ': ' + monitorRecortar_(body, 500));
    }

    let json;
    try {
      json = JSON.parse(body);
    } catch (e) {
      throw new Error('Make API devolvió JSON inválido.');
    }

    const lote = Array.isArray(json.scenarios) ? json.scenarios : [];
    todos.push.apply(todos, lote);

    if (lote.length < MONITOR_MAKE_CONFIG.PAGE_LIMIT) break;
    offset += lote.length;
  }

  return todos;
}

function monitorMakeLeerEstado_(props) {
  const raw = props.getProperty(MONITOR_MAKE_CONFIG.PROP_ESTADO);
  if (!raw) return {};

  try {
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (e) {
    return {};
  }
}

function monitorMakeEnviarProblemas_(titulo, escenarios, encabezado) {
  let mensaje = encabezado + '\n\n';
  const max = Math.min(escenarios.length, 12);

  for (let i = 0; i < max; i++) {
    const s = escenarios[i];
    mensaje += '• ' + s.name + '\n';
    mensaje += '  Estado: ' + s.razon + ' | ID: ' + s.id + '\n';
  }

  if (escenarios.length > max) {
    mensaje += '\n…y ' + (escenarios.length - max) + ' escenario(s) más.';
  }

  mensaje += '\n\nRevisión realizada directamente por Apps Script.';
  monitorEnviarPushover_(titulo, mensaje, 1);
}

function monitorMakeEnviarRecuperados_(escenarios) {
  let mensaje = 'Los siguientes escenarios volvieron a estar activos:\n\n';
  const max = Math.min(escenarios.length, 12);

  for (let i = 0; i < max; i++) {
    const s = escenarios[i];
    mensaje += '✅ ' + s.name + ' | ID: ' + s.id + '\n';
  }

  if (escenarios.length > max) {
    mensaje += '\n…y ' + (escenarios.length - max) + ' escenario(s) más.';
  }

  monitorEnviarPushover_('✅ Make: escenario recuperado', mensaje, 0);
}

/**
 * Prueba manual. Fuerza la consulta aunque todavía no hayan pasado 5 minutos.
 */
function probarMonitorMakeEscenarios() {
  return monitorMakeEscenarios(true);
}

/**
 * Reinicia únicamente la memoria del monitor de Make.
 * En el siguiente chequeo volverá a considerar el estado como primer chequeo.
 */
function reiniciarEstadoMonitorMake() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(MONITOR_MAKE_CONFIG.PROP_ESTADO);
  props.deleteProperty(MONITOR_MAKE_CONFIG.PROP_ULTIMO_CHECK);
  console.log('Estado del monitor Make reiniciado.');
}
