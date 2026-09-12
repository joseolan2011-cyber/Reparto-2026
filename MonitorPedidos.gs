/**
 * MonitorPedidos.gs
 * ARYS / Reparto 2026
 *
 * Vigila ENTRADASWEBOOK de forma independiente a Make/Treble.
 * Se monitorea por TELÉFONO, no por session_id.
 *
 * REGLA DE ALCANCE:
 * - SOLO trabaja con eventos del DÍA DE HOY según la zona horaria de la hoja.
 * - Por cada teléfono conserva únicamente la interacción MÁS RECIENTE de hoy.
 * - No lee ni alerta conversaciones de ayer ni de días anteriores.
 *
 * VIGILA DOS CASOS:
 * 1) PEDIDO / PEDIDO_COMPLETO que no genera IDPedido.
 * 2) Conversación que inicia en SALUDO y no avanza a otro flujo.
 *
 * Antes de alertar, también revisa la hoja PEDIDOS:
 * - mismo teléfono
 * - pedido de hoy
 * - Fecha y hora_recibido >= inicio de la interacción
 * - pedido_cancelado distinto de TRUE
 *
 * MONITOR_PEDIDOS columnas A:N:
 * A telefono
 * B nombre
 * C inicio
 * D ultima_actividad
 * E cantidad_webhooks
 * F ultimo_mensaje
 * G ultima_eleccion_gpt
 * H ultima_salida
 * I ultima_conversacion
 * J IDPedido
 * K estado
 * L alerta_5min_enviada
 * M alerta_10min_enviada
 * N recuperado_notificado
 */

const MONITOR_PEDIDOS_CONFIG = {
  HOJA_ENTRADAS: 'ENTRADASWEBOOK',
  HOJA_MONITOR: 'MONITOR_PEDIDOS',
  HOJA_PEDIDOS: 'PEDIDOS',
  VENTANA_NUEVO_INTENTO_MIN: 30,
  MINUTOS_ACTIVIDAD_RECIENTE: 2,
  ELECCIONES_PEDIDO: ['PEDIDO', 'PEDIDO_COMPLETO'],
  ELECCION_SALUDO: 'SALUDO',
  PUSHOVER_URL: 'https://api.pushover.net/1/messages.json'
};

function monitorPedidos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shEntradas = ss.getSheetByName(MONITOR_PEDIDOS_CONFIG.HOJA_ENTRADAS);
    const shMonitor = ss.getSheetByName(MONITOR_PEDIDOS_CONFIG.HOJA_MONITOR);
    const shPedidos = ss.getSheetByName(MONITOR_PEDIDOS_CONFIG.HOJA_PEDIDOS);

    if (!shEntradas) throw new Error('No existe la hoja ENTRADASWEBOOK.');
    if (!shMonitor) throw new Error('No existe la hoja MONITOR_PEDIDOS.');
    if (!shPedidos) throw new Error('No existe la hoja PEDIDOS.');

    monitorLimpiarMonitorHoy_(shMonitor);

    const props = PropertiesService.getScriptProperties();
    const minAlerta = Number(props.getProperty('MONITOR_MINUTOS_ALERTA') || 5);
    const minCritico = Number(props.getProperty('MONITOR_MINUTOS_CRITICO') || 10);

    const eventos = monitorLeerEntradasDeHoy_(shEntradas);
    if (!eventos.length) return;

    const pedidosHoyPorTelefono = monitorLeerPedidosDeHoy_(shPedidos);
    const intentos = monitorConstruirUltimaInteraccionPorTelefono_(eventos);
    const estadoGuardado = monitorLeerEstado_(shMonitor);
    const ahora = new Date();

    Object.keys(intentos).forEach(telefono => {
      const intento = intentos[telefono];

      if (!intento.tieneIntentoPedido && !intento.iniciaConSaludo) return;

      // Cualquier interacción puede quedar resuelta si el pedido ya existe en PEDIDOS.
      if (!intento.idPedido) {
        const pedidoEnPedidos = monitorBuscarPedidoPosterior_(
          pedidosHoyPorTelefono[telefono] || [],
          intento.inicio
        );

        if (pedidoEnPedidos) {
          intento.idPedido = pedidoEnPedidos.idPedido;
          intento.fechaPedido = pedidoEnPedidos.fechaRecibido;
          intento.completadoDesdePedidos = true;
        }
      }

      monitorProcesarInteraccion_(
        shMonitor,
        estadoGuardado[telefono] || null,
        intento,
        ahora,
        minAlerta,
        minCritico
      );
    });

  } catch (err) {
    console.error('monitorPedidos:', err);
    try {
      monitorEnviarPushover_(
        '⚠️ Error Monitor Pedidos',
        'El monitor de Apps Script presentó un error:\n\n' + err.message,
        1
      );
    } catch (pushErr) {
      console.error('No se pudo enviar alerta de error:', pushErr);
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function monitorLeerEntradasDeHoy_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const hoyKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  ['telefono', 'Entrada', 'Mensaje Inicial', 'Eleccion GPT', 'Salida', 'Conversacion', 'IDPedido', 'Fecha']
    .forEach(col => {
      if (idx[col] === undefined) {
        throw new Error('Falta la columna ' + col + ' en ENTRADASWEBOOK.');
      }
    });

  const eventos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const telefono = monitorNormalizarTelefono_(row[idx.telefono]);
    if (!telefono) continue;

    const fecha = monitorFecha_(row[idx.Fecha]) || monitorFecha_(row[idx.Entrada]);
    if (!fecha) continue;

    const fechaKey = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    if (fechaKey !== hoyKey) continue;

    eventos.push({
      telefono,
      fecha,
      nombre: idx.nombre !== undefined ? monitorTexto_(row[idx.nombre]) : '',
      mensaje: monitorTexto_(row[idx['Mensaje Inicial']]),
      eleccionGPT: monitorTexto_(row[idx['Eleccion GPT']]),
      salida: monitorTexto_(row[idx.Salida]),
      conversacion: monitorTexto_(row[idx.Conversacion]),
      idPedido: monitorTexto_(row[idx.IDPedido])
    });
  }

  eventos.sort((a, b) => a.fecha - b.fecha);
  return eventos;
}

function monitorLeerPedidosDeHoy_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const hoyKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  ['id_pedido', 'telefono', 'fecha', 'Fecha y hora_recibido', 'pedido_cancelado']
    .forEach(col => {
      if (idx[col] === undefined) {
        throw new Error('Falta la columna ' + col + ' en PEDIDOS.');
      }
    });

  const porTelefono = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const telefono = monitorTexto_(row[idx.telefono]);
    if (!telefono) continue;

    const fechaPedidoDia = monitorFecha_(row[idx.fecha]);
    const fechaRecibido = monitorFecha_(row[idx['Fecha y hora_recibido']]);
    if (!fechaPedidoDia || !fechaRecibido) continue;

    const fechaKey = Utilities.formatDate(fechaPedidoDia, tz, 'yyyy-MM-dd');
    if (fechaKey !== hoyKey) continue;
    if (monitorBool_(row[idx.pedido_cancelado])) continue;

    const idPedido = monitorTexto_(row[idx.id_pedido]);
    if (!idPedido) continue;

    if (!porTelefono[telefono]) porTelefono[telefono] = [];
    porTelefono[telefono].push({ idPedido, fechaRecibido });
  }

  Object.keys(porTelefono).forEach(telefono => {
    porTelefono[telefono].sort((a, b) => a.fechaRecibido - b.fechaRecibido);
  });

  return porTelefono;
}

function monitorBuscarPedidoPosterior_(pedidosTelefono, inicioIntento) {
  if (!inicioIntento || !pedidosTelefono.length) return null;

  for (let i = 0; i < pedidosTelefono.length; i++) {
    if (pedidosTelefono[i].fechaRecibido >= inicioIntento) {
      return pedidosTelefono[i];
    }
  }

  return null;
}

/**
 * Agrupa eventos de hoy por teléfono, separa interacciones por huecos > 30 min
 * y conserva únicamente la interacción más reciente de cada teléfono.
 */
function monitorConstruirUltimaInteraccionPorTelefono_(eventos) {
  const porTelefono = {};

  eventos.forEach(e => {
    if (!porTelefono[e.telefono]) porTelefono[e.telefono] = [];
    porTelefono[e.telefono].push(e);
  });

  const resultado = {};

  Object.keys(porTelefono).forEach(telefono => {
    const lista = porTelefono[telefono].sort((a, b) => a.fecha - b.fecha);
    const interacciones = [];
    let actual = [];

    lista.forEach(e => {
      if (!actual.length) {
        actual.push(e);
        return;
      }

      const previo = actual[actual.length - 1];
      const gapMin = (e.fecha - previo.fecha) / 60000;

      if (gapMin > MONITOR_PEDIDOS_CONFIG.VENTANA_NUEVO_INTENTO_MIN) {
        interacciones.push(actual);
        actual = [e];
      } else {
        actual.push(e);
      }
    });

    if (actual.length) interacciones.push(actual);
    if (!interacciones.length) return;

    const resumen = monitorResumirInteraccion_(telefono, interacciones[interacciones.length - 1]);
    if (resumen.tieneIntentoPedido || resumen.iniciaConSaludo) {
      resultado[telefono] = resumen;
    }
  });

  return resultado;
}

function monitorResumirInteraccion_(telefono, eventos) {
  let nombre = '';
  let ultimoMensaje = '';
  let ultimaEleccionGPT = '';
  let ultimaSalida = '';
  let ultimaConversacion = '';
  let idPedido = '';
  let fechaPedido = null;
  let tieneIntentoPedido = false;
  let primeraEleccion = '';
  let fechaPrimerSaludo = null;
  let avanceDespuesSaludo = false;

  eventos.forEach(e => {
    if (e.nombre) nombre = e.nombre;
    if (e.mensaje) ultimoMensaje = e.mensaje;
    if (e.eleccionGPT) ultimaEleccionGPT = e.eleccionGPT;
    if (e.salida) ultimaSalida = e.salida;
    if (e.conversacion) ultimaConversacion = e.conversacion;

    const eleccion = monitorNormalizarEleccion_(e.eleccionGPT);

    if (!primeraEleccion && eleccion) primeraEleccion = eleccion;

    if (eleccion === MONITOR_PEDIDOS_CONFIG.ELECCION_SALUDO && !fechaPrimerSaludo) {
      fechaPrimerSaludo = e.fecha;
    }

    if (fechaPrimerSaludo && eleccion && eleccion !== MONITOR_PEDIDOS_CONFIG.ELECCION_SALUDO) {
      avanceDespuesSaludo = true;
    }

    if (MONITOR_PEDIDOS_CONFIG.ELECCIONES_PEDIDO.includes(eleccion)) {
      tieneIntentoPedido = true;
    }

    if (!idPedido && e.idPedido) {
      idPedido = e.idPedido;
      fechaPedido = e.fecha;
    }
  });

  return {
    telefono,
    nombre,
    inicio: eventos[0].fecha,
    ultimaActividad: eventos[eventos.length - 1].fecha,
    cantidadWebhooks: eventos.length,
    ultimoMensaje,
    ultimaEleccionGPT,
    ultimaSalida,
    ultimaConversacion,
    idPedido,
    fechaPedido,
    tieneIntentoPedido,
    iniciaConSaludo: primeraEleccion === MONITOR_PEDIDOS_CONFIG.ELECCION_SALUDO,
    fechaPrimerSaludo,
    avanceDespuesSaludo,
    completadoDesdePedidos: false
  };
}

function monitorLeerEstado_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const rows = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  rows.forEach((r, i) => {
    const telefono = monitorNormalizarTelefono_(r[0]);
    if (!telefono) return;

    map[telefono] = {
      fila: i + 2,
      inicio: monitorFecha_(r[2]),
      alerta5: monitorBool_(r[11]),
      alerta10: monitorBool_(r[12]),
      recuperado: monitorBool_(r[13])
    };
  });

  return map;
}

function monitorProcesarInteraccion_(sheet, previo, intento, ahora, minAlerta, minCritico) {
  const nuevoIntento = !previo || !previo.inicio ||
    Math.abs(intento.inicio.getTime() - previo.inicio.getTime()) > 60000;

  const fila = previo ? previo.fila : monitorBuscarFilaTelefono_(sheet, intento.telefono);
  let alerta5 = nuevoIntento ? false : previo.alerta5;
  let alerta10 = nuevoIntento ? false : previo.alerta10;
  let recuperado = nuevoIntento ? false : previo.recuperado;

  const minDesdeInicio = (ahora - intento.inicio) / 60000;
  const minSinActividad = (ahora - intento.ultimaActividad) / 60000;

  // Si ya existe pedido, la interacción está resuelta sin importar si comenzó como SALUDO.
  if (intento.idPedido) {
    let estado = intento.completadoDesdePedidos
      ? 'COMPLETADO_EN_PEDIDOS'
      : 'COMPLETADO';

    if ((alerta5 || alerta10) && !recuperado) {
      const minHastaPedido = intento.fechaPedido
        ? (intento.fechaPedido - intento.inicio) / 60000
        : minDesdeInicio;

      monitorEnviarPushover_(
        '✅ Pedido recuperado',
        monitorMensajeRecuperado_(intento, minHastaPedido),
        0
      );

      recuperado = true;
      estado = intento.completadoDesdePedidos
        ? 'RECUPERADO_EN_PEDIDOS'
        : 'RECUPERADO';
    }

    monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado);
    return;
  }

  // Si ya apareció PEDIDO / PEDIDO_COMPLETO, aplica el watchdog de pedido existente.
  if (intento.tieneIntentoPedido) {
    let estado = 'EN_PROCESO';

    if (minDesdeInicio >= minAlerta) {
      estado = minSinActividad <= MONITOR_PEDIDOS_CONFIG.MINUTOS_ACTIVIDAD_RECIENTE
        ? 'DEMORADO_ACTIVO'
        : 'POSIBLEMENTE_DETENIDO';
    }

    if (minDesdeInicio >= minAlerta && !alerta5) {
      monitorEnviarPushover_(
        '⚠️ Pedido sin completar',
        monitorMensajeAlertaPedido_(intento, minDesdeInicio, minSinActividad, false),
        1
      );
      alerta5 = true;
    }

    if (minDesdeInicio >= minCritico && !alerta10) {
      monitorEnviarPushover_(
        '🚨 Pedido posiblemente detenido',
        monitorMensajeAlertaPedido_(intento, minDesdeInicio, minSinActividad, true),
        1
      );
      alerta10 = true;
      estado = 'CRITICO';
    }

    monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado);
    return;
  }

  // SALUDO que sí avanzó a otro flujo (ASESOR, PROMOCION, etc.): no se considera detenido.
  if (intento.iniciaConSaludo && intento.avanceDespuesSaludo) {
    let estado = 'SALUDO_CONTINUO_OTRO_FLUJO';

    if (alerta5 && !recuperado) {
      monitorEnviarPushover_(
        '✅ Conversación reanudada',
        monitorMensajeSaludoRecuperado_(intento),
        0
      );
      recuperado = true;
      estado = 'SALUDO_RECUPERADO';
    }

    monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado);
    return;
  }

  // SALUDO sin continuación: una sola alerta a los 5 minutos.
  if (intento.iniciaConSaludo) {
    let estado = 'SALUDO_ESPERANDO_CONTINUACION';

    if (minDesdeInicio >= minAlerta) {
      estado = 'SALUDO_SIN_CONTINUACION';
    }

    if (minDesdeInicio >= minAlerta && !alerta5) {
      monitorEnviarPushover_(
        '⚠️ Conversación detenida después del saludo',
        monitorMensajeAlertaSaludo_(intento, minDesdeInicio),
        1
      );
      alerta5 = true;
    }

    monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, false, recuperado);
  }
}

/**
 * MONITOR_PEDIDOS queda limitado al día de hoy.
 * Conserva filas de PEDIDO y SALUDO; otros estados resueltos se eliminan en el siguiente ciclo.
 */
function monitorLimpiarMonitorHoy_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const hoyKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const rows = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const conservar = [];

  rows.forEach(r => {
    const telefono = monitorNormalizarTelefono_(r[0]);
    if (!telefono) return;

    const inicio = monitorFecha_(r[2]);
    if (!inicio) return;

    const inicioKey = Utilities.formatDate(inicio, tz, 'yyyy-MM-dd');
    if (inicioKey !== hoyKey) return;

    const eleccion = monitorNormalizarEleccion_(r[6]);
    const esPedido = MONITOR_PEDIDOS_CONFIG.ELECCIONES_PEDIDO.includes(eleccion);
    const esSaludo = eleccion === MONITOR_PEDIDOS_CONFIG.ELECCION_SALUDO;
    const estado = monitorTexto_(r[10]);
    const esResueltoDesdePedido = estado.indexOf('COMPLETADO') === 0 || estado.indexOf('RECUPERADO') >= 0;

    if (!esPedido && !esSaludo && !esResueltoDesdePedido) return;
    conservar.push(r);
  });

  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 14).clearContent();

  if (conservar.length) {
    sheet.getRange(2, 1, conservar.length, 14).setValues(conservar);
  }
}

function monitorBuscarFilaTelefono_(sheet, telefono) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  const nums = sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(monitorNormalizarTelefono_);

  const i = nums.indexOf(telefono);
  return i >= 0 ? i + 2 : lastRow + 1;
}

function monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado) {
  sheet.getRange(fila, 1, 1, 14).setValues([[
    intento.telefono,
    intento.nombre,
    intento.inicio,
    intento.ultimaActividad,
    intento.cantidadWebhooks,
    intento.ultimoMensaje,
    intento.ultimaEleccionGPT,
    intento.ultimaSalida,
    intento.ultimaConversacion,
    intento.idPedido,
    estado,
    alerta5,
    alerta10,
    recuperado
  ]]);
}

function monitorEnviarPushover_(titulo, mensaje, prioridad) {
  const props = PropertiesService.getScriptProperties();
  const user = props.getProperty('PUSHOVER_USER_KEY');
  const token = props.getProperty('PUSHOVER_API_TOKEN');

  if (!user) throw new Error('Falta Script Property PUSHOVER_USER_KEY.');
  if (!token) throw new Error('Falta Script Property PUSHOVER_API_TOKEN.');

  const prioridadNumero = Number(prioridad);
  const prioridadValida = [-2, -1, 0, 1, 2].includes(prioridadNumero)
    ? prioridadNumero
    : 0;

  const payload = {
    token: String(token),
    user: String(user),
    title: String(titulo || ''),
    message: String(mensaje || '')
  };

  if (prioridadValida !== 0) payload.priority = String(prioridadValida);

  const resp = UrlFetchApp.fetch(MONITOR_PEDIDOS_CONFIG.PUSHOVER_URL, {
    method: 'post',
    payload,
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Pushover HTTP ' + code + ': ' + resp.getContentText());
  }
}

function monitorMensajeAlertaPedido_(intento, minDesdeInicio, minSinActividad, critica) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const inicio = Utilities.formatDate(intento.inicio, tz, 'dd/MM/yyyy HH:mm:ss');
  const ultima = Utilities.formatDate(intento.ultimaActividad, tz, 'HH:mm:ss');
  const actividad = minSinActividad <= MONITOR_PEDIDOS_CONFIG.MINUTOS_ACTIVIDAD_RECIENTE
    ? '🟡 Sigue recibiendo actividad'
    : '🔴 Sin actividad reciente';

  let m = critica
    ? 'La interacción lleva más tiempo del normal sin generar pedido.\n\n'
    : 'Han pasado 5 minutos y todavía no existe IDPedido.\n\n';

  if (intento.nombre) m += '👤 ' + intento.nombre + '\n';
  m += '📱 ' + intento.telefono + '\n';
  m += '🕐 Inicio: ' + inicio + '\n';
  m += '⏱️ Tiempo: ' + monitorFormatearMinutos_(minDesdeInicio) + '\n';
  m += '📥 Webhooks: ' + intento.cantidadWebhooks + '\n';
  m += '🕒 Última actividad: ' + ultima + '\n';
  m += actividad + '\n';

  if (intento.ultimoMensaje) m += '\n💬 Último mensaje:\n' + monitorRecortar_(intento.ultimoMensaje, 180) + '\n';
  if (intento.ultimaEleccionGPT) m += '\n🤖 GPT: ' + monitorRecortar_(intento.ultimaEleccionGPT, 100);
  if (intento.ultimaSalida) m += '\n➡️ Salida: ' + monitorRecortar_(intento.ultimaSalida, 100);
  if (intento.ultimaConversacion) m += '\n🧭 Conversación: ' + monitorRecortar_(intento.ultimaConversacion, 100);

  m += '\n\n❌ IDPedido: NO GENERADO';
  return m;
}

function monitorMensajeAlertaSaludo_(intento, minDesdeInicio) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const inicio = Utilities.formatDate(intento.inicio, tz, 'dd/MM/yyyy HH:mm:ss');

  let m = 'La conversación comenzó con un saludo y no registró ninguna continuación después de 5 minutos.\n\n';
  if (intento.nombre) m += '👤 ' + intento.nombre + '\n';
  m += '📱 ' + intento.telefono + '\n';
  m += '🕐 Inicio: ' + inicio + '\n';
  m += '⏱️ Tiempo: ' + monitorFormatearMinutos_(minDesdeInicio) + '\n';
  m += '📥 Webhooks: ' + intento.cantidadWebhooks + '\n';

  if (intento.ultimoMensaje) m += '\n💬 Último mensaje:\n' + monitorRecortar_(intento.ultimoMensaje, 180) + '\n';
  if (intento.ultimaSalida) m += '\n➡️ Salida: ' + monitorRecortar_(intento.ultimaSalida, 100);
  if (intento.ultimaConversacion) m += '\n🧭 Conversación: ' + monitorRecortar_(intento.ultimaConversacion, 100);

  m += '\n\n⚠️ Revisar si Treble/Make respondió correctamente al cliente.';
  return m;
}

function monitorMensajeSaludoRecuperado_(intento) {
  let m = '';
  if (intento.nombre) m += '👤 ' + intento.nombre + '\n';
  m += '📱 ' + intento.telefono + '\n';
  m += '✅ La conversación volvió a avanzar.\n';
  if (intento.ultimaEleccionGPT) m += '🤖 Nuevo flujo: ' + intento.ultimaEleccionGPT + '\n';
  if (intento.ultimaSalida) m += '➡️ Salida: ' + monitorRecortar_(intento.ultimaSalida, 120) + '\n';
  m += '\nLa conversación que había quedado detenida en SALUDO volvió a tener continuidad.';
  return m;
}

function monitorMensajeRecuperado_(intento, minHastaPedido) {
  let m = '';
  if (intento.nombre) m += '👤 ' + intento.nombre + '\n';
  m += '📱 ' + intento.telefono + '\n';
  m += '✅ IDPedido: ' + intento.idPedido + '\n';
  m += '⏱️ Tiempo total: ' + monitorFormatearMinutos_(minHastaPedido) + '\n';
  m += '📥 Webhooks: ' + intento.cantidadWebhooks;
  if (intento.completadoDesdePedidos) m += '\n📋 Confirmado en tabla PEDIDOS';
  if (intento.ultimaConversacion) m += '\n🧭 Conversación: ' + intento.ultimaConversacion;
  m += '\n\nLa interacción que había generado alerta terminó correctamente.';
  return m;
}

function probarPushoverMonitor() {
  monitorEnviarPushover_(
    '✅ Monitor ARYS conectado',
    'Apps Script ya puede enviar notificaciones directamente a Pushover.',
    0
  );
}

function instalarTriggerMonitorPedidos() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monitorPedidos')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('monitorPedidos')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function eliminarTriggerMonitorPedidos() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monitorPedidos')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function monitorNormalizarTelefono_(v) {
  if (v === null || v === undefined || v === '') return '';
  return String(v).replace(/\D/g, '').trim();
}

function monitorNormalizarEleccion_(v) {
  return String(v || '').trim().toUpperCase();
}

function monitorFecha_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function monitorTexto_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function monitorBool_(v) {
  if (v === true) return true;
  const s = String(v || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'SI' || s === 'SÍ' || s === '1';
}

function monitorFormatearMinutos_(minutos) {
  const totalSeg = Math.max(0, Math.round(minutos * 60));
  const min = Math.floor(totalSeg / 60);
  const seg = totalSeg % 60;
  return min > 0 ? min + ' min ' + seg + ' seg' : seg + ' seg';
}

function monitorRecortar_(texto, max) {
  texto = String(texto || '');
  return texto.length <= max ? texto : texto.substring(0, max - 3) + '...';
}
