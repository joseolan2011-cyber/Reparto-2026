/**
 * MonitorPedidos.gs
 * ARYS / Reparto 2026
 *
 * Vigila ENTRADASWEBOOK de forma independiente a Make/Treble.
 * Se monitorea por TELÉFONO, no por session_id.
 *
 * Reglas:
 * - Si aparece IDPedido, el intento se considera terminado correctamente.
 * - Si pasan MONITOR_MINUTOS_ALERTA (default 5) sin IDPedido, envía Pushover.
 * - Si pasan MONITOR_MINUTOS_CRITICO (default 10) sin IDPedido, envía alerta crítica.
 * - Si después de una alerta aparece IDPedido, envía notificación de recuperación.
 * - MONITOR_PEDIDOS guarda el estado para no repetir alertas.
 *
 * Script Properties requeridas:
 *   PUSHOVER_USER_KEY
 *   PUSHOVER_API_TOKEN
 *   MONITOR_MINUTOS_ALERTA   (opcional, default 5)
 *   MONITOR_MINUTOS_CRITICO  (opcional, default 10)
 */

const MONITOR_PEDIDOS_CONFIG = {
  HOJA_ENTRADAS: 'ENTRADASWEBOOK',
  HOJA_MONITOR: 'MONITOR_PEDIDOS',
  VENTANA_NUEVO_INTENTO_MIN: 30,
  HORAS_HISTORICO: 24,
  MINUTOS_ACTIVIDAD_RECIENTE: 2,
  PUSHOVER_URL: 'https://api.pushover.net/1/messages.json'
};

function monitorPedidos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shEntradas = ss.getSheetByName(MONITOR_PEDIDOS_CONFIG.HOJA_ENTRADAS);
    const shMonitor = ss.getSheetByName(MONITOR_PEDIDOS_CONFIG.HOJA_MONITOR);

    if (!shEntradas) throw new Error('No existe la hoja ENTRADASWEBOOK.');
    if (!shMonitor) throw new Error('No existe la hoja MONITOR_PEDIDOS.');

    const props = PropertiesService.getScriptProperties();
    const minAlerta = Number(props.getProperty('MONITOR_MINUTOS_ALERTA') || 5);
    const minCritico = Number(props.getProperty('MONITOR_MINUTOS_CRITICO') || 10);

    const eventos = monitorLeerEntradasRecientes_(shEntradas);
    if (!eventos.length) return;

    const intentos = monitorConstruirUltimoIntentoPorTelefono_(eventos);
    const estadoGuardado = monitorLeerEstado_(shMonitor);
    const ahora = new Date();

    Object.keys(intentos).forEach(telefono => {
      monitorProcesarIntento_(
        shMonitor,
        estadoGuardado[telefono] || null,
        intentos[telefono],
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

function monitorLeerEntradasRecientes_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

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

  const limite = new Date(Date.now() - MONITOR_PEDIDOS_CONFIG.HORAS_HISTORICO * 3600000);
  const eventos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const telefono = monitorNormalizarTelefono_(row[idx.telefono]);
    if (!telefono) continue;

    const fecha = monitorFecha_(row[idx.Fecha]) || monitorFecha_(row[idx.Entrada]);
    if (!fecha || fecha < limite) continue;

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

function monitorConstruirUltimoIntentoPorTelefono_(eventos) {
  const porTelefono = {};
  eventos.forEach(e => {
    if (!porTelefono[e.telefono]) porTelefono[e.telefono] = [];
    porTelefono[e.telefono].push(e);
  });

  const resultado = {};

  Object.keys(porTelefono).forEach(telefono => {
    const lista = porTelefono[telefono].sort((a, b) => a.fecha - b.fecha);
    const intentos = [];
    let actual = [];

    lista.forEach(e => {
      if (!actual.length) {
        actual.push(e);
        return;
      }

      const previo = actual[actual.length - 1];
      const gapMin = (e.fecha - previo.fecha) / 60000;

      if (gapMin > MONITOR_PEDIDOS_CONFIG.VENTANA_NUEVO_INTENTO_MIN) {
        intentos.push(actual);
        actual = [e];
      } else {
        actual.push(e);
      }
    });

    if (actual.length) intentos.push(actual);
    resultado[telefono] = monitorResumirIntento_(telefono, intentos[intentos.length - 1]);
  });

  return resultado;
}

function monitorResumirIntento_(telefono, eventos) {
  let nombre = '';
  let ultimoMensaje = '';
  let ultimaEleccionGPT = '';
  let ultimaSalida = '';
  let ultimaConversacion = '';
  let idPedido = '';
  let fechaPedido = null;

  eventos.forEach(e => {
    if (e.nombre) nombre = e.nombre;
    if (e.mensaje) ultimoMensaje = e.mensaje;
    if (e.eleccionGPT) ultimaEleccionGPT = e.eleccionGPT;
    if (e.salida) ultimaSalida = e.salida;
    if (e.conversacion) ultimaConversacion = e.conversacion;
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
    fechaPedido
  };
}

function monitorLeerEstado_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  rows.forEach((r, i) => {
    const telefono = monitorNormalizarTelefono_(r[0]);
    if (!telefono) return;

    map[telefono] = {
      fila: i + 2,
      inicio: monitorFecha_(r[1]),
      alerta5: monitorBool_(r[10]),
      alerta10: monitorBool_(r[11]),
      recuperado: monitorBool_(r[12])
    };
  });

  return map;
}

function monitorProcesarIntento_(sheet, previo, intento, ahora, minAlerta, minCritico) {
  const nuevoIntento = !previo || !previo.inicio ||
    Math.abs(intento.inicio.getTime() - previo.inicio.getTime()) > 60000;

  let fila = previo ? previo.fila : monitorBuscarFilaTelefono_(sheet, intento.telefono);
  let alerta5 = nuevoIntento ? false : previo.alerta5;
  let alerta10 = nuevoIntento ? false : previo.alerta10;
  let recuperado = nuevoIntento ? false : previo.recuperado;

  const minDesdeInicio = (ahora - intento.inicio) / 60000;
  const minSinActividad = (ahora - intento.ultimaActividad) / 60000;

  if (intento.idPedido) {
    let estado = 'COMPLETADO';

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
      estado = 'RECUPERADO';
    }

    monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado);
    return;
  }

  let estado = 'EN_PROCESO';

  if (minDesdeInicio >= minAlerta) {
    estado = minSinActividad <= MONITOR_PEDIDOS_CONFIG.MINUTOS_ACTIVIDAD_RECIENTE
      ? 'DEMORADO_ACTIVO'
      : 'POSIBLEMENTE_DETENIDO';
  }

  if (minDesdeInicio >= minAlerta && !alerta5) {
    monitorEnviarPushover_(
      '⚠️ Pedido sin completar',
      monitorMensajeAlerta_(intento, minDesdeInicio, minSinActividad, false),
      1
    );
    alerta5 = true;
  }

  if (minDesdeInicio >= minCritico && !alerta10) {
    monitorEnviarPushover_(
      '🚨 Pedido posiblemente detenido',
      monitorMensajeAlerta_(intento, minDesdeInicio, minSinActividad, true),
      1
    );
    alerta10 = true;
    estado = 'CRITICO';
  }

  monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado);
}

function monitorBuscarFilaTelefono_(sheet, telefono) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  const nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(monitorNormalizarTelefono_);
  const i = nums.indexOf(telefono);
  return i >= 0 ? i + 2 : lastRow + 1;
}

function monitorGuardarEstado_(sheet, fila, intento, estado, alerta5, alerta10, recuperado) {
  sheet.getRange(fila, 1, 1, 13).setValues([[
    intento.telefono,
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

  // Para prioridad normal (0), Pushover no necesita el parámetro.
  // Lo omitimos para evitar que Apps Script lo serialice de forma inválida.
  if (prioridadValida !== 0) {
    payload.priority = String(prioridadValida);
  }

  const resp = UrlFetchApp.fetch(MONITOR_PEDIDOS_CONFIG.PUSHOVER_URL, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Pushover HTTP ' + code + ': ' + resp.getContentText());
  }
}

function monitorMensajeAlerta_(intento, minDesdeInicio, minSinActividad, critica) {
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

function monitorMensajeRecuperado_(intento, minHastaPedido) {
  let m = '';
  if (intento.nombre) m += '👤 ' + intento.nombre + '\n';
  m += '📱 ' + intento.telefono + '\n';
  m += '✅ IDPedido: ' + intento.idPedido + '\n';
  m += '⏱️ Tiempo total: ' + monitorFormatearMinutos_(minHastaPedido) + '\n';
  m += '📥 Webhooks: ' + intento.cantidadWebhooks;
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
