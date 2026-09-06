// ============================================
// SISTEMA DE PEDIDOS SUPERAGUA - PRODUCCIÓN v4.2
// Correcciones principales:
// 1) TOMAR PEDIDO queda simple: solo asigna vendedor, limpia el grupo y sale.
// 2) El script principal es el ÚNICO que manda mensajes con botones al vendedor.
// 3) Los botones de hora registran Hora_esti_entrega de inmediato y luego avisan a Make.
// 4) Todos los mensajes y respuestas de Telegram muestran la versión actual para diagnóstico visual.
// 5) Los botones usan formato principal 30_, 60_, 90_, 120_, 180_.
// 6) v3.8: base v3.7; NO tocar función principal, NO tocar botones, NO tocar Tomar pedido.
// 7) v3.8: solo se refuerza botón de hora: respuesta inmediata, quitar botones y anti-duplicados.
// 8) v3.9: SOLO se refuerza doPost: responde popup al INICIO absoluto del callback y filtra duplicados.
// 9) v4.0: Tomar pedido YA NO crea una fila nueva. Reusa la misma fila, la limpia y la deja Asignado.
// 10) v4.1: Pedidos del grupo muestran 30 min / 1 hora / 1.5 horas; al pulsar se toman y fijan la hora en una sola acción.
// 11) v4.1: El pedido tomado desde grupo conserva la misma fila y todos los datos del cliente; Procedencia = Tomado desde grupo.
// 12) v4.1: El aviso privado del pedido tomado desde grupo sale en el trigger de cada minuto, SIN botones, con datos completos y hora prometida.
// 13) v4.1: Al elegir hora en un pedido asignado directamente, se conservan visibles los datos completos del cliente.
// 14) v4.2: Los avisos de ENTREGA ATRASADA al vendedor salen por TELEGRAM_TOKEN_RETRASOS.
// ============================================

const VERSION = '4.2';
const HOJA_PEDIDOS = 'PEDIDOS';
const HOJA_PARAMETROS = 'PARAMETROS';
const ID_VENDEDOR_DEFAULT = 'contacto@superagua.com.mx';
const TZ_LOCAL = 'America/Merida';

// ============================================
// CONFIGURACIÓN
// La información sensible se lee desde Configuración / Propiedades del script.
// No pegues tokens ni webhooks directamente en el código.
// ============================================

function CFG() {
  const p = PropertiesService.getScriptProperties();

  return {
    TELEGRAM_TOKEN: p.getProperty('TELEGRAM_TOKEN'),
    TELEGRAM_TOKEN_RETRASOS: p.getProperty('TELEGRAM_TOKEN_RETRASOS'),
    GRUPO_PEDIDOS_ID: p.getProperty('GRUPO_PEDIDOS_ID'),
    GRUPO_MANDO_ID: p.getProperty('GRUPO_MANDO_ID'),
    WEBHOOK_MAKE: p.getProperty('WEBHOOK_MAKE'),
    URL_WEBAPP: p.getProperty('URL_WEBAPP'),

    RESPETAR_HORARIOS: normalizarSiNo_(p.getProperty('RESPETAR_HORARIOS') || 'SI'),

    HORARIO_VENDEDOR_INICIO: Number(p.getProperty('HORARIO_VENDEDOR_INICIO') || 7),
    HORARIO_VENDEDOR_FIN: Number(p.getProperty('HORARIO_VENDEDOR_FIN') || 22),

    HORARIO_GRUPO_INICIO: Number(p.getProperty('HORARIO_GRUPO_INICIO') || 7),
    HORARIO_GRUPO_FIN: Number(p.getProperty('HORARIO_GRUPO_FIN') || 19),

    HORARIO_RETRASOS_INICIO: Number(p.getProperty('HORARIO_RETRASOS_INICIO') || 8),
    HORARIO_RETRASOS_FIN: Number(p.getProperty('HORARIO_RETRASOS_FIN') || 22),

    BLOQUEAR_DOMINGO: normalizarSiNo_(p.getProperty('BLOQUEAR_DOMINGO') || 'SI'),
    BLOQUEAR_SABADO_DESDE: Number(p.getProperty('BLOQUEAR_SABADO_DESDE') || 19),
    BLOQUEAR_LUNES_ANTES: Number(p.getProperty('BLOQUEAR_LUNES_ANTES') || 7)
  };
}

function normalizarSiNo_(valor) {
  return String(valor || '')
    .trim()
    .toUpperCase()
    .replace('Í', 'I') === 'SI';
}

// ============================================
// WEBHOOK DE TELEGRAM
// ============================================

function doPost(e) {
  // v3.9: puerta rápida del webhook.
  // Objetivo: si Telegram manda un callback, responder popup INMEDIATO antes de cualquier trabajo pesado.
  // NO toca función principal, NO toca botones, NO toca Tomar pedido.
  try {
    console.log('=== WEBHOOK TELEGRAM RECIBIDO ' + VERSION + ' ===');

    if (!e || !e.postData || !e.postData.contents) {
      console.log('doPost sin postData. Sale rápido. ' + VERSION);
      return ok();
    }

    let data;

    try {
      data = JSON.parse(e.postData.contents);
    } catch (jsonErr) {
      console.log('JSON inválido en doPost ' + VERSION + ': ' + jsonErr.message);
      console.log(String(e.postData.contents).substring(0, 500));
      return ok();
    }

    const updateId = data.update_id !== undefined && data.update_id !== null
      ? String(data.update_id)
      : '';

    // Si no es botón, no debe hacer nada. /start y mensajes normales también entran por doPost.
    if (!data.callback_query) {
      const textoMsg = data.message && data.message.text
        ? String(data.message.text)
        : '';

      const chatMsg = data.message && data.message.chat
        ? String(data.message.chat.id)
        : '';

      console.log(
        'doPost sin callback_query. Ignorado. ' + VERSION +
        ' | update_id=' + updateId +
        ' | chat=' + chatMsg +
        ' | texto=' + textoMsg
      );

      return ok();
    }

    const cb = data.callback_query;
    const callbackId = String(cb.id || '');
    const callbackData = String(cb.data || '').trim();

    const fromId = cb.from && cb.from.id
      ? String(cb.from.id)
      : '';

    const chatId = cb.message && cb.message.chat
      ? String(cb.message.chat.id)
      : '';

    const messageId = cb.message && cb.message.message_id
      ? String(cb.message.message_id)
      : '';

    // =====================================================
    // RESPUESTA INMEDIATA ABSOLUTA
    // =====================================================
    // Esta es la diferencia de v3.9:
    // Antes el popup salía hasta que la función de hora empezaba.
    // Ahora cualquier botón recibe popup ANTES de cache, Sheets, Make o lógica pesada.
    try {
      responderCallbackTelegram(callbackId, '✅ Recibido');
    } catch (ePopup) {
      console.log('No se pudo lanzar popup inmediato ' + VERSION + ': ' + ePopup.message);
    }

    console.log(
      'CALLBACK ' + VERSION +
      ' | update_id=' + updateId +
      ' | callback_id=' + callbackId +
      ' | callback_data=' + callbackData +
      ' | from=' + fromId +
      ' | chat=' + chatId +
      ' | message_id=' + messageId
    );

    if (!callbackData) {
      console.log('Callback vacío. Sale rápido. ' + VERSION);
      return ok();
    }

    const cache = CacheService.getScriptCache();

    // =====================================================
    // BLOQUEO POR UPDATE_ID
    // =====================================================
    // Si Telegram reintenta el mismo update, no debe procesarse dos veces.
    if (updateId) {
      const updateKey = 'DP_UPDATE_' + updateId;

      if (cache.get(updateKey)) {
        console.log('Update duplicado ignorado ' + VERSION + ': ' + updateId);
        return ok();
      }

      cache.put(updateKey, 'SI', 300);
    }

    // =====================================================
    // BLOQUEO POR CALLBACK_ID
    // =====================================================
    // Cada toque trae callback_id. Si se repite, sale rápido.
    if (callbackId) {
      const cbKey = 'DP_CB_' + callbackId;

      if (cache.get(cbKey)) {
        console.log('callback_id duplicado ignorado ' + VERSION + ': ' + callbackId + ' | ' + callbackData);
        return ok();
      }

      cache.put(cbKey, 'SI', 300);
    }

    // =====================================================
    // BLOQUEO POR ACCIÓN EXACTA
    // =====================================================
    // Si el usuario pica muchas veces el mismo botón, pueden llegar callback_id distintos
    // pero con el mismo callback_data. Solo dejamos pasar el primero durante 60 segundos.
    const accionKey = 'DP_ACCION_' + callbackData;

    if (cache.get(accionKey)) {
      console.log('Acción duplicada ignorada ' + VERSION + ': ' + callbackData);
      return ok();
    }

    cache.put(accionKey, 'SI', 60);

    // =====================================================
    // ENRUTAMIENTO NORMAL
    // =====================================================
    // No se cambia la lógica interna de las funciones.
    // Solo se evita que entren duplicados y se responde popup al inicio.
    if (esCallbackGrupoConTiempo_(callbackData)) {
      console.log('Ruta doPost ' + VERSION + ': TOMAR DESDE GRUPO CON HORA | ' + callbackData);
      return tomarPedidoConHoraDesdeGrupo_(cb);
    }

    // Compatibilidad con mensajes viejos del grupo que todavía tengan el botón "Tomar pedido".
    if (callbackData.startsWith('tomar_')) {
      console.log('Ruta doPost ' + VERSION + ': TOMAR PEDIDO LEGACY | ' + callbackData);
      return asignarPedidoAVendedor(cb);
    }

    if (callbackData.startsWith('reasignar_')) {
      console.log('Ruta doPost ' + VERSION + ': REASIGNAR | ' + callbackData);
      return devolverPedidoAGrupo(cb);
    }

    if (esCallbackDeTiempo_(callbackData)) {
      console.log('Ruta doPost ' + VERSION + ': BOTÓN DE HORA | ' + callbackData);
      return notificarHoraAMake(cb);
    }

    console.log('Callback no reconocido ' + VERSION + ': ' + callbackData);
    return ok();

  } catch (err) {
    console.error('ERROR doPost ' + VERSION + ':', err.message);
    console.error(err.stack);
    return ok();
  }
}

function esCallbackDeTiempo_(callbackData) {
  return /^(30_|60_|90_|120_|180_|0\.5h_|1h_|1\.5h_|2h_|3h_)/.test(String(callbackData || ''));
}

function esCallbackGrupoConTiempo_(callbackData) {
  return /^(g30_|g60_|g90_)/.test(String(callbackData || ''));
}

function obtenerMinutosGrupoDesdeCallback_(callbackData) {
  const data = String(callbackData || '');

  if (data.startsWith('g30_')) return 30;
  if (data.startsWith('g60_')) return 60;
  if (data.startsWith('g90_')) return 90;

  return 0;
}

function extraerTokenGrupoDesdeCallback_(callbackData) {
  return String(callbackData || '')
    .replace(/^g(?:30|60|90)_/, '')
    .trim();
}

// ============================================
// EJECUTOR CADA MINUTO
// ============================================

function ejecutarMonitoreosCadaMinuto() {
  try {
    procesarPedidosPendientes();
  } catch (e) {
    console.log('Error en procesarPedidosPendientes:', e.message);
    console.log(e.stack);
  }

  try {
    revisarEntregasAtrasadas();
  } catch (e) {
    console.log('Error en revisarEntregasAtrasadas:', e.message);
    console.log(e.stack);
  }
}

// ============================================
// PROCESAR PEDIDOS PENDIENTES
// ============================================

function procesarPedidosPendientes() {
  // v3.7: ya no usa ScriptLock global.
  // Antes el trigger podía quedarse con el candado mientras mandaba Telegram o revisaba retrasos.
  // Eso hacía que los callbacks de pedidos robados se sintieran lentos.
  // Ahora se usa CacheService solo para evitar doble ejecución, pero sin bloquear botones.
  const cache = CacheService.getScriptCache();
  const runKey = 'RUN_PROCESAR_PEDIDOS_' + VERSION;

  if (cache.get(runKey)) {
    console.log('Otra ejecución de procesarPedidosPendientes sigue activa. Se cancela esta vuelta sin bloquear callbacks. ' + VERSION);
    return;
  }

  cache.put(runKey, 'SI', 55);

  try {
    const C = CFG();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);

    if (!sheet) {
      console.log('No existe hoja PEDIDOS.');
      return;
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return;
    }

    const headers = data[0];
    const idx = nombre => headers.indexOf(nombre);

    const idxHora = idx('Hora_esti_entrega');
    const idxFlag = idx('notificacion_entregada');
    const idxEjec = idx('ejecuciones_webhook');
    const idxPrimer = idx('fecha_primer_envio');
    const idxEstado = idx('estado');
    const idxChatId = idx('chatid');
    const idxMensajeId = idx('mensaje_telegram_id');
    const idxIdVendedor = idx('id_vendedor');
    const idxNombreVend = idx('Vendedor_Nombre');

    if (
      idxHora === -1 ||
      idxFlag === -1 ||
      idxEjec === -1 ||
      idxPrimer === -1 ||
      idxEstado === -1 ||
      idxChatId === -1 ||
      idxMensajeId === -1 ||
      idxIdVendedor === -1 ||
      idxNombreVend === -1
    ) {
      console.log('Faltan columnas necesarias en PEDIDOS.');
      return;
    }

    const ahora = new Date();

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const row = i + 1;

      const estado = String(fila[idxEstado] || '').trim();
      const horaEntrega = fila[idxHora];

      const payload = {};
      headers.forEach((h, j) => payload[h] = fila[j]);

      // =====================================================
      // -1) AVISO PRIVADO DE PEDIDO TOMADO DESDE GRUPO
      // =====================================================
      // El callback del grupo ya asignó vendedor + hora.
      // El trigger de cada minuto manda UNA SOLA VEZ el detalle privado, SIN botones.
      if (
        estado === 'En Camino' &&
        horaEntrega &&
        fila[idxChatId] &&
        esProcedenciaTomadoDesdeGrupo_(payload.Procedencia) &&
        !fila[idxMensajeId]
      ) {
        enviarAvisoPrivadoPedidoTomadoGrupo_(sheet, headers, row);
        continue;
      }

      // =====================================================
      // 0) PEDIDOS TOMADOS DESDE GRUPO
      // =====================================================
      // v3.7: Tomar pedido deja estado = Tomado.
      // El trigger principal lo convierte a Asignado y manda el mensaje oficial.
      // Así el pedido robado entra al mismo carril que el pedido original.
      if (estado === 'Tomado') {
        procesarPedidoTomado_(sheet, headers, row, payload, ahora);
        continue;
      }

      // =====================================================
      // 1) PEDIDOS SIN VENDEDOR
      // =====================================================
      if (estado === 'Sin Vendedor') {

        const estadoReal = String(sheet.getRange(row, idxEstado + 1).getValue() || '').trim();
        const horaReal = sheet.getRange(row, idxHora + 1).getValue();
        const idPedidoReal = obtenerIdPedidoDesdeFila_(sheet, row, headers);

        if (estadoReal !== 'Sin Vendedor' || horaReal) {
          console.log('Se evita reenviar. El pedido ya cambió. Pedido: ' + idPedidoReal);
          continue;
        }

        if (estaPedidoEnProcesoDeToma_(idPedidoReal)) {
          console.log('Se evita reenviar. Pedido en proceso de toma: ' + idPedidoReal);
          continue;
        }

        if (!esHorarioPermitidoParaGrupo(ahora)) {
          console.log('Pedido Sin Vendedor retenido fuera de horario de grupo. Pedido: ' + idPedidoReal);
          continue;
        }

        enviarPedidoDisponibleAlGrupo_(sheet, headers, row, payload, C);
        continue;
      }

      // =====================================================
      // 2) PEDIDOS ASIGNADOS SIN HORA
      // =====================================================
      if (estado === 'Asignado' && fila[idxChatId] && !horaEntrega) {

        const tokenAsignado = String(payload.callback_token || payload.id_pedido || '').trim();
        if (estaHoraEnProceso_(tokenAsignado)) {
          console.log('Se evita mandar botones. Hora en proceso para pedido/token: ' + tokenAsignado + ' ' + VERSION);
          continue;
        }

        if (!esHorarioPermitidoParaVendedor(ahora)) {
          console.log('Pedido asignado retenido fuera de horario vendedor. Pedido: ' + (payload.id_pedido || 'SIN_ID'));
          continue;
        }

        const estadoReal = String(sheet.getRange(row, idxEstado + 1).getValue() || '').trim();
        const horaReal = sheet.getRange(row, idxHora + 1).getValue();
        const chatIdReal = sheet.getRange(row, idxChatId + 1).getValue();

        if (estadoReal !== 'Asignado' || horaReal || !chatIdReal) {
          console.log('Se evita reenviar al vendedor. El pedido ya cambió. Pedido: ' + (payload.id_pedido || 'SIN_ID'));
          continue;
        }

        procesarPedidoAsignadoSinHora_(sheet, headers, row, payload, ahora);
      }
    }

  } finally {
    cache.remove(runKey);
  }
}

function procesarPedidoTomado_(sheet, headers, row, payload, ahora) {
  const idx = nombre => headers.indexOf(nombre);

  const idxEstado = idx('estado');
  const idxHora = idx('Hora_esti_entrega');
  const idxChatId = idx('chatid');
  const idxEjec = idx('ejecuciones_webhook');
  const idxMensajeId = idx('mensaje_telegram_id');
  const idxPrimer = idx('fecha_primer_envio');

  if (idxEstado === -1 || idxHora === -1 || idxChatId === -1 || idxEjec === -1 || idxMensajeId === -1 || idxPrimer === -1) {
    console.log('Faltan columnas para procesar pedido Tomado. ' + VERSION);
    return;
  }

  const idPedido = payload.id_pedido || obtenerIdPedidoDesdeFila_(sheet, row, headers) || 'SIN_ID';

  const estadoReal = String(sheet.getRange(row, idxEstado + 1).getValue() || '').trim();
  const horaReal = sheet.getRange(row, idxHora + 1).getValue();
  const chatIdReal = sheet.getRange(row, idxChatId + 1).getValue();

  if (estadoReal !== 'Tomado') {
    console.log('No se procesa Tomado. Estado real: ' + estadoReal + ' | Pedido: ' + idPedido);
    return;
  }

  if (horaReal) {
    console.log('No se procesa Tomado. Ya tiene hora. Pedido: ' + idPedido);
    return;
  }

  if (!chatIdReal) {
    console.log('No se procesa Tomado. Falta chatid. Pedido: ' + idPedido);
    return;
  }

  if (!esHorarioPermitidoParaVendedor(ahora)) {
    console.log('Pedido Tomado retenido fuera de horario vendedor. Pedido: ' + idPedido);
    return;
  }

  // Entra al mismo carril que un pedido original.
  sheet.getRange(row, idxEstado + 1).setValue('Asignado');
  sheet.getRange(row, idxEjec + 1).setValue(0);
  sheet.getRange(row, idxMensajeId + 1).setValue('');
  sheet.getRange(row, idxPrimer + 1).setValue('');
  SpreadsheetApp.flush();

  enviarMensajeTiempoAVendedor_(sheet, headers, row, {
    origen: 'TRIGGER_TOMADO_A_ASIGNADO',
    editarAnterior: false
  });
}

function enviarPedidoDisponibleAlGrupo_(sheet, headers, row, payload, C) {
  const idx = nombre => headers.indexOf(nombre);

  const idxMensajeId = idx('mensaje_telegram_id');
  const idxEjec = idx('ejecuciones_webhook');
  const idxPrimer = idx('fecha_primer_envio');

  const idPedido = payload.id_pedido || obtenerIdPedidoDesdeFila_(sheet, row, headers) || 'SIN_ID';
  const tokenBoton = obtenerOCrearCallbackToken_(sheet, headers, row, idPedido);

  const mensajeAnterior = idxMensajeId !== -1
    ? sheet.getRange(row, idxMensajeId + 1).getValue()
    : '';

  const intentosActuales = idxEjec !== -1
    ? Number(sheet.getRange(row, idxEjec + 1).getValue()) || 0
    : 0;

  const nuevoIntento = intentosActuales + 1;

  if (mensajeAnterior) {
    try {
      editarMensajeTelegram(
        C.GRUPO_PEDIDOS_ID,
        mensajeAnterior,
        `⌛ <b>Tiempo expirado ${VERSION}</b>\n\nUsa el mensaje más reciente para tomar este pedido.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar mensaje anterior del grupo:', e.message);
    }
  }

  const ubicacion = formatearUbicacion(payload.ubicacion);

  const mensaje =
    `🔔 <b>PEDIDO NUEVO ${VERSION}</b>\n\n` +
    `⏱ <b>Intento ${nuevoIntento}</b>\n\n` +
    `📦 Producto: ${payload.producto || 'Sin producto'} x${payload.cantidad || ''}\n` +
    `📍 Ubicación: ${ubicacion}\n` +
    `${formatearTelefonoConSeparadores(payload.telefono)}\n` +
    `🧾 Pedido: ${idPedido}\n` +
    `💬 Mensaje: ${payload['Mensaje Original'] || 'Sin mensaje'}\n` +
    `📝 Detalles: ${payload['Detalles_Pedido'] || 'Sin detalles'}\n\n` +
    `Primer vendedor que seleccione un tiempo se lo queda:`;

  // v4.1: en el grupo NO existe "No puedo atender".
  // Seleccionar el tiempo toma el pedido y fija la hora prometida en la misma acción.
  const botones = {
    inline_keyboard: [[
      { text: '30 min', callback_data: 'g30_' + String(tokenBoton) },
      { text: '1 hora', callback_data: 'g60_' + String(tokenBoton) },
      { text: '1.5 horas', callback_data: 'g90_' + String(tokenBoton) }
    ]]
  };

  const r = enviarMensajeTelegram(C.GRUPO_PEDIDOS_ID, mensaje, botones);

  if (r && r.ok && r.result && r.result.message_id && idxMensajeId !== -1) {
    sheet.getRange(row, idxMensajeId + 1).setValue(r.result.message_id);
  }

  if (idxPrimer !== -1 && !sheet.getRange(row, idxPrimer + 1).getValue()) {
    sheet.getRange(row, idxPrimer + 1).setValue(new Date());
  }

  if (idxEjec !== -1) {
    sheet.getRange(row, idxEjec + 1).setValue(nuevoIntento);
  }

  SpreadsheetApp.flush();
}

// ============================================
// ENVÍO OFICIAL DE MENSAJE AL VENDEDOR
// Esta es la ÚNICA función que debe mandar botones de tiempo al vendedor.
// En v3.7 la llama el trigger principal. Tomar pedido NO manda botones.
// ============================================

function enviarMensajeTiempoAVendedor_(sheet, headers, row, opciones) {
  opciones = opciones || {};

  const origen = opciones.origen || 'SIN_ORIGEN';
  const editarAnterior = opciones.editarAnterior !== false;

  const idx = nombre => headers.indexOf(nombre);

  const idxHora = idx('Hora_esti_entrega');
  const idxEjec = idx('ejecuciones_webhook');
  const idxPrimer = idx('fecha_primer_envio');
  const idxEstado = idx('estado');
  const idxChatId = idx('chatid');
  const idxMensajeId = idx('mensaje_telegram_id');
  const idxWebhookHora = idx('webhook_hora_enviado');

  if (
    idxHora === -1 ||
    idxEjec === -1 ||
    idxPrimer === -1 ||
    idxEstado === -1 ||
    idxChatId === -1 ||
    idxMensajeId === -1
  ) {
    console.log('Faltan columnas para enviar mensaje al vendedor. Origen: ' + origen);
    return { ok: false, motivo: 'FALTAN_COLUMNAS' };
  }

  const estadoReal = String(sheet.getRange(row, idxEstado + 1).getValue() || '').trim();
  const horaReal = sheet.getRange(row, idxHora + 1).getValue();
  const chatIdReal = sheet.getRange(row, idxChatId + 1).getValue();

  if (estadoReal !== 'Asignado') {
    console.log('No se envía mensaje al vendedor. Estado real: ' + estadoReal + ' | Origen: ' + origen);
    return { ok: false, motivo: 'ESTADO_NO_ASIGNADO' };
  }

  if (horaReal) {
    console.log('No se envía mensaje al vendedor. Ya tiene hora. Origen: ' + origen);
    return { ok: false, motivo: 'YA_TIENE_HORA' };
  }

  if (!chatIdReal) {
    console.log('No se envía mensaje al vendedor. Falta chatid. Origen: ' + origen);
    return { ok: false, motivo: 'SIN_CHATID' };
  }

  let intentos = Number(sheet.getRange(row, idxEjec + 1).getValue()) || 0;
  const nuevoIntento = intentos + 1;

  const mensajeAnterior = sheet.getRange(row, idxMensajeId + 1).getValue();

  if (editarAnterior && mensajeAnterior) {
    try {
      editarMensajeTelegram(
        chatIdReal,
        mensajeAnterior,
        `⌛ <b>Tiempo expirado ${VERSION}</b>\n\nUsa el mensaje más reciente.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar mensaje anterior del vendedor. Origen: ' + origen + ' | ' + e.message);
    }
  }

  const filaActual = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const payload = {};
  headers.forEach((h, j) => payload[h] = filaActual[j]);

  const idPedido = payload.id_pedido || obtenerIdPedidoDesdeFila_(sheet, row, headers) || 'SIN_ID';
  const tokenBoton = obtenerOCrearCallbackToken_(sheet, headers, row, idPedido);

  if (estaHoraEnProceso_(tokenBoton) || estaHoraEnProceso_(idPedido)) {
    console.log('No se envía mensaje al vendedor. Hora en proceso. Pedido: ' + idPedido + ' | Token: ' + tokenBoton);
    return { ok: false, motivo: 'HORA_EN_PROCESO' };
  }

  const mensaje = construirMensajePedidoAsignado_(payload, idPedido, nuevoIntento, origen, tokenBoton);
  const botones = construirBotonesTiempo_(tokenBoton, row);

  const r = enviarMensajeTelegram(chatIdReal, mensaje, botones);

  if (r && r.ok && r.result && r.result.message_id) {
    // Revisión de seguridad: si el vendedor alcanzó a pulsar hora mientras se enviaba este mensaje,
    // este mensaje recién enviado queda sin botones para que no haya mensajes fantasma.
    const horaDespues = sheet.getRange(row, idxHora + 1).getValue();
    if (horaDespues || estaHoraEnProceso_(tokenBoton) || estaHoraEnProceso_(idPedido)) {
      editarMensajeTelegram(
        chatIdReal,
        r.result.message_id,
        `✅ <b>Hora ya en proceso ${VERSION}</b>\n\n🧾 Pedido: ${html_(idPedido)}\n\nEste mensaje fue cancelado para evitar botones duplicados.`,
        null
      );
      return { ok: false, motivo: 'CANCELADO_POR_HORA_EN_PROCESO', message_id: r.result.message_id };
    }

    sheet.getRange(row, idxMensajeId + 1).setValue(r.result.message_id);
  } else {
    console.log('Telegram no confirmó envío al vendedor. Origen: ' + origen + ' | Respuesta: ' + JSON.stringify(r));
    return { ok: false, motivo: 'TELEGRAM_NO_OK', respuesta: r };
  }

  sheet.getRange(row, idxEjec + 1).setValue(nuevoIntento);

  if (!sheet.getRange(row, idxPrimer + 1).getValue()) {
    sheet.getRange(row, idxPrimer + 1).setValue(new Date());
  }

  SpreadsheetApp.flush();

  console.log('Mensaje oficial enviado al vendedor. Pedido: ' + idPedido + ' | Intento: ' + nuevoIntento + ' | Origen: ' + origen);

  return {
    ok: true,
    idPedido: idPedido,
    tokenBoton: tokenBoton,
    intento: nuevoIntento,
    message_id: r.result.message_id,
    origen: origen
  };
}

function procesarPedidoAsignadoSinHora_(sheet, headers, row, payload, ahora) {
  const idx = nombre => headers.indexOf(nombre);

  const idxHora = idx('Hora_esti_entrega');
  const idxEjec = idx('ejecuciones_webhook');
  const idxPrimer = idx('fecha_primer_envio');
  const idxEstado = idx('estado');
  const idxChatId = idx('chatid');
  const idxMensajeId = idx('mensaje_telegram_id');
  const idxIdVendedor = idx('id_vendedor');
  const idxNombreVend = idx('Vendedor_Nombre');
  const idxFechaAceptado = idx('Fechahora aceptado');
  const idxWebhookHora = idx('webhook_hora_enviado');

  const idPedido = payload.id_pedido || obtenerIdPedidoDesdeFila_(sheet, row, headers) || 'SIN_ID';

  let intentos = Number(sheet.getRange(row, idxEjec + 1).getValue()) || 0;
  const primerEnvio = sheet.getRange(row, idxPrimer + 1).getValue();

  if (intentos >= 10 && esHorarioPermitidoParaGrupo(ahora) && esPedidoDeNocheAnterior(primerEnvio, ahora)) {
    intentos = 0;
    sheet.getRange(row, idxEjec + 1).setValue(0);
    sheet.getRange(row, idxPrimer + 1).setValue(ahora);
    SpreadsheetApp.flush();
    console.log('Se reiniciaron los 10 minutos de tolerancia matutina para pedido: ' + idPedido + ' ' + VERSION);
  }

  const mensajeActivo = sheet.getRange(row, idxMensajeId + 1).getValue();

  if (intentos >= 10) {
    if (esHorarioPermitidoParaGrupo(ahora)) {
      if (mensajeActivo) {
        try {
          editarMensajeTelegram(
            sheet.getRange(row, idxChatId + 1).getValue(),
            mensajeActivo,
            `⌛ <b>Tiempo agotado ${VERSION}</b>\n\nPedido devuelto al grupo.`,
            null
          );
        } catch (e) {
          console.log('No se pudo editar mensaje del vendedor:', e.message);
        }
      }

      sheet.getRange(row, idxIdVendedor + 1).setValue(ID_VENDEDOR_DEFAULT);
      sheet.getRange(row, idxEstado + 1).setValue('Sin Vendedor');
      sheet.getRange(row, idxHora + 1).setValue('');
      sheet.getRange(row, idxChatId + 1).setValue('');
      sheet.getRange(row, idxNombreVend + 1).setValue('');
      sheet.getRange(row, idxMensajeId + 1).setValue('');
      sheet.getRange(row, idxEjec + 1).setValue(0);

      if (idxWebhookHora !== -1) {
        sheet.getRange(row, idxWebhookHora + 1).setValue('');
      }

      const idxNotificacionLocal = headers.indexOf('notificacion_entregada');
      if (idxNotificacionLocal !== -1) {
        sheet.getRange(row, idxNotificacionLocal + 1).setValue('No');
      }

      if (idxFechaAceptado !== -1) {
        sheet.getRange(row, idxFechaAceptado + 1).setValue('');
      }

      SpreadsheetApp.flush();
      console.log('Pedido quitado al vendedor y preparado para grupo: ' + idPedido + ' ' + VERSION);
    }

    return;
  }

  enviarMensajeTiempoAVendedor_(sheet, headers, row, {
    origen: 'TRIGGER_ASIGNADO',
    editarAnterior: true
  });
}


// ============================================
// TOMAR PEDIDO DESDE GRUPO + FIJAR HORA
// v4.1: nuevo flujo para los tres botones del grupo.
// NO manda mensaje privado aquí: lo hace el trigger de cada minuto.
// ============================================

function tomarPedidoConHoraDesdeGrupo_(cb) {
  const cache = CacheService.getScriptCache();
  const callbackId = String(cb.id || '');
  const callbackData = String(cb.data || '').trim();
  const tokenPedido = extraerTokenGrupoDesdeCallback_(callbackData);
  const minutos = obtenerMinutosGrupoDesdeCallback_(callbackData);

  if (!tokenPedido || !minutos) {
    try {
      responderCallbackTelegram(callbackId, '❌ Botón inválido');
    } catch (e) {
      console.log('No se pudo responder botón de grupo inválido:', e.message);
    }
    return ok();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = ss.getSheetByName(HOJA_PEDIDOS);
  const parametros = ss.getSheetByName(HOJA_PARAMETROS);

  if (!pedidos || !parametros) {
    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `❌ <b>Error ${VERSION}</b>\n\nNo encontré PEDIDOS o PARAMETROS.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar error de hojas en grupo:', e.message);
    }
    return ok();
  }

  const telegramUserId = String(cb.from && cb.from.id ? cb.from.id : '').trim();
  const vendedor = buscarVendedorPorTelegram(parametros, telegramUserId);

  if (!vendedor) {
    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `❌ <b>Vendedor no registrado ${VERSION}</b>\n\nNo estás registrado como vendedor activo.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar vendedor no registrado en grupo:', e.message);
    }
    return ok();
  }

  // Doble protección local. Se activa solo después de confirmar que quien pulsó
  // el botón sí es un vendedor registrado, para no bloquear el pedido por un toque ajeno.
  const keyGrupo = 'GRUPO_HORA_' + tokenPedido;
  if (cache.get(keyGrupo)) {
    try {
      responderCallbackTelegram(callbackId, '⏳ Este pedido ya se está tomando');
    } catch (e) {
      console.log('No se pudo responder pedido grupo en proceso:', e.message);
    }
    return ok();
  }

  cache.put(keyGrupo, 'SI', 120);
  marcarPedidoEnProcesoDeToma_(tokenPedido);
  marcarHoraEnProceso_(tokenPedido);

  const lock = LockService.getScriptLock();
  let resultado = null;

  if (!lock.tryLock(2500)) {
    cache.remove(keyGrupo);
    try {
      responderCallbackTelegram(callbackId, '⚠ Sistema ocupado. Intenta otra vez.');
    } catch (e) {
      console.log('No se pudo responder sistema ocupado grupo:', e.message);
    }
    return ok();
  }

  try {
    const filaPedido = buscarFilaPedidoPorTokenRapido_(
      pedidos,
      obtenerHeadersRapido_(pedidos),
      tokenPedido
    );

    if (filaPedido === -1) {
      resultado = {
        ok: false,
        tipo: 'NO_ENCONTRADO',
        mensaje: `❌ <b>Pedido no encontrado ${VERSION}</b>\n\nIntenta usar el mensaje más reciente.`
      };
    } else {
      const headers = obtenerHeadersRapido_(pedidos);
      const idxEstado = headers.indexOf('estado');
      const idxHora = headers.indexOf('Hora_esti_entrega');
      const idxIdPedido = headers.indexOf('id_pedido');
      const idxWebhookHora = headers.indexOf('webhook_hora_enviado');

      if (idxEstado === -1 || idxHora === -1 || idxWebhookHora === -1) {
        resultado = {
          ok: false,
          tipo: 'FALTAN_COLUMNAS',
          mensaje: `❌ <b>Faltan columnas ${VERSION}</b>\n\nSe requieren estado, Hora_esti_entrega y webhook_hora_enviado.`
        };
      } else {
        const estadoActual = String(
          pedidos.getRange(filaPedido, idxEstado + 1).getValue() || ''
        ).trim();

        const horaActual = pedidos.getRange(filaPedido, idxHora + 1).getValue();
        const idPedido = idxIdPedido !== -1
          ? pedidos.getRange(filaPedido, idxIdPedido + 1).getValue()
          : tokenPedido;

        if (estadoActual !== 'Sin Vendedor' || horaActual) {
          resultado = {
            ok: false,
            tipo: 'YA_TOMADO',
            idPedido: idPedido,
            mensaje:
              `❌ <b>Pedido tomado por otro vendedor ${VERSION}</b>\n\n` +
              `🧾 Pedido: ${html_(idPedido)}\n` +
              `Estado actual: ${html_(estadoActual)}`
          };
        } else {
          resultado = asignarMismaFilaAVendedorConHoraDesdeGrupo_(
            pedidos,
            headers,
            filaPedido,
            vendedor,
            tokenPedido,
            minutos
          );
        }
      }
    }

    SpreadsheetApp.flush();

  } catch (e) {
    console.log('Error en tomarPedidoConHoraDesdeGrupo_ ' + VERSION + ':', e.message);
    console.log(e.stack);

    resultado = {
      ok: false,
      tipo: 'ERROR',
      mensaje: `❌ <b>Error al tomar pedido ${VERSION}</b>\n\n${html_(e.message)}`
    };
  } finally {
    lock.releaseLock();
  }

  // El mensaje del grupo se limpia como ya ocurría antes:
  // NO se dejan nombre, teléfono, ubicación, mensaje ni detalles del cliente.
  try {
    if (resultado && resultado.ok) {
      const vendedorMostrar = String(vendedor.nombre || vendedor.id_vendedor || 'Sin nombre')
        .replace(/\bnull\b/gi, '')
        .replace(/\bundefined\b/gi, '')
        .trim() || 'Sin nombre';

      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `✅ <b>Pedido tomado ${VERSION}</b>\n\n` +
        `🧾 Pedido: ${html_(resultado.idPedido)}\n` +
        `👤 Vendedor: ${html_(vendedorMostrar)}\n` +
        `⏱ Tiempo comprometido: ${html_(resultado.tiempoTexto)}\n` +
        `🕒 Hora prometida: ${html_(resultado.horaTexto)}`,
        null
      );

      console.log(
        'Pedido tomado desde grupo con hora ' + VERSION +
        ' | Pedido: ' + resultado.idPedido +
        ' | Fila: ' + resultado.fila +
        ' | Vendedor: ' + vendedor.id_vendedor +
        ' | Hora: ' + resultado.horaTexto
      );

      // Make conserva exactamente su función: avisar al cliente de la hora prometida.
      // Se hace fuera del lock para no bloquear a otros vendedores.
      enviarHoraAMakeDesdeGrupo_(cb, resultado);
    } else {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        resultado && resultado.mensaje
          ? resultado.mensaje
          : `❌ <b>No se pudo tomar pedido ${VERSION}</b>`,
        null
      );
    }
  } catch (e) {
    console.log('No se pudo cerrar mensaje del grupo después de tomar con hora:', e.message);
  }

  return ok();
}

function asignarMismaFilaAVendedorConHoraDesdeGrupo_(
  sheet,
  headers,
  filaPedido,
  vendedor,
  tokenPedido,
  minutos
) {
  const lastCol = sheet.getLastColumn();
  const ahora = new Date();
  const horaEstimada = new Date(ahora.getTime() + (minutos * 60000));
  const valores = sheet.getRange(filaPedido, 1, 1, lastCol).getValues()[0];

  const idPedido = obtenerValorArray_(valores, headers, 'id_pedido') || tokenPedido || 'SIN_ID';

  // Mismos datos operativos del flujo normal después de que el vendedor
  // acepta el pedido y selecciona la hora. Todos los demás datos de la fila
  // (cliente, teléfono, ubicación, producto, mensaje, detalles, etc.) SE CONSERVAN.
  setValorArraySiExiste_(valores, headers, 'estado', 'En Camino');
  setValorArraySiExiste_(valores, headers, 'id_vendedor', vendedor.id_vendedor);
  setValorArraySiExiste_(valores, headers, 'Vendedor_Nombre', vendedor.nombre);
  setValorArraySiExiste_(valores, headers, 'chatid', vendedor.chatid);
  setValorArraySiExiste_(valores, headers, 'Fechahora aceptado', ahora);
  setValorArraySiExiste_(valores, headers, 'Hora_esti_entrega', horaEstimada);
  setValorArraySiExiste_(valores, headers, 'hora_entrega_real', '');

  // mensaje_telegram_id queda vacío a propósito:
  // el trigger de cada minuto detecta esto y manda el detalle privado SIN botones.
  setValorArraySiExiste_(valores, headers, 'mensaje_telegram_id', '');

  // Igual que el botón normal de hora: ya no debe volver a enviar botones/reintentos.
  setValorArraySiExiste_(valores, headers, 'ejecuciones_webhook', 99);
  setValorArraySiExiste_(valores, headers, 'avisos_retraso_entrega', 0);
  setValorArraySiExiste_(valores, headers, 'notificacion_entregada', 'No');

  // Igual que el flujo normal: se marca ANTES de llamar a Make para evitar duplicados.
  setValorArraySiExiste_(valores, headers, 'webhook_hora_enviado', ahora);

  // Requisito específico del grupo.
  setValorArraySiExiste_(valores, headers, 'Procedencia', 'Tomado desde grupo');

  const idxToken = headers.indexOf('callback_token');
  if (idxToken !== -1 && !String(valores[idxToken] || '').trim()) {
    valores[idxToken] = generarCallbackToken_();
  }

  sheet.getRange(filaPedido, 1, 1, lastCol).setValues([valores]);

  const idxHora = headers.indexOf('Hora_esti_entrega');
  const idxWebhookHora = headers.indexOf('webhook_hora_enviado');

  if (idxHora !== -1) {
    sheet
      .getRange(filaPedido, idxHora + 1)
      .setNumberFormat('dd/MM/yyyy hh:mm AM/PM');
  }

  if (idxWebhookHora !== -1) {
    sheet
      .getRange(filaPedido, idxWebhookHora + 1)
      .setNumberFormat('dd/MM/yyyy hh:mm AM/PM');
  }

  marcarPedidoEnProcesoDeToma_(idPedido);
  marcarPedidoEnProcesoDeToma_(tokenPedido);
  marcarHoraEnProceso_(idPedido);
  marcarHoraEnProceso_(tokenPedido);

  return {
    ok: true,
    idPedido: idPedido,
    token: idxToken !== -1 ? valores[idxToken] : tokenPedido,
    fila: filaPedido,
    minutos: minutos,
    tiempoTexto: textoTiempoDesdeMinutos_(minutos),
    horaEstimada: horaEstimada,
    horaTexto: formatearFechaParaWhatsApp_(horaEstimada),
    webhookHoraEnviado: ahora
  };
}

function enviarHoraAMakeDesdeGrupo_(cb, resultado) {
  try {
    const C = CFG();

    if (!C.WEBHOOK_MAKE) {
      console.log(
        'Make NO se envía desde grupo porque WEBHOOK_MAKE está vacío. Pedido: ' +
        resultado.idPedido + ' ' + VERSION
      );
      return;
    }

    const pedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);
    if (!pedidos) {
      console.log('Make NO se envía desde grupo porque no existe PEDIDOS. ' + VERSION);
      return;
    }

    const headers = obtenerHeadersRapido_(pedidos);
    const filaActual = pedidos
      .getRange(resultado.fila, 1, 1, headers.length)
      .getValues()[0];

    const pedidoCompleto = {};
    headers.forEach((h, j) => pedidoCompleto[h] = filaActual[j]);

    const payloadMake = {
      callback_query: cb,
      callback_data: String(cb.data || ''),
      telegram_user_id: cb.from && cb.from.id ? cb.from.id : '',
      telegram_user_name: cb.from && cb.from.first_name ? cb.from.first_name : '',
      chat_id: cb.message && cb.message.chat ? cb.message.chat.id : '',
      message_id: cb.message ? cb.message.message_id : '',
      id_pedido: resultado.idPedido,
      token_pedido: resultado.token,
      minutos: resultado.minutos,
      tiempo_texto: resultado.tiempoTexto,

      hora_estimada: resultado.horaEstimada,
      hora_estimada_iso: resultado.horaEstimada.toISOString(),
      hora_estimada_texto: resultado.horaTexto,
      hora_estimada_whatsapp: resultado.horaTexto,

      fecha_callback: new Date(),
      actualizado_por_apps_script: true,
      version_apps_script: VERSION,
      origen_apps_script: 'TOMADO_DESDE_GRUPO_CON_HORA',
      webhook_hora_enviado: resultado.webhookHoraEnviado,
      pedido: pedidoCompleto
    };

    const respuesta = UrlFetchApp.fetch(C.WEBHOOK_MAKE, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payloadMake),
      muteHttpExceptions: true
    });

    console.log(
      'Hora de pedido tomado desde grupo enviada a Make UNA SOLA VEZ ' + VERSION +
      '. Código: ' + respuesta.getResponseCode() +
      ' | Pedido: ' + resultado.idPedido +
      ' | Respuesta: ' + respuesta.getContentText()
    );

  } catch (e) {
    // Igual que el flujo normal: no se borra webhook_hora_enviado para evitar duplicados.
    console.log(
      'Make falló para pedido tomado desde grupo, pero webhook_hora_enviado queda marcado ' +
      VERSION + ': ' + e.message
    );
  }
}

function esProcedenciaTomadoDesdeGrupo_(procedencia) {
  return String(procedencia || '')
    .trim()
    .toLowerCase()
    .startsWith('tomado desde grupo');
}

function enviarAvisoPrivadoPedidoTomadoGrupo_(sheet, headers, row) {
  const idx = nombre => headers.indexOf(nombre);

  const idxEstado = idx('estado');
  const idxHora = idx('Hora_esti_entrega');
  const idxChatId = idx('chatid');
  const idxMensajeId = idx('mensaje_telegram_id');
  const idxProcedencia = idx('Procedencia');

  if (
    idxEstado === -1 ||
    idxHora === -1 ||
    idxChatId === -1 ||
    idxMensajeId === -1 ||
    idxProcedencia === -1
  ) {
    console.log('Faltan columnas para aviso privado de pedido tomado desde grupo. ' + VERSION);
    return;
  }

  // Relectura real para evitar que dos ejecuciones del minuto manden el mismo mensaje.
  const filaReal = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const estadoReal = String(filaReal[idxEstado] || '').trim();
  const horaReal = filaReal[idxHora];
  const chatIdReal = filaReal[idxChatId];
  const mensajeIdReal = filaReal[idxMensajeId];
  const procedenciaReal = filaReal[idxProcedencia];

  if (
    estadoReal !== 'En Camino' ||
    !horaReal ||
    !chatIdReal ||
    mensajeIdReal ||
    !esProcedenciaTomadoDesdeGrupo_(procedenciaReal)
  ) {
    return;
  }

  const payload = {};
  headers.forEach((h, j) => payload[h] = filaReal[j]);

  const idPedido = payload.id_pedido || obtenerIdPedidoDesdeFila_(sheet, row, headers) || 'SIN_ID';
  const horaEstimada = parsearFechaEntrega(horaReal);

  if (!horaEstimada) {
    console.log('No se manda aviso privado del grupo: Hora_esti_entrega inválida. Pedido: ' + idPedido);
    return;
  }

  const mensaje = construirMensajePrivadoTomadoDesdeGrupo_(
    payload,
    idPedido,
    horaEstimada
  );

  const r = enviarMensajeTelegram(chatIdReal, mensaje, null);

  if (r && r.ok && r.result && r.result.message_id) {
    sheet.getRange(row, idxMensajeId + 1).setValue(r.result.message_id);
    SpreadsheetApp.flush();

    console.log(
      'Aviso privado SIN botones enviado al vendedor por pedido tomado desde grupo ' + VERSION +
      ' | Pedido: ' + idPedido +
      ' | Chat: ' + chatIdReal
    );
  } else {
    console.log(
      'Telegram no confirmó aviso privado de pedido tomado desde grupo. ' + VERSION +
      ' | Pedido: ' + idPedido +
      ' | Respuesta: ' + JSON.stringify(r)
    );
  }
}

function construirMensajePrivadoTomadoDesdeGrupo_(payload, idPedido, horaEstimada) {
  const ubicacion = formatearUbicacion(payload.ubicacion);
  const nombreCliente = String(payload.nombre || '').trim();
  const lineaCliente = nombreCliente
    ? `👤 Cliente: ${html_(nombreCliente)}\n`
    : '';

  return (
    `📋 <b>PEDIDO TOMADO DESDE GRUPO ${VERSION}</b>\n\n` +
    lineaCliente +
    `📦 Producto: ${html_(payload.producto || 'Sin producto')} x${html_(payload.cantidad || '')}\n` +
    `📍 Ubicación: ${ubicacion}\n` +
    `${formatearTelefonoConSeparadores(payload.telefono)}\n` +
    `🧾 Pedido: ${html_(idPedido)}\n` +
    `💬 Mensaje Original: ${html_(payload['Mensaje Original'] || 'Sin mensaje')}\n` +
    `📝 Detalles: ${html_(payload['Detalles_Pedido'] || 'Sin detalles')}\n\n` +
    `🕒 <b>HORA PROMETIDA: ${html_(formatearFechaParaWhatsApp_(horaEstimada))}</b>\n\n` +
    `⚠️ <b>IMPORTANTE: CUMPLE CON EL TIEMPO ESTABLECIDO.</b>\n` +
    `Al tomar este pedido te comprometiste con esta hora de entrega. Organiza tu ruta para cumplirla.`
  );
}

// ============================================
// TOMAR PEDIDO DESDE GRUPO
// ============================================

function asignarPedidoAVendedor(cb) {
  const cache = CacheService.getScriptCache();
  const callbackId = String(cb.id || '');
  const tokenPedido = String(cb.data || '').replace('tomar_', '').trim();

  // v4.0: respuesta inmediata. Tomar pedido debe sentirse rápido.
  responderCallbackTelegram(callbackId, '✅ Tomando pedido ' + VERSION);

  if (!tokenPedido) {
    return ok();
  }

  if (cache.get('CB_' + callbackId)) {
    console.log('Callback tomar duplicado ignorado:', callbackId);
    return ok();
  }

  cache.put('CB_' + callbackId, 'SI', 300);
  marcarPedidoEnProcesoDeToma_(tokenPedido);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = ss.getSheetByName(HOJA_PEDIDOS);
  const parametros = ss.getSheetByName(HOJA_PARAMETROS);

  if (!pedidos || !parametros) {
    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `❌ <b>Error ${VERSION}</b>

No encontré PEDIDOS o PARAMETROS.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar error de hojas:', e.message);
    }
    return ok();
  }

  const telegramUserId = String(cb.from && cb.from.id ? cb.from.id : '').trim();
  const vendedor = buscarVendedorPorTelegram(parametros, telegramUserId);

  if (!vendedor) {
    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `❌ <b>Vendedor no registrado ${VERSION}</b>

No estás registrado como vendedor activo.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar vendedor no registrado:', e.message);
    }
    return ok();
  }

  let resultado = null;

  // v4.0: lock corto solo para escribir la misma fila.
  // Ya NO se crea fila nueva. Esto evita el costo de appendRow y evita que el pedido nuevo nazca "caliente".
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(2500)) {
    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `⚠ <b>Sistema ocupado ${VERSION}</b>

Intenta tocar el botón otra vez en unos segundos.`,
        null
      );
    } catch (e) {
      console.log('No se pudo avisar sistema ocupado:', e.message);
    }
    return ok();
  }

  try {
    const filaPedido = buscarFilaPedidoPorToken(pedidos, tokenPedido);

    if (filaPedido === -1) {
      resultado = {
        ok: false,
        tipo: 'NO_ENCONTRADO',
        mensaje: `❌ <b>Pedido no encontrado ${VERSION}</b>\n\nIntenta usar el mensaje más reciente.`
      };
    } else {
      const headers = obtenerHeadersRapido_(pedidos);
      const idxEstado = headers.indexOf('estado');
      const idxIdPedido = headers.indexOf('id_pedido');

      if (idxEstado === -1) {
        resultado = {
          ok: false,
          tipo: 'FALTAN_COLUMNAS',
          mensaje: `❌ <b>Falta columna estado ${VERSION}</b>`
        };
      } else {
        const estadoActual = String(pedidos.getRange(filaPedido, idxEstado + 1).getValue() || '').trim();
        const idPedido = idxIdPedido !== -1
          ? pedidos.getRange(filaPedido, idxIdPedido + 1).getValue()
          : tokenPedido;

        if (estadoActual !== 'Sin Vendedor') {
          resultado = {
            ok: false,
            tipo: 'YA_TOMADO',
            idPedido: idPedido,
            mensaje: `❌ <b>Pedido tomado por otro vendedor ${VERSION}</b>\n\n🧾 Pedido: ${html_(idPedido)}\nEstado actual: ${html_(estadoActual)}`
          };
        } else {
          resultado = asignarMismaFilaAVendedor_(pedidos, headers, filaPedido, vendedor, tokenPedido);
        }
      }
    }

    SpreadsheetApp.flush();

  } catch (e) {
    console.log('Error en asignarPedidoAVendedor ' + VERSION + ':', e.message);
    console.log(e.stack);

    resultado = {
      ok: false,
      tipo: 'ERROR',
      mensaje: `❌ <b>Error al tomar pedido ${VERSION}</b>\n\n${html_(e.message)}`
    };

  } finally {
    lock.releaseLock();
  }

  // Editamos Telegram fuera del lock. No se manda mensaje privado aquí.
  // El único que debe mandar el mensaje con botones al vendedor sigue siendo la función principal.
  try {
    if (resultado && resultado.ok) {
      const pedidoTomadoMostrar = String(resultado.idPedido || tokenPedido || 'SIN_ID')
        .replace(/\bnull\b/gi, '')
        .replace(/\bundefined\b/gi, '')
        .replace(/,+\s*$/g, '')
        .trim() || 'SIN_ID';

      const vendedorTomadoMostrar = String(vendedor.nombre || vendedor.id_vendedor || 'Sin nombre')
        .replace(/\bnull\b/gi, '')
        .replace(/\bundefined\b/gi, '')
        .trim() || 'Sin nombre';

      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `✅ <b>Pedido tomado ${VERSION}</b>

` +
        `🧾 Pedido: ${html_(pedidoTomadoMostrar)}
` +
        `👤 Vendedor: ${html_(vendedorTomadoMostrar)}`,
        null
      );

      console.log(
        'Pedido tomado ' + VERSION + ' usando la misma fila. Pedido: ' + resultado.idPedido +
        ' | Fila: ' + resultado.fila +
        ' | Vendedor: ' + vendedor.id_vendedor
      );
    } else {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        resultado && resultado.mensaje
          ? resultado.mensaje
          : `❌ <b>No se pudo tomar pedido ${VERSION}</b>`,
        null
      );
    }
  } catch (e) {
    console.log('No se pudo editar mensaje del grupo después de tomar pedido:', e.message);
  }

  return ok();
}

// ============================================
// NO PUEDO ATENDER / DEVOLVER AL GRUPO
// ============================================

function devolverPedidoAGrupo(cb) {
  const cache = CacheService.getScriptCache();
  const callbackId = String(cb.id || '');

  responderCallbackTelegram(callbackId, '🔄 Reasignando ' + VERSION + '...');

  if (cache.get('CB_' + callbackId)) {
    console.log('Callback reasignar duplicado ignorado:', callbackId);
    return ok();
  }

  cache.put('CB_' + callbackId, 'SI', 300);

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1200)) {
    console.log('Sistema ocupado en reasignar. Callback: ' + callbackId);
    return ok();
  }

  try {
    const ahora = new Date();

    if (!esHorarioPermitidoParaGrupo(ahora)) {
      responderCallbackTelegram(callbackId, '⛔ Fuera de horario para reasignar. El pedido no se mandó al grupo.');
      return ok();
    }

    const tokenPedido = String(cb.data || '').replace('reasignar_', '').replace(/_r\d+$/i, '').trim();
    const pedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);

    if (!pedidos) {
      responderCallbackTelegram(callbackId, '❌ No encontré PEDIDOS');
      return ok();
    }

    const filaPedido = buscarFilaPedidoPorToken(pedidos, tokenPedido);

    if (filaPedido === -1) {
      responderCallbackTelegram(callbackId, '❌ Pedido no encontrado');
      return ok();
    }

    const headers = pedidos.getDataRange().getValues()[0];
    const idx = nombre => headers.indexOf(nombre);

    try {
      editarMensajeTelegram(
        cb.message.chat.id,
        cb.message.message_id,
        `⌛ <b>Pedido rechazado ${VERSION}</b>\n\nDevuelto al grupo.`,
        null
      );
    } catch (e) {
      console.log('No se pudo editar mensaje del vendedor que rechazó:', e.message);
    }

    setCeldaSiExiste_(pedidos, filaPedido, headers, 'estado', 'Sin Vendedor');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'id_vendedor', ID_VENDEDOR_DEFAULT);
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'Vendedor_Nombre', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'chatid', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'Hora_esti_entrega', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'mensaje_telegram_id', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'ejecuciones_webhook', 0);
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'Fechahora aceptado', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'webhook_hora_enviado', '');
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'avisos_retraso_entrega', 0);
    setCeldaSiExiste_(pedidos, filaPedido, headers, 'notificacion_entregada', 'No');

    SpreadsheetApp.flush();

    console.log('Pedido devuelto al grupo ' + VERSION + '. Token: ' + tokenPedido);
    return ok();

  } finally {
    lock.releaseLock();
  }
}

// ============================================
// BOTONES DE HORA
// Aquí está la corrección importante:
// Se escribe Hora_esti_entrega inmediatamente en la hoja.
// Make se conserva para notificar al cliente, pero ya no controla la velocidad.
// ============================================

function notificarHoraAMake(cb) {
  // v4.0:
  // 1) El botón de hora registra la hora localmente.
  // 2) Make SOLO se llama si hay Hora_esti_entrega y webhook_hora_enviado está vacío.
  // 3) Antes de llamar a Make, se marca webhook_hora_enviado para cortar cualquier duplicado.
  // 4) La hora se guarda como fecha real y se manda a Make también como texto WhatsApp:
  //    dd/MM/yyyy hh:mm AM/PM
  const cache = CacheService.getScriptCache();
  const callbackId = String(cb.id || '');
  const callbackData = String(cb.data || '').trim();
  const chatId = cb.message && cb.message.chat ? cb.message.chat.id : '';
  const messageId = cb.message ? cb.message.message_id : '';

  const tokenPedido = extraerTokenDesdeCallback_(callbackData);
  const minutos = obtenerMinutosDesdeCallback_(callbackData);

  function editarActualSinBotones_(texto) {
    try {
      if (chatId && messageId) {
        editarMensajeTelegram(chatId, messageId, texto, null);
      }
    } catch (e) {
      console.log('No se pudo editar mensaje actual sin botones ' + VERSION + ':', e.message);
    }
  }

  function responderSeguro_(texto) {
    try {
      responderCallbackTelegram(callbackId, texto);
    } catch (e) {
      console.log('No se pudo responder callback ' + VERSION + ':', e.message);
    }
  }

  function editarExtraSinBotones_(chatExtra, mensajeExtra, texto) {
    try {
      if (chatExtra && mensajeExtra) {
        editarMensajeTelegram(chatExtra, mensajeExtra, texto, null);
      }
    } catch (e) {
      console.log('No se pudo limpiar mensaje extra ' + VERSION + ':', e.message);
    }
  }

  if (!tokenPedido || !minutos) {
    responderSeguro_('❌ Botón de hora inválido');
    editarActualSinBotones_(
      `❌ <b>Botón de hora inválido ${VERSION}</b>

Intenta usar el mensaje más reciente.`
    );
    return ok();
  }

  // Anti-duplicado por callback exacto.
  if (callbackId && cache.get('CB_' + callbackId)) {
    responderSeguro_('⏳ Ya recibido');
    editarActualSinBotones_(
      `⏳ <b>Botón ya recibido ${VERSION}</b>

Esta selección ya fue recibida. Estoy evitando duplicados.`
    );
    return ok();
  }

  if (callbackId) {
    cache.put('CB_' + callbackId, 'SI', 300);
  }

  // Anti-duplicado por token de pedido.
  if (cache.get('HORA_TOKEN_' + tokenPedido)) {
    responderSeguro_('⏳ Hora ya recibida');
    editarActualSinBotones_(
      `⏳ <b>Hora ya recibida ${VERSION}</b>

Este pedido ya está en proceso de registro.`
    );
    return ok();
  }

  cache.put('HORA_TOKEN_' + tokenPedido, 'SI', 300);
  marcarHoraEnProceso_(tokenPedido);

  // Respuesta visible inmediata, antes de Sheets y Make.
  responderSeguro_('✅ Hora recibida');

  // v4.1: quitamos SOLO los botones. El texto completo del cliente permanece visible.
  // Antes se reemplazaba todo el mensaje por "Registrando hora" y por eso parecían
  // desaparecer los datos del cliente.
  try {
    quitarBotonesTelegram_(chatId, messageId);
  } catch (e) {
    console.log('No se pudieron quitar botones conservando el texto ' + VERSION + ':', e.message);
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    editarActualSinBotones_(
      `⚠️ <b>Sistema ocupado ${VERSION}</b>

El botón ya fue recibido, pero otro proceso está usando la hoja.

Intenta nuevamente en unos segundos si no se registra la hora.`
    );
    return ok();
  }

  try {
    const C = CFG();
    const pedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);

    if (!pedidos) {
      editarActualSinBotones_(`❌ <b>No encontré PEDIDOS ${VERSION}</b>`);
      return ok();
    }

    const headers = obtenerHeadersRapido_(pedidos);
    const idx = nombre => headers.indexOf(nombre);

    const idxHora = idx('Hora_esti_entrega');
    const idxEstado = idx('estado');
    const idxWebhookHora = idx('webhook_hora_enviado');
    const idxEjec = idx('ejecuciones_webhook');
    const idxAvisosRetraso = idx('avisos_retraso_entrega');
    const idxIdPedido = idx('id_pedido');
    const idxMensajeId = idx('mensaje_telegram_id');
    const idxChatId = idx('chatid');

    if (idxHora === -1 || idxEstado === -1) {
      editarActualSinBotones_(
        `❌ <b>Faltan columnas ${VERSION}</b>

Faltan columnas base: estado u Hora_esti_entrega.`
      );
      return ok();
    }

    if (idxWebhookHora === -1) {
      editarActualSinBotones_(
        `❌ <b>Falta columna webhook_hora_enviado ${VERSION}</b>

Por seguridad NO se manda a Make, porque no hay columna para marcar que ya se mandó.`
      );
      return ok();
    }

    // Buscar desde abajo para preferir la fila más reciente y evitar filas viejas.
    let filaPedido = -1;
    const data = pedidos.getDataRange().getValues();
    const idxCallbackToken = idx('callback_token');

    if (idxCallbackToken !== -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        const tokenFila = String(data[i][idxCallbackToken] || '').trim();
        const estadoFila = String(data[i][idxEstado] || '').trim();

        if (
          tokenFila === tokenPedido &&
          estadoFila !== 'Reasignado' &&
          estadoFila.toUpperCase() !== 'REASIGNADO'
        ) {
          filaPedido = i + 1;
          break;
        }
      }
    }

    if (filaPedido === -1 && idxIdPedido !== -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        const idFila = String(data[i][idxIdPedido] || '').trim();
        const estadoFila = String(data[i][idxEstado] || '').trim();

        if (
          idFila === tokenPedido &&
          estadoFila !== 'Reasignado' &&
          estadoFila.toUpperCase() !== 'REASIGNADO'
        ) {
          filaPedido = i + 1;
          break;
        }
      }
    }

    if (filaPedido === -1) {
      editarActualSinBotones_(
        `❌ <b>Pedido no encontrado ${VERSION}</b>

No encontré una fila activa para este botón.

Intenta usar el mensaje más reciente.`
      );
      return ok();
    }

    const lastCol = Math.max(pedidos.getLastColumn(), headers.length);
    const rangoFila = pedidos.getRange(filaPedido, 1, 1, lastCol);
    const filaValores = rangoFila.getValues()[0];

    const idPedido = idxIdPedido !== -1
      ? String(filaValores[idxIdPedido] || tokenPedido).trim()
      : tokenPedido;

    const estadoActual = String(filaValores[idxEstado] || '').trim();
    const horaExistente = filaValores[idxHora];
    const webhookYaEnviado = filaValores[idxWebhookHora];
    const mensajeGuardadoAntes = idxMensajeId !== -1 ? filaValores[idxMensajeId] : '';
    const chatGuardado = idxChatId !== -1 ? filaValores[idxChatId] : chatId;

    if (estadoActual === 'Reasignado' || estadoActual.toUpperCase() === 'REASIGNADO') {
      editarActualSinBotones_(
        `⚠️ <b>Pedido cerrado ${VERSION}</b>

Este pedido ya está como Reasignado y no debe recibir hora.

🧾 Pedido: ${html_(idPedido)}`
      );
      return ok();
    }

    let horaEstimada;

    if (horaExistente) {
      horaEstimada = parsearFechaEntrega(horaExistente);

      if (!horaEstimada) {
        editarActualSinBotones_(
          `❌ <b>Hora existente inválida ${VERSION}</b>

El pedido ya tiene Hora_esti_entrega, pero no pude interpretarla.

🧾 Pedido: ${html_(idPedido)}`
        );
        return ok();
      }
    } else {
      horaEstimada = new Date();
      horaEstimada.setMinutes(horaEstimada.getMinutes() + minutos);

      filaValores[idxHora] = horaEstimada;
      filaValores[idxEstado] = 'En Camino';

      if (idxEjec !== -1) {
        filaValores[idxEjec] = 99;
      }

      if (idxAvisosRetraso !== -1) {
        filaValores[idxAvisosRetraso] = 0;
      }

      if (idxMensajeId !== -1 && messageId) {
        filaValores[idxMensajeId] = messageId;
      }
    }

    const tiempoTexto = textoTiempoDesdeMinutos_(minutos);
    const horaTextoWhatsapp = formatearFechaParaWhatsApp_(horaEstimada);

    // Si Make ya fue avisado, NO se vuelve a consumir crédito.
    let debeAvisarMake = false;

    if (!webhookYaEnviado) {
      // Se marca ANTES de llamar a Make para cortar duplicados aunque entren varios doPost.
      filaValores[idxWebhookHora] = new Date();
      debeAvisarMake = true;
    }

    rangoFila.setValues([filaValores]);

    // Mantener la celda como fecha real, pero visible en formato 12h AM/PM.
    pedidos
      .getRange(filaPedido, idxHora + 1)
      .setNumberFormat('dd/MM/yyyy hh:mm AM/PM');

    pedidos
      .getRange(filaPedido, idxWebhookHora + 1)
      .setNumberFormat('dd/MM/yyyy hh:mm AM/PM');

    SpreadsheetApp.flush();

    // v4.1: el mensaje privado conserva TODOS los datos del cliente.
    const filaConfirmada = pedidos.getRange(filaPedido, 1, 1, headers.length).getValues()[0];
    const payloadConfirmado = {};
    headers.forEach((h, j) => payloadConfirmado[h] = filaConfirmada[j]);

    editarActualSinBotones_(
      construirMensajePedidoAsignadoConHora_(
        payloadConfirmado,
        idPedido,
        tiempoTexto,
        horaTextoWhatsapp
      )
    );

    // Cierra otro mensaje guardado si es distinto del que se acaba de pulsar.
    if (
      mensajeGuardadoAntes &&
      messageId &&
      String(mensajeGuardadoAntes) !== String(messageId)
    ) {
      editarExtraSinBotones_(
        chatGuardado,
        mensajeGuardadoAntes,
        `✅ <b>Hora registrada ${VERSION}</b>

` +
        `🧾 Pedido: ${html_(idPedido)}

` +
        `Este mensaje fue cerrado porque ya se registró una hora.`
      );
    }

    // Make SOLO al final, SOLO si hay Hora_esti_entrega y webhook_hora_enviado estaba vacío.
    if (!debeAvisarMake) {
      console.log(
        'Make NO se envía. Ya estaba marcado webhook_hora_enviado. ' +
        VERSION + ' | Pedido: ' + idPedido + ' | Valor: ' + webhookYaEnviado
      );
      return ok();
    }

    if (!horaEstimada) {
      console.log('Make NO se envía porque no hay hora estimada. Pedido: ' + idPedido + ' ' + VERSION);
      return ok();
    }

    try {
      if (C.WEBHOOK_MAKE) {
        const filaActual = pedidos.getRange(filaPedido, 1, 1, headers.length).getValues()[0];
        const pedidoCompleto = {};
        headers.forEach((h, j) => pedidoCompleto[h] = filaActual[j]);

        const payloadMake = {
          callback_query: cb,
          callback_data: callbackData,
          telegram_user_id: cb.from && cb.from.id ? cb.from.id : '',
          telegram_user_name: cb.from && cb.from.first_name ? cb.from.first_name : '',
          chat_id: chatId,
          message_id: messageId,
          id_pedido: idPedido,
          token_pedido: tokenPedido,
          minutos: minutos,
          tiempo_texto: tiempoTexto,

          // Fecha en varios formatos para evitar que Make/WhatsApp la deformen.
          hora_estimada: horaEstimada,
          hora_estimada_iso: horaEstimada.toISOString(),
          hora_estimada_texto: horaTextoWhatsapp,
          hora_estimada_whatsapp: horaTextoWhatsapp,

          fecha_callback: new Date(),
          actualizado_por_apps_script: true,
          version_apps_script: VERSION,
          origen_apps_script: 'BOTON_HORA_LOCAL_INMEDIATO',
          webhook_hora_enviado: filaValores[idxWebhookHora],
          pedido: pedidoCompleto
        };

        const respuesta = UrlFetchApp.fetch(C.WEBHOOK_MAKE, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payloadMake),
          muteHttpExceptions: true
        });

        console.log(
          'Callback enviado a Make UNA SOLA VEZ después de registrar hora ' + VERSION +
          '. Código: ' + respuesta.getResponseCode() +
          ' | Pedido: ' + idPedido +
          ' | Respuesta: ' + respuesta.getContentText()
        );
      } else {
        console.log('Make NO se envía porque WEBHOOK_MAKE está vacío. Pedido: ' + idPedido + ' ' + VERSION);
      }
    } catch (eMake) {
      // No borramos webhook_hora_enviado. Esto evita quemar créditos por reintentos infinitos.
      console.log('Make falló, pero webhook_hora_enviado quedó marcado para evitar duplicados ' + VERSION + ':', eMake.message);
    }

  } catch (e) {
    console.log('Error en notificarHoraAMake ' + VERSION + ':', e.message);
    console.log(e.stack);

    editarActualSinBotones_(
      `❌ <b>Error al registrar hora ${VERSION}</b>

` +
      `El botón sí fue recibido, pero ocurrió un error al escribir en la hoja.

` +
      `Detalle: ${html_(e.message)}`
    );
  } finally {
    lock.releaseLock();
  }

  return ok();
}

// ============================================
// REVISAR ENTREGAS ATRASADAS
// ============================================

function revisarEntregasAtrasadas() {
  // v3.7: no usa ScriptLock global para no bloquear callbacks de Telegram.
  const cache = CacheService.getScriptCache();
  const runKey = 'RUN_RETRASOS_' + VERSION;

  if (cache.get(runKey)) {
    console.log('Otra ejecución de revisarEntregasAtrasadas sigue activa. Se cancela sin bloquear callbacks. ' + VERSION);
    return;
  }

  cache.put(runKey, 'SI', 55);

  try {
    const C = CFG();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);

    if (!sheet) {
      console.log('No existe hoja PEDIDOS.');
      return;
    }

    const ahora = new Date();
    const horaActual = ahora.getHours();

    if (
      C.RESPETAR_HORARIOS &&
      (horaActual < C.HORARIO_RETRASOS_INICIO || horaActual >= C.HORARIO_RETRASOS_FIN)
    ) {
      console.log('Fuera de horario para revisar entregas atrasadas.');
      return;
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return;
    }

    const headers = data[0];
    const idx = nombre => headers.indexOf(nombre);

    const idxIdPedido = idx('id_pedido');
    const idxTelefono = idx('telefono');
    const idxNombre = idx('nombre');
    const idxProducto = idx('producto');
    const idxCantidad = idx('cantidad');
    const idxIdVendedor = idx('id_vendedor');
    const idxNombreVend = idx('Vendedor_Nombre');
    const idxHoraEst = idx('Hora_esti_entrega');
    const idxHoraReal = idx('hora_entrega_real');
    const idxChatId = idx('chatid');
    const idxAvisos = idx('avisos_retraso_entrega');
    const idxDetalles = idx('Detalles_Pedido');
    const idxMensajeOriginal = idx('Mensaje Original');

    if (idxHoraEst === -1 || idxHoraReal === -1 || idxChatId === -1 || idxAvisos === -1) {
      console.log('Faltan columnas necesarias para revisar entregas atrasadas.');
      return;
    }

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const row = i + 1;

      const horaEst = fila[idxHoraEst];
      const horaReal = fila[idxHoraReal];
      const chatId = fila[idxChatId];

      if (!horaEst || horaReal || !chatId) {
        continue;
      }

      const horaEstimada = parsearFechaEntrega(horaEst);

      if (!horaEstimada) {
        console.log('Hora estimada inválida en fila ' + row);
        continue;
      }

      if (ahora <= horaEstimada) {
        continue;
      }

      const avisosActuales = Number(sheet.getRange(row, idxAvisos + 1).getValue()) || 0;
      const minutosRetraso = Math.floor((ahora - horaEstimada) / 60000);

      const idPedido = idxIdPedido !== -1 ? fila[idxIdPedido] : 'SIN_ID';
      const nombreCliente = idxNombre !== -1 ? fila[idxNombre] : '';
      const telefono = idxTelefono !== -1 ? fila[idxTelefono] : '';
      const producto = idxProducto !== -1 ? fila[idxProducto] : '';
      const cantidad = idxCantidad !== -1 ? fila[idxCantidad] : '';
      const idVendedor = idxIdVendedor !== -1 ? fila[idxIdVendedor] : '';
      const nombreVendedor = idxNombreVend !== -1 ? fila[idxNombreVend] : '';
      const detalles = idxDetalles !== -1 ? fila[idxDetalles] : '';
      const mensajeOriginal = idxMensajeOriginal !== -1 ? fila[idxMensajeOriginal] : '';

      if (avisosActuales < 10) {
        const nuevoAviso = avisosActuales + 1;

        const mensajeVendedor =
          `🚨 <b>ENTREGA ATRASADA ${VERSION}</b>\n\n` +
          `⏱ <b>Aviso ${nuevoAviso} de 10</b>\n\n` +
          `Este pedido ya pasó de la hora estimada y todavía no aparece como entregado en AppSheet.\n\n` +
          `👤 Cliente: ${nombreCliente || 'Sin nombre'}\n` +
          `📦 Producto: ${producto || 'Sin producto'} x${cantidad || ''}\n` +
          `${formatearTelefonoConSeparadores(telefono)}\n` +
          `🧾 Pedido: ${idPedido || 'SIN_ID'}\n\n` +
          `🕒 Hora estimada: ${formatearFechaEntrega_(horaEstimada)}\n` +
          `⏰ Retraso aproximado: ${minutosRetraso} minutos\n\n` +
          `⚠ Si ya entregaste, entra a AppSheet y registra la entrega real.\n` +
          `⚠ Si no has entregado, atiende este pedido cuanto antes.`;

        const rRetraso = enviarMensajeTelegramRetrasos(chatId, mensajeVendedor, null);

        // Solo contamos el aviso si el bot de retrasos confirmó el envío.
        // Esto evita quemar los 10 avisos si el vendedor todavía no inició el bot nuevo.
        if (rRetraso && rRetraso.ok) {
          sheet.getRange(row, idxAvisos + 1).setValue(nuevoAviso);
          SpreadsheetApp.flush();
        } else {
          console.log(
            'No se contó aviso de retraso porque TELEGRAM_TOKEN_RETRASOS no confirmó el envío. ' +
            'Pedido: ' + (idPedido || 'SIN_ID') +
            ' | Chat: ' + chatId +
            ' | Respuesta: ' + JSON.stringify(rRetraso)
          );
        }

        continue;
      }

      if (avisosActuales === 10) {
        const mensajeAdmin =
          `🚨 <b>VENDEDOR NO COMPLETÓ ENTREGA ${VERSION}</b>\n\n` +
          `Se le avisó <b>10 veces</b> al vendedor y el pedido sigue sin hora de entrega real en AppSheet.\n\n` +
          `🛵 Vendedor: ${nombreVendedor || idVendedor || 'Sin vendedor'}\n` +
          `👤 Cliente: ${nombreCliente || 'Sin nombre'}\n` +
          `📦 Producto: ${producto || 'Sin producto'} x${cantidad || ''}\n` +
          `${formatearTelefonoConSeparadores(telefono)}\n` +
          `🧾 Pedido: ${idPedido || 'SIN_ID'}\n\n` +
          `🕒 Hora estimada: ${formatearFechaEntrega_(horaEstimada)}\n` +
          `⏰ Retraso aproximado: ${minutosRetraso} minutos\n\n` +
          `💬 Mensaje Original: ${mensajeOriginal || 'Sin mensaje'}\n` +
          `📝 Detalles: ${detalles || 'Sin detalles'}`;

        enviarMensajeTelegram(C.GRUPO_MANDO_ID, mensajeAdmin, null);
        sheet.getRange(row, idxAvisos + 1).setValue(11);
        SpreadsheetApp.flush();
        continue;
      }
    }

  } finally {
    cache.remove(runKey);
  }
}

// ============================================
// TELEGRAM
// ============================================

function enviarMensajeTelegram(chatId, mensaje, botones) {
  const C = CFG();

  if (!C.TELEGRAM_TOKEN) {
    throw new Error('Falta TELEGRAM_TOKEN en Propiedades del script.');
  }

  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/sendMessage`;

  const payload = {
    chat_id: String(chatId).trim(),
    text: asegurarVersionTelegram_(mensaje, 'SEND_MESSAGE'),
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  if (botones) {
    payload.reply_markup = botones;
  }

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const texto = response.getContentText();
  console.log('Telegram sendMessage ' + VERSION + ':', texto);

  try {
    return JSON.parse(texto);
  } catch (e) {
    return { ok: false, raw: texto };
  }
}

// Envío exclusivo para alertas de entregas retrasadas.
// Usa un bot separado para permitir un ringtone distinto en Telegram.
function enviarMensajeTelegramRetrasos(chatId, mensaje, botones) {
  const C = CFG();

  if (!C.TELEGRAM_TOKEN_RETRASOS) {
    console.log('Falta TELEGRAM_TOKEN_RETRASOS en Propiedades del script. ' + VERSION);
    return { ok: false, error: 'FALTA_TELEGRAM_TOKEN_RETRASOS' };
  }

  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN_RETRASOS}/sendMessage`;

  const payload = {
    chat_id: String(chatId).trim(),
    text: asegurarVersionTelegram_(mensaje, 'SEND_MESSAGE_RETRASOS'),
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  if (botones) {
    payload.reply_markup = botones;
  }

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const texto = response.getContentText();
  console.log('Telegram RETRASOS sendMessage ' + VERSION + ':', texto);

  try {
    return JSON.parse(texto);
  } catch (e) {
    return { ok: false, raw: texto };
  }
}

function editarMensajeTelegram(chatId, messageId, texto, botones) {
  const C = CFG();

  if (!chatId || !messageId) {
    console.log('No se edita mensaje: falta chatId o messageId. ' + VERSION);
    return { ok: false };
  }

  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/editMessageText`;

  const payload = {
    chat_id: chatId,
    message_id: messageId,
    text: asegurarVersionTelegram_(texto, 'EDIT_MESSAGE'),
    parse_mode: 'HTML',
    disable_web_page_preview: true,
    reply_markup: botones || { inline_keyboard: [] }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const respuesta = response.getContentText();
  console.log('Telegram editMessageText ' + VERSION + ':', respuesta);

  try {
    return JSON.parse(respuesta);
  } catch (e) {
    return { ok: false, raw: respuesta };
  }
}

function quitarBotonesTelegram_(chatId, messageId) {
  const C = CFG();

  if (!chatId || !messageId) {
    return { ok: false };
  }

  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/editMessageReplyMarkup`;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: chatId,
      message_id: messageId,
      reply_markup: { inline_keyboard: [] }
    }),
    muteHttpExceptions: true
  });

  const respuesta = response.getContentText();
  console.log('Telegram editMessageReplyMarkup ' + VERSION + ':', respuesta);

  try {
    return JSON.parse(respuesta);
  } catch (e) {
    return { ok: false, raw: respuesta };
  }
}

function responderCallbackTelegram(callbackId, texto) {
  const C = CFG();

  if (!callbackId) {
    return;
  }

  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/answerCallbackQuery`;

  const textoFinal = `${VERSION} | ${String(texto || '').trim()}`;

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      callback_query_id: callbackId,
      text: textoFinal.substring(0, 190),
      show_alert: true
    }),
    muteHttpExceptions: true
  });
}

function asegurarVersionTelegram_(texto, origen) {
  const t = String(texto || '');

  if (t.indexOf(VERSION) !== -1) {
    return t;
  }

  return `<b>SuperAgua ${VERSION}</b>

` + t;
}

function html_(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


function ok() {
  return ContentService.createTextOutput('ok');
}

// ============================================
// MENSAJES Y BOTONES
// ============================================

function construirMensajePedidoAsignado_(payload, idPedido, intento, origen, tokenBoton) {
  const ubicacion = formatearUbicacion(payload.ubicacion);

  return (
    `📋 <b>PEDIDO ASIGNADO ${VERSION}</b>\n\n` +
    `⏱ <b>Intento ${intento} de 10</b>\n\n` +
    `👤 Cliente: ${html_(payload.nombre || 'Sin nombre')}\n` +
    `📦 Producto: ${html_(payload.producto || 'Sin producto')} x${html_(payload.cantidad || '')}\n` +
    `📍 Ubicación: ${ubicacion}\n` +
    `${formatearTelefonoConSeparadores(payload.telefono)}\n` +
    `🧾 Pedido: ${html_(idPedido)}\n` +
    `💬 Mensaje Original: ${html_(payload['Mensaje Original'] || 'Sin mensaje')}\n` +
    `📝 Detalles: ${html_(payload['Detalles_Pedido'] || 'Sin detalles')}\n\n` +
    `⚠ SELECCIONA EL TIEMPO ESTIMADO`
  );
}

function construirMensajePedidoAsignadoConHora_(payload, idPedido, tiempoTexto, horaTextoWhatsapp) {
  const ubicacion = formatearUbicacion(payload.ubicacion);

  return (
    `📋 <b>PEDIDO ASIGNADO ${VERSION}</b>\n\n` +
    `👤 Cliente: ${html_(payload.nombre || 'Sin nombre')}\n` +
    `📦 Producto: ${html_(payload.producto || 'Sin producto')} x${html_(payload.cantidad || '')}\n` +
    `📍 Ubicación: ${ubicacion}\n` +
    `${formatearTelefonoConSeparadores(payload.telefono)}\n` +
    `🧾 Pedido: ${html_(idPedido)}\n` +
    `💬 Mensaje Original: ${html_(payload['Mensaje Original'] || 'Sin mensaje')}\n` +
    `📝 Detalles: ${html_(payload['Detalles_Pedido'] || 'Sin detalles')}\n\n` +
    `✅ <b>HORA REGISTRADA</b>\n` +
    `⏱ Tiempo seleccionado: ${html_(tiempoTexto)}\n` +
    `🕒 Hora estimada: ${html_(horaTextoWhatsapp)}`
  );
}

function construirBotonesTiempo_(tokenPedido, row) {
  // v3.7: REGRESO TOTAL AL FORMATO QUE YA FUNCIONABA EN EL PEDIDO ORIGINAL.
  // Nada de _rFILA, nada de metadatos extra en callback_data.
  // El botón visible y el callback interno quedan como el flujo principal estable:
  // 30_TOKEN, 60_TOKEN, 90_TOKEN, 120_TOKEN, 180_TOKEN.
  // El script buscará la fila por callback_token, empezando desde abajo para encontrar la fila nueva.
  const token = String(tokenPedido || '').trim();

  return {
    inline_keyboard: [
      [
        { text: '30 min', callback_data: '30_' + token },
        { text: '1 hora', callback_data: '60_' + token },
        { text: '1.5 horas', callback_data: '90_' + token }
      ],
      [
        { text: '2 horas', callback_data: '120_' + token },
        { text: '3 horas', callback_data: '180_' + token }
      ],
      [
        { text: 'No puedo atender', callback_data: 'reasignar_' + token }
      ]
    ]
  };
}


// ============================================
// HORARIOS
// ============================================

function esHorarioPermitidoParaVendedor(fecha) {
  const C = CFG();

  if (!C.RESPETAR_HORARIOS) {
    return true;
  }

  const hora = fecha.getHours();
  return hora >= C.HORARIO_VENDEDOR_INICIO && hora < C.HORARIO_VENDEDOR_FIN;
}

function esHorarioPermitidoParaGrupo(fecha) {
  const C = CFG();

  if (!C.RESPETAR_HORARIOS) {
    return true;
  }

  const dia = fecha.getDay();
  const hora = fecha.getHours();

  if (dia === 0 && C.BLOQUEAR_DOMINGO) {
    return false;
  }

  if (dia === 6 && hora >= C.BLOQUEAR_SABADO_DESDE) {
    return false;
  }

  if (dia === 1 && hora < C.BLOQUEAR_LUNES_ANTES) {
    return false;
  }

  return hora >= C.HORARIO_GRUPO_INICIO && hora < C.HORARIO_GRUPO_FIN;
}

function esPedidoDeNocheAnterior(fechaPrimerEnvio, ahora) {
  if (!fechaPrimerEnvio) {
    return false;
  }

  const fechaEnvio = new Date(fechaPrimerEnvio);

  if (isNaN(fechaEnvio.getTime())) {
    return false;
  }

  const inicioHoy = new Date(
    ahora.getFullYear(),
    ahora.getMonth(),
    ahora.getDate(),
    7,
    0,
    0
  );

  return fechaEnvio < inicioHoy;
}

// ============================================
// BÚSQUEDAS EN HOJAS
// ============================================

function buscarVendedorPorTelegram(sheet, telegramUserId) {
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return null;
  }

  const headers = data[0];
  const idx = nombre => headers.indexOf(nombre);

  const idxChatId = idx('CHATID');
  const idxNombre = idx('nombre');
  const idxIdVend = idx('id_vendedor');
  const idxActivo = idx('activo');

  if (idxChatId === -1 || idxNombre === -1 || idxIdVend === -1 || idxActivo === -1) {
    console.log('Faltan columnas en PARAMETROS.');
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    const chatIdHoja = String(data[i][idxChatId] || '').trim();
    const activo = data[i][idxActivo];

    const esActivo =
      activo === true ||
      String(activo).trim().toUpperCase() === 'TRUE' ||
      String(activo).trim().toUpperCase() === 'SI' ||
      String(activo).trim().toUpperCase() === 'SÍ' ||
      String(activo).trim() === '1';

    if (chatIdHoja === String(telegramUserId).trim() && esActivo) {
      return {
        nombre: data[i][idxNombre],
        id_vendedor: data[i][idxIdVend],
        chatid: data[i][idxChatId]
      };
    }
  }

  return null;
}

function buscarFilaPedidoPorToken(sheet, token) {
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return -1;
  }

  const headers = data[0];
  const idxToken = headers.indexOf('callback_token');
  const idxId = headers.indexOf('id_pedido');
  const t = String(token || '').trim();

  if (!t) {
    return -1;
  }

  if (idxToken !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxToken] || '').trim() === t) {
        return i + 1;
      }
    }
  }

  if (idxId !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxId] || '').trim() === t) {
        return i + 1;
      }
    }
  }

  return -1;
}

function obtenerHeadersRapido_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function buscarFilaPedidoPorTokenRapido_(sheet, headers, token) {
  const t = String(token || '').trim();
  if (!t) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const idxToken = headers.indexOf('callback_token');
  const idxId = headers.indexOf('id_pedido');

  if (idxToken !== -1) {
    const valores = sheet.getRange(2, idxToken + 1, lastRow - 1, 1).getValues();
    for (let i = valores.length - 1; i >= 0; i--) {
      if (String(valores[i][0] || '').trim() === t) {
        return i + 2;
      }
    }
  }

  if (idxId !== -1) {
    const valoresId = sheet.getRange(2, idxId + 1, lastRow - 1, 1).getValues();
    for (let i = valoresId.length - 1; i >= 0; i--) {
      if (String(valoresId[i][0] || '').trim() === t) {
        return i + 2;
      }
    }
  }

  return -1;
}

function filaCoincideConToken_(sheet, headers, row, token) {
  const t = String(token || '').trim();
  if (!t || row < 2 || row > sheet.getLastRow()) return false;

  const idxToken = headers.indexOf('callback_token');
  const idxId = headers.indexOf('id_pedido');

  if (idxToken !== -1) {
    const valorToken = String(sheet.getRange(row, idxToken + 1).getValue() || '').trim();
    if (valorToken && valorToken === t) return true;
  }

  if (idxId !== -1) {
    const valorId = String(sheet.getRange(row, idxId + 1).getValue() || '').trim();
    if (valorId && valorId === t) return true;
  }

  return false;
}


function obtenerIdPedidoDesdeFila_(sheet, row, headers) {
  const idxId = headers.indexOf('id_pedido');

  if (idxId === -1) {
    return '';
  }

  return sheet.getRange(row, idxId + 1).getValue();
}


function obtenerOCrearCallbackToken_(sheet, headers, row, fallbackId) {
  const idxToken = headers.indexOf('callback_token');

  if (idxToken === -1) {
    return String(fallbackId || '').trim();
  }

  let token = String(sheet.getRange(row, idxToken + 1).getValue() || '').trim();

  if (!token) {
    token = generarCallbackToken_();
    sheet.getRange(row, idxToken + 1).setValue(token);
    SpreadsheetApp.flush();
    console.log('Callback token generado ' + VERSION + ': ' + token + ' | Fila: ' + row);
  }

  return token;
}


function asignarMismaFilaAVendedor_(sheet, headers, filaPedido, vendedor, tokenPedido) {
  const lastCol = sheet.getLastColumn();
  const ahora = new Date();
  const valores = sheet.getRange(filaPedido, 1, 1, lastCol).getValues()[0];

  const idPedido = obtenerValorArray_(valores, headers, 'id_pedido') || tokenPedido || 'SIN_ID';

  // v4.0: se limpia la MISMA fila. No appendRow, no id nuevo, no token nuevo.
  // La función principal verá: estado=Asignado + chatid lleno + Hora vacía.
  setValorArraySiExiste_(valores, headers, 'estado', 'Asignado');
  setValorArraySiExiste_(valores, headers, 'id_vendedor', vendedor.id_vendedor);
  setValorArraySiExiste_(valores, headers, 'Vendedor_Nombre', vendedor.nombre);
  setValorArraySiExiste_(valores, headers, 'chatid', vendedor.chatid);
  setValorArraySiExiste_(valores, headers, 'Fechahora aceptado', ahora);

  // Se eliminan restos del ciclo de grupo/vendedor anterior.
  setValorArraySiExiste_(valores, headers, 'Hora_esti_entrega', '');
  setValorArraySiExiste_(valores, headers, 'hora_entrega_real', '');
  setValorArraySiExiste_(valores, headers, 'mensaje_telegram_id', '');
  setValorArraySiExiste_(valores, headers, 'ejecuciones_webhook', 0);
  setValorArraySiExiste_(valores, headers, 'fecha_primer_envio', '');
  setValorArraySiExiste_(valores, headers, 'webhook_hora_enviado', '');
  setValorArraySiExiste_(valores, headers, 'notificacion_entregada', 'No');
  setValorArraySiExiste_(valores, headers, 'avisos_retraso_entrega', 0);
  setValorArraySiExiste_(valores, headers, 'Procedencia', 'Tomado desde grupo sin crear fila nueva | TOKEN ' + tokenPedido);

  // Si no existe callback_token, se genera. Si ya existe, se conserva.
  // No cambia el formato de botones: la función principal decide eso.
  const idxToken = headers.indexOf('callback_token');
  if (idxToken !== -1 && !String(valores[idxToken] || '').trim()) {
    valores[idxToken] = generarCallbackToken_();
  }

  sheet.getRange(filaPedido, 1, 1, lastCol).setValues([valores]);

  marcarPedidoEnProcesoDeToma_(idPedido);
  marcarPedidoEnProcesoDeToma_(tokenPedido);

  return {
    ok: true,
    idPedido: idPedido,
    token: idxToken !== -1 ? valores[idxToken] : tokenPedido,
    fila: filaPedido
  };
}

function cerrarPedidoViejoYCrearNuevoAsignado_(sheet, headers, filaVieja, vendedor, tokenPedido) {
  const lastCol = sheet.getLastColumn();
  const ahora = new Date();

  const valoresViejos = sheet.getRange(filaVieja, 1, 1, lastCol).getValues()[0];
  const idViejo = obtenerValorArray_(valoresViejos, headers, 'id_pedido') || tokenPedido || 'SIN_ID';
  const mensajeGrupo = obtenerValorArray_(valoresViejos, headers, 'mensaje_telegram_id');

  const idNuevo = generarIdPedidoNuevo_(sheet, headers);
  const tokenNuevo = generarCallbackTokenUnico_(sheet, headers);
  const filaNueva = sheet.getLastRow() + 1;

  const filaViejaNueva = valoresViejos.slice();
  const filaNuevaValores = valoresViejos.slice();

  // Fila vieja: queda cerrada y ya no vuelve a participar en el ciclo.
  setValorArraySiExiste_(filaViejaNueva, headers, 'estado', 'Reasignado');
  setValorArraySiExiste_(filaViejaNueva, headers, 'id_vendedor', ID_VENDEDOR_DEFAULT);
  setValorArraySiExiste_(filaViejaNueva, headers, 'Vendedor_Nombre', '');
  setValorArraySiExiste_(filaViejaNueva, headers, 'chatid', '');
  setValorArraySiExiste_(filaViejaNueva, headers, 'Hora_esti_entrega', '');
  setValorArraySiExiste_(filaViejaNueva, headers, 'mensaje_telegram_id', mensajeGrupo || '');
  setValorArraySiExiste_(filaViejaNueva, headers, 'ejecuciones_webhook', 99);
  setValorArraySiExiste_(filaViejaNueva, headers, 'webhook_hora_enviado', '');
  setValorArraySiExiste_(filaViejaNueva, headers, 'notificacion_entregada', 'No');
  setValorArraySiExiste_(filaViejaNueva, headers, 'avisos_retraso_entrega', 0);
  setValorArraySiExiste_(filaViejaNueva, headers, 'Procedencia', 'Reasignado a nuevo pedido ' + idNuevo + ' | ORIGINAL ' + idViejo);

  // Fila nueva: pedido limpio, igual que si hubiera llegado asignado desde el inicio.
  setValorArraySiExiste_(filaNuevaValores, headers, 'id_pedido', idNuevo);
  setValorArraySiExiste_(filaNuevaValores, headers, 'callback_token', tokenNuevo);
  setValorArraySiExiste_(filaNuevaValores, headers, 'estado', 'Asignado');
  setValorArraySiExiste_(filaNuevaValores, headers, 'id_vendedor', vendedor.id_vendedor);
  setValorArraySiExiste_(filaNuevaValores, headers, 'Vendedor_Nombre', vendedor.nombre);
  setValorArraySiExiste_(filaNuevaValores, headers, 'chatid', vendedor.chatid);
  setValorArraySiExiste_(filaNuevaValores, headers, 'Fechahora aceptado', ahora);
  setValorArraySiExiste_(filaNuevaValores, headers, 'Hora_esti_entrega', '');
  setValorArraySiExiste_(filaNuevaValores, headers, 'hora_entrega_real', '');
  setValorArraySiExiste_(filaNuevaValores, headers, 'mensaje_telegram_id', '');
  setValorArraySiExiste_(filaNuevaValores, headers, 'ejecuciones_webhook', 0);
  setValorArraySiExiste_(filaNuevaValores, headers, 'fecha_primer_envio', '');
  setValorArraySiExiste_(filaNuevaValores, headers, 'webhook_hora_enviado', '');
  setValorArraySiExiste_(filaNuevaValores, headers, 'notificacion_entregada', 'No');
  setValorArraySiExiste_(filaNuevaValores, headers, 'avisos_retraso_entrega', 0);
  setValorArraySiExiste_(filaNuevaValores, headers, 'Procedencia', 'Reasignado desde grupo | ORIGINAL ' + idViejo + ' | TOKEN_ANTERIOR ' + tokenPedido);

  sheet.getRange(filaVieja, 1, 1, lastCol).setValues([filaViejaNueva]);
  sheet.appendRow(filaNuevaValores);

  marcarPedidoEnProcesoDeToma_(idViejo);
  marcarPedidoEnProcesoDeToma_(tokenPedido);
  marcarPedidoEnProcesoDeToma_(idNuevo);
  marcarPedidoEnProcesoDeToma_(tokenNuevo);

  return {
    ok: true,
    idViejo: idViejo,
    idNuevo: idNuevo,
    tokenNuevo: tokenNuevo,
    filaVieja: filaVieja,
    filaNueva: filaNueva
  };
}

function obtenerValorArray_(filaValores, headers, nombreColumna) {
  const idx = headers.indexOf(nombreColumna);
  if (idx === -1) return '';
  return filaValores[idx];
}

function setValorArraySiExiste_(filaValores, headers, nombreColumna, valor) {
  const idx = headers.indexOf(nombreColumna);
  if (idx !== -1) {
    filaValores[idx] = valor;
  }
}

function generarIdPedidoNuevo_(sheet, headers) {
  for (let i = 0; i < 20; i++) {
    const id = Utilities.getUuid().replace(/-/g, '').substring(0, 8);
    if (buscarFilaPedidoPorTokenRapido_(sheet, headers, id) === -1) {
      return id;
    }
  }
  return Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function generarCallbackTokenUnico_(sheet, headers) {
  for (let i = 0; i < 20; i++) {
    const token = generarCallbackToken_();
    if (buscarFilaPedidoPorTokenRapido_(sheet, headers, token) === -1) {
      return token;
    }
  }
  return generarCallbackToken_();
}

function setCeldaSiExiste_(sheet, row, headers, nombreColumna, valor) {
  const idx = headers.indexOf(nombreColumna);

  if (idx !== -1) {
    sheet.getRange(row, idx + 1).setValue(valor);
  }
}

// ============================================
// CACHE ANTI-CARRERA PARA TOMAR PEDIDOS
// ============================================

function marcarPedidoEnProcesoDeToma_(tokenPedido) {
  if (!tokenPedido) {
    return;
  }

  CacheService.getScriptCache().put('TOMANDO_' + String(tokenPedido).trim(), 'SI', 120);
}

function estaPedidoEnProcesoDeToma_(tokenPedido) {
  if (!tokenPedido) {
    return false;
  }

  return CacheService.getScriptCache().get('TOMANDO_' + String(tokenPedido).trim()) === 'SI';
}

function marcarHoraEnProceso_(tokenPedido) {
  if (!tokenPedido) {
    return;
  }

  CacheService.getScriptCache().put('HORA_' + String(tokenPedido).trim(), 'SI', 180);
}

function estaHoraEnProceso_(tokenPedido) {
  if (!tokenPedido) {
    return false;
  }

  return CacheService.getScriptCache().get('HORA_' + String(tokenPedido).trim()) === 'SI';
}


// ============================================
// FORMATO Y FECHAS
// ============================================

function formatearTelefonoConSeparadores(telefono) {
  const limpio = String(telefono || '').replace(/\D/g, '');

  if (!limpio) {
    return '━━━━━━━━━━━━\n📞 Sin teléfono\n━━━━━━━━━━━━';
  }

  return '━━━━━━━━━━━━\n📞 +52' + limpio + '\n━━━━━━━━━━━━';
}

function formatearUbicacion(ubicacion) {
  const coords = ubicacion ? String(ubicacion).trim() : '';

  if (!coords) {
    return 'Sin ubicación';
  }

  const link = 'https://maps.google.com/?q=' + encodeURIComponent(coords);
  return `<a href="${link}">Maps</a> - ${html_(coords)}`;
}

function parsearFechaEntrega(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (!valor) {
    return null;
  }

  const texto = String(valor).trim().replace(/\s+/g, ' ');

  const match = texto.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)?$/
  );

  if (!match) {
    console.log('Fecha de entrega no reconocida: ' + texto);
    return null;
  }

  const dia = Number(match[1]);
  const mes = Number(match[2]) - 1;
  const anio = Number(match[3]);
  let hora = Number(match[4]);
  const minuto = Number(match[5]);
  const segundo = match[6] ? Number(match[6]) : 0;
  const ampm = match[7] ? String(match[7]).toUpperCase() : '';

  if (ampm === 'PM' && hora < 12) {
    hora += 12;
  }

  if (ampm === 'AM' && hora === 12) {
    hora = 0;
  }

  const fecha = new Date(anio, mes, dia, hora, minuto, segundo);

  if (isNaN(fecha.getTime())) {
    console.log('Fecha inválida después de convertir: ' + texto);
    return null;
  }

  return fecha;
}

function formatearFechaEntrega_(fecha) {
  return formatearFechaParaWhatsApp_(fecha);
}

function formatearFechaParaWhatsApp_(fecha) {
  const fechaReal = parsearFechaEntrega(fecha);

  if (!fechaReal || isNaN(fechaReal.getTime())) {
    return '';
  }

  const dia = String(fechaReal.getDate()).padStart(2, '0');
  const mes = String(fechaReal.getMonth() + 1).padStart(2, '0');
  const anio = fechaReal.getFullYear();

  let hora = fechaReal.getHours();
  const minuto = String(fechaReal.getMinutes()).padStart(2, '0');
  const ampm = hora >= 12 ? 'PM' : 'AM';

  hora = hora % 12;
  if (hora === 0) hora = 12;

  const horaTexto = String(hora).padStart(2, '0');

  return `${dia}/${mes}/${anio} ${horaTexto}:${minuto} ${ampm}`;
}

function obtenerMinutosDesdeCallback_(callbackData) {
  const data = String(callbackData || '');

  if (data.startsWith('30_') || data.startsWith('0.5h_')) return 30;
  if (data.startsWith('60_') || data.startsWith('1h_')) return 60;
  if (data.startsWith('90_') || data.startsWith('1.5h_')) return 90;
  if (data.startsWith('120_') || data.startsWith('2h_')) return 120;
  if (data.startsWith('180_') || data.startsWith('3h_')) return 180;

  return 0;
}

function textoTiempoDesdeMinutos_(minutos) {
  if (minutos === 30) return '30 minutos';
  if (minutos === 60) return '1 hora';
  if (minutos === 90) return '1.5 horas';
  if (minutos === 120) return '2 horas';
  if (minutos === 180) return '3 horas';
  return minutos + ' minutos';
}

function extraerTokenDesdeCallback_(callbackData) {
  const data = String(callbackData || '');
  const pos = data.indexOf('_');

  if (pos === -1) {
    return '';
  }

  let token = data.substring(pos + 1).trim();
  token = token.replace(/_r\d+$/i, '').trim();
  return token;
}

function extraerFilaDesdeCallback_(callbackData) {
  const data = String(callbackData || '');
  const m = data.match(/_r(\d+)$/i);
  return m ? Number(m[1]) : 0;
}

function generarCallbackToken_() {
  return Utilities.getUuid()
    .replace(/-/g, '')
    .substring(0, 12);
}

// ============================================
// WEBHOOK Y TRIGGERS
// ============================================

function activarWebhook() {
  const C = CFG();
  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/setWebhook?url=${encodeURIComponent(C.URL_WEBAPP)}`;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  console.log(r.getContentText());
}

function revisarWebhook() {
  const C = CFG();
  const url = `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/getWebhookInfo`;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  console.log(r.getContentText());
}

function limpiarYReactivar() {
  const C = CFG();

  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/deleteWebhook?drop_pending_updates=true`,
    { muteHttpExceptions: true }
  );

  const r = UrlFetchApp.fetch(
    `https://api.telegram.org/bot${C.TELEGRAM_TOKEN}/setWebhook?url=${encodeURIComponent(C.URL_WEBAPP)}`,
    { muteHttpExceptions: true }
  );

  console.log('setWebhook:', r.getContentText());
  revisarWebhook();
}

function revisarTriggersInstalados() {
  const triggers = ScriptApp.getProjectTriggers();

  console.log('=== TRIGGERS INSTALADOS ===');

  triggers.forEach(t => {
    console.log(
      'ID: ' + t.getUniqueId() +
      ' | Función: ' + t.getHandlerFunction() +
      ' | Evento: ' + t.getEventType()
    );
  });

  console.log('Total triggers: ' + triggers.length);
}

function dejarSoloUnTriggerMonitoreosCadaMinuto() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    const funcion = t.getHandlerFunction();

    if (
      funcion === 'procesarPedidosPendientes' ||
      funcion === 'revisarEntregasAtrasadas' ||
      funcion === 'revisarEntregasRetrasadas' ||
      funcion === 'enviarPendientesAMake' ||
      funcion === 'ejecutarMonitoreosCadaMinuto'
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp
    .newTrigger('ejecutarMonitoreosCadaMinuto')
    .timeBased()
    .everyMinutes(1)
    .create();

  console.log('Listo: quedó un solo trigger cada minuto para ejecutarMonitoreosCadaMinuto.');
}

// ============================================
// TESTS ÚTILES
// ============================================

function testConfig() {
  const C = CFG();
  console.log(JSON.stringify({
    tieneTelegramToken: !!C.TELEGRAM_TOKEN,
    tieneGrupoPedidos: !!C.GRUPO_PEDIDOS_ID,
    tieneGrupoMando: !!C.GRUPO_MANDO_ID,
    tieneWebhookMake: !!C.WEBHOOK_MAKE,
    tieneUrlWebapp: !!C.URL_WEBAPP,
    respetarHorarios: C.RESPETAR_HORARIOS,
    horarioVendedor: C.HORARIO_VENDEDOR_INICIO + '-' + C.HORARIO_VENDEDOR_FIN,
    horarioGrupo: C.HORARIO_GRUPO_INICIO + '-' + C.HORARIO_GRUPO_FIN,
    version: VERSION
  }, null, 2));
}
