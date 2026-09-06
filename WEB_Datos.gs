// ============================================================
// WEB_Datos.gs
// CAPA DE LECTURA - REPARTO 2026
// ============================================================
// SOLO LECTURA.
// Este archivo NO debe:
// - setValue()
// - setValues()
// - appendRow()
// - deleteRow()
// - clear()
// ============================================================


/**
 * Devuelve un diagnóstico completo del archivo.
 *
 * Incluye:
 * - nombre del spreadsheet
 * - número de hojas
 * - filas utilizadas
 * - columnas utilizadas
 * - encabezados
 * - clasificación tentativa
 */
function WEB_getDiagnosticoBase() {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('No fue posible identificar el Spreadsheet activo.');
    }

    const hojas = ss.getSheets();

    const resultado = {
      ok: true,
      archivo: ss.getName(),
      spreadsheetId: ss.getId(),
      totalHojas: hojas.length,
      totalRegistros: 0,
      hojas: []
    };


    hojas.forEach(sheet => {

      const nombre = sheet.getName();

      const ultimaFila = sheet.getLastRow();
      const ultimaColumna = sheet.getLastColumn();

      let encabezados = [];

      if (ultimaFila > 0 && ultimaColumna > 0) {

        encabezados = sheet
          .getRange(1, 1, 1, ultimaColumna)
          .getDisplayValues()[0]
          .map(valor => String(valor || '').trim());

      }


      const registros =
        ultimaFila > 1
          ? ultimaFila - 1
          : 0;


      resultado.totalRegistros += registros;


      resultado.hojas.push({

        nombre: nombre,

        filas: ultimaFila,

        registros: registros,

        columnas: ultimaColumna,

        encabezados: encabezados,

        tipo: WEB_clasificarHoja_(nombre, encabezados)

      });

    });


    // Ordenar primero hojas con más registros.
    resultado.hojas.sort((a, b) => b.registros - a.registros);


    return resultado;

  } catch (error) {

    console.error(error);

    return {
      ok: false,
      mensaje: error.message,
      stack: error.stack
    };

  }

}


/**
 * Intenta identificar para qué sirve una hoja.
 * Esto es solamente diagnóstico.
 */
function WEB_clasificarHoja_(nombre, encabezados) {

  const n = String(nombre || '').toUpperCase();

  const headers = encabezados
    .join(' ')
    .toUpperCase();


  // ------------------------------------------------------------
  // CLIENTES
  // ------------------------------------------------------------

  if (
    n.includes('CLIENT') ||
    (
      headers.includes('CLIENTE') &&
      (
        headers.includes('TELEFONO') ||
        headers.includes('TELÉFONO')
      )
    )
  ) {
    return 'Clientes';
  }


  // ------------------------------------------------------------
  // PEDIDOS
  // ------------------------------------------------------------

  if (
    n.includes('PEDIDO') ||
    (
      headers.includes('ID_PEDIDO') &&
      headers.includes('ESTADO')
    )
  ) {
    return 'Pedidos';
  }


  // ------------------------------------------------------------
  // VENTAS
  // ------------------------------------------------------------

  if (
    n.includes('VENTA') ||
    (
      headers.includes('IMPORTE') &&
      (
        headers.includes('CANTIDAD') ||
        headers.includes('PRECIO')
      )
    )
  ) {
    return 'Ventas';
  }


  // ------------------------------------------------------------
  // TRANSACCIONES
  // ------------------------------------------------------------

  if (n.includes('TRANSAC')) {
    return 'Transacciones';
  }


  // ------------------------------------------------------------
  // WEBHOOK / BOT
  // ------------------------------------------------------------

  if (
    n.includes('WEBHOOK') ||
    n.includes('WEBHOOK') ||
    headers.includes('SESSION_ID')
  ) {
    return 'Automatización';
  }


  // ------------------------------------------------------------
  // RECORRIDOS / RUTAS
  // ------------------------------------------------------------

  if (
    n.includes('RECORRIDO') ||
    n.includes('RUTA')
  ) {
    return 'Rutas';
  }


  // ------------------------------------------------------------
  // LIQUIDACIONES
  // ------------------------------------------------------------

  if (n.includes('LIQUID')) {
    return 'Liquidaciones';
  }


  // ------------------------------------------------------------
  // PROMOCIONES
  // ------------------------------------------------------------

  if (
    n.includes('PROMO') ||
    headers.includes('CODIGO_ENVIADO') ||
    headers.includes('CODIGO_RECIBIDO')
  ) {
    return 'Promociones';
  }


  // ------------------------------------------------------------
  // CONFIGURACIÓN
  // ------------------------------------------------------------

  if (
    n.includes('PARAMETRO') ||
    n.includes('CONFIG')
  ) {
    return 'Configuración';
  }


  return 'Otros';
}


/**
 * Obtiene una pequeña muestra de una hoja.
 *
 * Útil posteriormente para identificar tipos de información
 * sin descargar miles de registros.
 */
function WEB_getMuestraHoja(nombreHoja, limite) {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      throw new Error(
        'No existe la hoja: ' + nombreHoja
      );
    }


    const ultimaFila = sheet.getLastRow();
    const ultimaColumna = sheet.getLastColumn();


    if (
      ultimaFila === 0 ||
      ultimaColumna === 0
    ) {

      return {
        ok: true,
        nombre: nombreHoja,
        encabezados: [],
        registros: []
      };

    }


    limite = Number(limite || 5);

    limite = Math.max(
      1,
      Math.min(limite, 20)
    );


    const filasLectura = Math.min(
      ultimaFila,
      limite + 1
    );


    const datos = sheet
      .getRange(
        1,
        1,
        filasLectura,
        ultimaColumna
      )
      .getDisplayValues();


    const encabezados =
      datos.length
        ? datos[0]
        : [];


    const registros =
      datos.length > 1
        ? datos.slice(1)
        : [];


    return {

      ok: true,

      nombre: nombreHoja,

      totalFilas: ultimaFila,

      totalColumnas: ultimaColumna,

      encabezados: encabezados,

      registros: registros

    };

  } catch (error) {

    console.error(error);

    return {
      ok: false,
      mensaje: error.message
    };

  }

}
