// ============================================================
// WEB_Datos.gs
// CAPA DE LECTURA + CACHE COMPARTIDA - REPARTO 2026
// ============================================================
// SOLO LECTURA.
// ============================================================

const WEB_DATA_CACHE_TTL_SEGUNDOS = 180;
const WEB_DATA_CACHE_CHUNK = 60000;
const WEB_DATA_CACHE_VERSION = 'ATLAS_R1';

function WEB_DATA_getValues_(ss, nombreHoja, forzar) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('No fue posible identificar el Spreadsheet activo.');
  }

  const hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) return [];

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila <= 0 || ultimaColumna <= 0) return [];

  const cache = CacheService.getScriptCache();

  const baseKey = WEB_DATA_cacheKey_(
    ss,
    hoja,
    ultimaFila,
    ultimaColumna
  );

  const metaKey = baseKey + '_META';

  if (!forzar) {
    const metaTexto = cache.get(metaKey);

    if (metaTexto) {
      try {
        const meta = JSON.parse(metaTexto);
        const claves = [];

        for (let i = 0; i < meta.partes; i++) {
          claves.push(baseKey + '_P' + i);
        }

        const partes = cache.getAll(claves);

        let completo = true;
        let json = '';

        for (let i = 0; i < claves.length; i++) {
          if (
            partes[claves[i]] === undefined ||
            partes[claves[i]] === null
          ) {
            completo = false;
            break;
          }

          json += partes[claves[i]];
        }

        if (completo) {
          const datos = JSON.parse(json);

          if (Array.isArray(datos)) {
            return datos;
          }
        }
      } catch (errorCache) {
        console.warn(
          'Cache Atlas inválida para ' +
          nombreHoja +
          ': ' +
          errorCache.message
        );
      }
    }
  }

  const datos = hoja
    .getRange(
      1,
      1,
      ultimaFila,
      ultimaColumna
    )
    .getValues();

  WEB_DATA_guardarCache_(
    cache,
    baseKey,
    metaKey,
    datos
  );

  return datos;
}

function WEB_precalentarDatosAtlas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );
    }

    const hojas = [
      'VENTAS',
      'PARAMETROS',
      'GIROS'
    ];

    const resultado = {
      ok: true,
      generado: new Date(),
      hojas: []
    };

    hojas.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);

      if (!hoja) {
        resultado.hojas.push({
          nombre: nombre,
          existe: false,
          filas: 0
        });

        return;
      }

      const datos = WEB_DATA_getValues_(
        ss,
        nombre,
        true
      );

      resultado.hojas.push({
        nombre: nombre,
        existe: true,
        filas: Math.max(0, datos.length - 1)
      });
    });

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

function WEB_getEstadoCacheAtlas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );
    }

    const hojas = [
      'VENTAS',
      'PARAMETROS',
      'GIROS'
    ];

    return {
      ok: true,
      ttlSegundos: WEB_DATA_CACHE_TTL_SEGUNDOS,
      hojas: hojas.map(nombre => {
        const hoja = ss.getSheetByName(nombre);

        return {
          nombre: nombre,
          existe: !!hoja,
          filas:
            hoja
              ? Math.max(0, hoja.getLastRow() - 1)
              : 0,
          columnas:
            hoja
              ? hoja.getLastColumn()
              : 0
        };
      })
    };

  } catch (error) {
    return {
      ok: false,
      mensaje: error.message
    };
  }
}

function WEB_DATA_guardarCache_(
  cache,
  baseKey,
  metaKey,
  datos
) {
  try {
    const json = JSON.stringify(datos);
    const partes = [];

    for (
      let i = 0;
      i < json.length;
      i += WEB_DATA_CACHE_CHUNK
    ) {
      partes.push(
        json.substring(
          i,
          i + WEB_DATA_CACHE_CHUNK
        )
      );
    }

    const guardar = {};

    partes.forEach((parte, indice) => {
      guardar[
        baseKey +
        '_P' +
        indice
      ] = parte;
    });

    guardar[metaKey] = JSON.stringify({
      partes: partes.length,
      creado: Date.now()
    });

    cache.putAll(
      guardar,
      WEB_DATA_CACHE_TTL_SEGUNDOS
    );

  } catch (error) {
    console.warn(
      'No fue posible guardar cache Atlas: ' +
      error.message
    );
  }
}

function WEB_DATA_cacheKey_(
  ss,
  hoja,
  filas,
  columnas
) {
  const idArchivo = String(
    ss.getId() || ''
  )
    .replace(
      /[^a-zA-Z0-9]/g,
      ''
    )
    .slice(-24);

  const nombre = String(
    hoja.getName() || ''
  )
    .replace(
      /[^a-zA-Z0-9]/g,
      '_'
    )
    .slice(0, 45);

  return [
    WEB_DATA_CACHE_VERSION,
    idArchivo,
    hoja.getSheetId(),
    nombre,
    filas + 'x' + columnas
  ].join('_');
}

function WEB_getDiagnosticoBase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );
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
          .getRange(
            1,
            1,
            1,
            ultimaColumna
          )
          .getDisplayValues()[0]
          .map(valor =>
            String(valor || '').trim()
          );
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
        tipo: WEB_clasificarHoja_(
          nombre,
          encabezados
        )
      });
    });

    resultado.hojas.sort(
      (a, b) =>
        b.registros -
        a.registros
    );

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

function WEB_clasificarHoja_(
  nombre,
  encabezados
) {
  const n = String(
    nombre || ''
  ).toUpperCase();

  const headers = encabezados
    .join(' ')
    .toUpperCase();

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

  if (
    n.includes('PEDIDO') ||
    (
      headers.includes('ID_PEDIDO') &&
      headers.includes('ESTADO')
    )
  ) {
    return 'Pedidos';
  }

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

  if (n.includes('TRANSAC')) {
    return 'Transacciones';
  }

  if (
    n.includes('WEBHOOK') ||
    headers.includes('SESSION_ID')
  ) {
    return 'Automatización';
  }

  if (
    n.includes('RECORRIDO') ||
    n.includes('RUTA')
  ) {
    return 'Rutas';
  }

  if (n.includes('LIQUID')) {
    return 'Liquidaciones';
  }

  if (
    n.includes('PROMO') ||
    headers.includes('CODIGO_ENVIADO') ||
    headers.includes('CODIGO_RECIBIDO')
  ) {
    return 'Promociones';
  }

  if (
    n.includes('PARAMETRO') ||
    n.includes('CONFIG')
  ) {
    return 'Configuración';
  }

  return 'Otros';
}

function WEB_getMuestraHoja(
  nombreHoja,
  limite
) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      throw new Error(
        'No existe la hoja: ' +
        nombreHoja
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
