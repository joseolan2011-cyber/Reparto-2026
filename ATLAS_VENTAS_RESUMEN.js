// ============================================================================
// ATLAS_VENTAS_RESUMEN.gs
// REPARTO 2026 - TABLA ANALITICA MATERIALIZADA PARA ATLAS
// ============================================================================
//
// OBJETIVO:
// - Tomar VENTAS como fuente oficial.
// - Crear / mantener ATLAS_VENTAS_RESUMEN.
// - Agrupar por:
//     Fecha + Ruta + Vendedor + Giro
// - Evitar que Atlas tenga que procesar VENTAS completa en cada filtro.
//
// IMPORTANTE:
// - NO modifica VENTAS.
// - NO modifica PARAMETROS.
// - NO modifica GIROS.
// - SOLO escribe en ATLAS_VENTAS_RESUMEN.
// - Puede ejecutarse manualmente o cada 5 minutos.
//
// NOTA SOBRE CLIENTES:
// - La columna Clientes es el número de clientes únicos DENTRO de esa fila.
// - La columna ClientesKeys conserva las llaves para poder obtener el número
//   exacto de clientes únicos al combinar varios días/rutas/giros, sin volver
//   a leer VENTAS.
// ============================================================================


const AVR_TZ = 'America/Merida';

const AVR_HOJA_VENTAS = 'VENTAS';
const AVR_HOJA_PARAMETROS = 'PARAMETROS';
const AVR_HOJA_GIROS = 'GIROS';
const AVR_HOJA_RESUMEN = 'ATLAS_VENTAS_RESUMEN';

const AVR_FUNCION_TRIGGER = 'ATLAS_VENTAS_RESUMEN_ACTUALIZAR';


const AVR_ENCABEZADOS = [
  'Fecha',
  'FechaKey',
  'Ruta',
  'id_vendedor',
  'Vendedor',
  'GiroCodigo',
  'Giro',
  'Venta',
  'Agua',
  'SuperIce',
  'Comision',
  'Tickets',
  'Clientes',
  'ClientesKeys',
  'Ultima_Actualizacion'
];


// ============================================================================
// CONFIGURACION INICIAL
// ============================================================================

/**
 * Ejecutar UNA SOLA VEZ manualmente.
 *
 * Hace tres cosas:
 * 1. Crea la hoja ATLAS_VENTAS_RESUMEN si no existe.
 * 2. Construye el resumen inmediatamente.
 * 3. Instala un trigger cada 5 minutos.
 */
function ATLAS_VENTAS_RESUMEN_CONFIGURAR() {

  const resultado =
    ATLAS_VENTAS_RESUMEN_ACTUALIZAR();


  if (
    !resultado ||
    !resultado.ok
  ) {

    return resultado;

  }


  const trigger =
    AVR_instalarTrigger_();


  return {

    ok:
      true,

    hoja:
      AVR_HOJA_RESUMEN,

    filasResumen:
      resultado.filasResumen,

    ventasLeidas:
      resultado.ventasLeidas,

    trigger:
      trigger,

    mensaje:
      'ATLAS_VENTAS_RESUMEN quedó creada y programada cada 5 minutos.'

  };

}


// ============================================================================
// ACTUALIZACION PRINCIPAL
// ============================================================================

/**
 * Reconstruye ATLAS_VENTAS_RESUMEN desde VENTAS.
 *
 * El proceso es intencionalmente completo para mantener consistencia si una
 * venta existente es corregida posteriormente.
 */
function ATLAS_VENTAS_RESUMEN_ACTUALIZAR() {

  const lock =
    LockService.getScriptLock();


  if (
    !lock.tryLock(
      1000
    )
  ) {

    return {

      ok:
        true,

      omitido:
        true,

      mensaje:
        'Ya existe otra actualización de ATLAS_VENTAS_RESUMEN en curso.'

    };

  }


  try {

    const inicioMs =
      Date.now();


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    if (!ss) {

      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );

    }


    const shVentas =
      ss.getSheetByName(
        AVR_HOJA_VENTAS
      );


    if (!shVentas) {

      throw new Error(
        'No existe la hoja VENTAS.'
      );

    }


    const ultimaFila =
      shVentas.getLastRow();


    const ultimaColumna =
      shVentas.getLastColumn();


    const hojaResumen =
      AVR_obtenerOCrearHojaResumen_(
        ss
      );


    if (
      ultimaFila < 2 ||
      ultimaColumna < 1
    ) {

      AVR_escribirResumen_(
        hojaResumen,
        []
      );


      return {

        ok:
          true,

        ventasLeidas:
          0,

        filasResumen:
          0,

        duracionMs:
          Date.now() -
          inicioMs

      };

    }


    const data =
      shVentas
        .getRange(
          1,
          1,
          ultimaFila,
          ultimaColumna
        )
        .getValues();


    const columnas =
      AVR_columnasVentas_(
        data[0]
      );


    if (
      columnas.importe === -1
    ) {

      throw new Error(
        'No se encontró la columna Importe en VENTAS.'
      );

    }


    if (
      columnas.fecha === -1 &&
      columnas.fechaHora === -1
    ) {

      throw new Error(
        'No se encontró Fecha ni fecha_hora en VENTAS.'
      );

    }


    const mapaVendedores =
      AVR_mapaVendedores_(
        ss
      );


    const mapaGiros =
      AVR_mapaGiros_(
        ss
      );


    const grupos = {};


    let ventasLeidas =
      0;


    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const fila =
        data[i];


      let fecha =
        AVR_parseFecha_(
          AVR_val_(
            fila,
            columnas.fecha
          )
        );


      let fechaHora =
        AVR_parseFecha_(
          AVR_val_(
            fila,
            columnas.fechaHora
          )
        );


      if (!fecha) {

        fecha =
          fechaHora;

      }


      if (!fecha) {

        continue;

      }


      ventasLeidas++;


      const fechaKey =
        Utilities.formatDate(
          fecha,
          AVR_TZ,
          'yyyy-MM-dd'
        );


      const fechaDia =
        AVR_fechaDesdeKey_(
          fechaKey
        );


      const ruta =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.ruta
          )
        ) ||
        'Sin ruta';


      const vendedorId =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.vendedor
          )
        ) ||
        'Sin vendedor';


      const vendedorNombre =
        mapaVendedores[
          vendedorId.toLowerCase()
        ] ||
        vendedorId;


      const giroCodigo =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.giro
          )
        );


      const giroNombre =
        mapaGiros[
          giroCodigo
        ] ||
        mapaGiros[
          giroCodigo.toLowerCase()
        ] ||
        giroCodigo ||
        'Sin giro';


      const telefono =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.telefono
          )
        );


      const cliente =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.cliente
          )
        );


      const clienteKey =
        AVR_clienteKey_(
          telefono,
          cliente,
          i + 1
        );


      const idVenta =
        AVR_texto_(
          AVR_val_(
            fila,
            columnas.idVenta
          )
        ) ||
        '__ROW_' +
        (i + 1);


      const key =
        [
          fechaKey,
          ruta,
          vendedorId,
          giroCodigo,
          giroNombre
        ].join(
          '\u001F'
        );


      if (
        !grupos[key]
      ) {

        grupos[key] = {

          fecha:
            fechaDia,

          fechaKey:
            fechaKey,

          ruta:
            ruta,

          vendedorId:
            vendedorId,

          vendedor:
            vendedorNombre,

          giroCodigo:
            giroCodigo,

          giro:
            giroNombre,

          venta:
            0,

          agua:
            0,

          superIce:
            0,

          comision:
            0,

          tickets:
            new Set(),

          clientes:
            new Set()

        };

      }


      grupos[key].venta +=
        AVR_numero_(
          AVR_val_(
            fila,
            columnas.importe
          )
        );


      grupos[key].agua +=
        AVR_numero_(
          AVR_val_(
            fila,
            columnas.agua
          )
        );


      grupos[key].superIce +=
        AVR_numero_(
          AVR_val_(
            fila,
            columnas.superIce
          )
        );


      grupos[key].comision +=
        AVR_numero_(
          AVR_val_(
            fila,
            columnas.comision
          )
        );


      grupos[key]
        .tickets
        .add(
          idVenta
        );


      grupos[key]
        .clientes
        .add(
          clienteKey
        );

    }


    const ahora =
      new Date();


    const filas =
      Object
        .values(
          grupos
        )
        .sort(
          (a, b) => {

            if (
              a.fechaKey !==
              b.fechaKey
            ) {

              return a.fechaKey <
                b.fechaKey
                ? -1
                : 1;

            }


            const rutaCmp =
              String(
                a.ruta
              )
                .localeCompare(
                  String(
                    b.ruta
                  ),
                  'es',
                  {
                    numeric:
                      true,
                    sensitivity:
                      'base'
                  }
                );


            if (
              rutaCmp !== 0
            ) {

              return rutaCmp;

            }


            const vendedorCmp =
              String(
                a.vendedor
              )
                .localeCompare(
                  String(
                    b.vendedor
                  ),
                  'es',
                  {
                    sensitivity:
                      'base'
                  }
                );


            if (
              vendedorCmp !== 0
            ) {

              return vendedorCmp;

            }


            return String(
              a.giro
            )
              .localeCompare(
                String(
                  b.giro
                ),
                'es',
                {
                  sensitivity:
                    'base'
                }
              );

          }
        )
        .map(
          g => {

            const clientesKeys =
              Array
                .from(
                  g.clientes
                )
                .sort();


            return [

              g.fecha,

              g.fechaKey,

              g.ruta,

              g.vendedorId,

              g.vendedor,

              g.giroCodigo,

              g.giro,

              AVR_redondear_(
                g.venta
              ),

              AVR_redondear_(
                g.agua
              ),

              AVR_redondear_(
                g.superIce
              ),

              AVR_redondear_(
                g.comision
              ),

              g.tickets.size,

              g.clientes.size,

              clientesKeys.join(
                '|'
              ),

              ahora

            ];

          }
        );


    AVR_escribirResumen_(
      hojaResumen,
      filas
    );


    SpreadsheetApp.flush();


    return {

      ok:
        true,

      hoja:
        AVR_HOJA_RESUMEN,

      ventasLeidas:
        ventasLeidas,

      filasResumen:
        filas.length,

      reduccionPct:
        ventasLeidas
          ? AVR_redondear_(
              (
                1 -
                filas.length /
                ventasLeidas
              ) *
              100
            )
          : 0,

      actualizado:
        Utilities.formatDate(
          ahora,
          AVR_TZ,
          'dd/MM/yyyy HH:mm:ss'
        ),

      duracionMs:
        Date.now() -
        inicioMs

    };


  } catch (error) {

    console.error(
      error
    );


    return {

      ok:
        false,

      mensaje:
        error.message,

      stack:
        error.stack

    };


  } finally {

    lock.releaseLock();

  }

}


// ============================================================================
// TRIGGER
// ============================================================================

/**
 * Instala un trigger cada 5 minutos.
 * Si ya existe, no crea duplicados.
 */
function ATLAS_VENTAS_RESUMEN_INSTALAR_TRIGGER() {

  return AVR_instalarTrigger_();

}


/**
 * Elimina únicamente los triggers de este resumen.
 */
function ATLAS_VENTAS_RESUMEN_QUITAR_TRIGGER() {

  const triggers =
    ScriptApp.getProjectTriggers();


  let eliminados =
    0;


  triggers.forEach(
    trigger => {

      if (
        trigger.getHandlerFunction() ===
        AVR_FUNCION_TRIGGER
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );


        eliminados++;

      }

    }
  );


  return {

    ok:
      true,

    eliminados:
      eliminados

  };

}


function AVR_instalarTrigger_() {

  const triggers =
    ScriptApp.getProjectTriggers();


  const existente =
    triggers.some(
      trigger =>
        trigger.getHandlerFunction() ===
        AVR_FUNCION_TRIGGER
    );


  if (
    existente
  ) {

    return {

      creado:
        false,

      frecuencia:
        '5 minutos',

      mensaje:
        'El trigger ya existía.'

    };

  }


  ScriptApp
    .newTrigger(
      AVR_FUNCION_TRIGGER
    )
    .timeBased()
    .everyMinutes(
      5
    )
    .create();


  return {

    creado:
      true,

    frecuencia:
      '5 minutos',

    mensaje:
      'Trigger creado correctamente.'

  };

}


// ============================================================================
// HOJA DESTINO
// ============================================================================

function AVR_obtenerOCrearHojaResumen_(
  ss
) {

  let hoja =
    ss.getSheetByName(
      AVR_HOJA_RESUMEN
    );


  if (!hoja) {

    hoja =
      ss.insertSheet(
        AVR_HOJA_RESUMEN
      );

  }


  return hoja;

}


function AVR_escribirResumen_(
  hoja,
  filas
) {

  if (!hoja) {

    throw new Error(
      'No fue posible obtener la hoja de resumen.'
    );

  }


  const columnas =
    AVR_ENCABEZADOS.length;


  // Limpiamos únicamente la hoja analítica.
  hoja.clearContents();


  hoja
    .getRange(
      1,
      1,
      1,
      columnas
    )
    .setValues([
      AVR_ENCABEZADOS
    ]);


  if (
    filas.length
  ) {

    hoja
      .getRange(
        2,
        1,
        filas.length,
        columnas
      )
      .setValues(
        filas
      );

  }


  hoja
    .setFrozenRows(
      1
    );


  hoja
    .getRange(
      'A:A'
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  hoja
    .getRange(
      'H:K'
    )
    .setNumberFormat(
      '#,##0.00'
    );


  hoja
    .getRange(
      'L:M'
    )
    .setNumberFormat(
      '0'
    );


  hoja
    .getRange(
      'O:O'
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm:ss'
    );


  hoja
    .autoResizeColumns(
      1,
      Math.min(
        columnas,
        13
      )
    );


  // ClientesKeys puede ser largo; evitamos que ensanche toda la hoja.
  hoja
    .setColumnWidth(
      14,
      180
    );

}


// ============================================================================
// MAPA DE VENDEDORES
// ============================================================================

function AVR_mapaVendedores_(
  ss
) {

  const resultado = {};


  const hoja =
    ss.getSheetByName(
      AVR_HOJA_PARAMETROS
    );


  if (!hoja) {

    return resultado;

  }


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2 ||
    ultimaColumna < 1
  ) {

    return resultado;

  }


  const data =
    hoja
      .getRange(
        1,
        1,
        ultimaFila,
        ultimaColumna
      )
      .getValues();


  const idCol =
    AVR_findHeader_(
      data[0],
      [
        'id_vendedor',
        'id vendedor'
      ]
    );


  const nombreCol =
    AVR_findHeader_(
      data[0],
      [
        'nombre',
        'vendedor_nombre',
        'vendedor nombre'
      ]
    );


  if (
    idCol === -1 ||
    nombreCol === -1
  ) {

    return resultado;

  }


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const id =
      AVR_texto_(
        data[i][idCol]
      );


    const nombre =
      AVR_texto_(
        data[i][nombreCol]
      );


    if (id) {

      resultado[
        id.toLowerCase()
      ] =
        nombre ||
        id;

    }

  }


  return resultado;

}


// ============================================================================
// MAPA DE GIROS
// ============================================================================

function AVR_mapaGiros_(
  ss
) {

  const resultado = {};


  const hoja =
    ss.getSheetByName(
      AVR_HOJA_GIROS
    );


  if (!hoja) {

    return resultado;

  }


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2 ||
    ultimaColumna < 1
  ) {

    return resultado;

  }


  const data =
    hoja
      .getRange(
        1,
        1,
        ultimaFila,
        ultimaColumna
      )
      .getValues();


  const idCol =
    AVR_findHeader_(
      data[0],
      [
        'idgiro',
        'id_giro'
      ]
    );


  const nombreCol =
    AVR_findHeader_(
      data[0],
      [
        'giro'
      ]
    );


  if (
    idCol === -1 ||
    nombreCol === -1
  ) {

    return resultado;

  }


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const id =
      AVR_texto_(
        data[i][idCol]
      );


    const nombre =
      AVR_texto_(
        data[i][nombreCol]
      );


    if (id) {

      resultado[
        id
      ] =
        nombre ||
        id;


      resultado[
        id.toLowerCase()
      ] =
        nombre ||
        id;

    }

  }


  return resultado;

}


// ============================================================================
// COLUMNAS VENTAS
// ============================================================================

function AVR_columnasVentas_(
  headers
) {

  return {

    idVenta:
      AVR_findHeader_(
        headers,
        [
          'id_venta',
          'id venta'
        ]
      ),

    telefono:
      AVR_findHeader_(
        headers,
        [
          'telefono_cliente',
          'telefono cliente',
          'telefono'
        ]
      ),

    cliente:
      AVR_findHeader_(
        headers,
        [
          'cliente',
          'nombre_cliente'
        ]
      ),

    ruta:
      AVR_findHeader_(
        headers,
        [
          'ruta'
        ]
      ),

    vendedor:
      AVR_findHeader_(
        headers,
        [
          'id_vendedor',
          'vendedor'
        ]
      ),

    fechaHora:
      AVR_findHeader_(
        headers,
        [
          'fecha_hora',
          'fecha hora'
        ]
      ),

    fecha:
      AVR_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    agua:
      AVR_findHeader_(
        headers,
        [
          'agua 20 litros',
          'agua_20_litros',
          'agua20litros'
        ]
      ),

    superIce:
      AVR_findHeader_(
        headers,
        [
          'superice',
          'super ice'
        ]
      ),

    importe:
      AVR_findHeader_(
        headers,
        [
          'importe'
        ]
      ),

    giro:
      AVR_findHeader_(
        headers,
        [
          'giro'
        ]
      ),

    comision:
      AVR_findHeader_(
        headers,
        [
          'comision',
          'comisión'
        ]
      )

  };

}


// ============================================================================
// UTILIDADES
// ============================================================================

function AVR_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      AVR_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const buscado =
      AVR_normalizar_(
        nombres[i]
      );


    const indice =
      normalizados.indexOf(
        buscado
      );


    if (
      indice !== -1
    ) {

      return indice;

    }

  }


  return -1;

}


function AVR_normalizar_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  )
    .normalize(
      'NFD'
    )
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .trim()
    .toLowerCase()
    .replace(
      /[^a-z0-9]+/g,
      '_'
    )
    .replace(
      /^_+|_+$/g,
      ''
    );

}


function AVR_val_(
  fila,
  indice
) {

  if (
    indice === -1 ||
    indice === null ||
    indice === undefined
  ) {

    return '';

  }


  return fila[
    indice
  ];

}


function AVR_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function AVR_numero_(
  valor
) {

  if (
    typeof valor ===
    'number'
  ) {

    return isFinite(
      valor
    )
      ? valor
      : 0;

  }


  const limpio =
    String(
      valor || ''
    )
      .replace(
        /,/g,
        ''
      )
      .replace(
        /[^0-9.\-]/g,
        ''
      );


  const numero =
    Number(
      limpio
    );


  return isFinite(
    numero
  )
    ? numero
    : 0;

}


function AVR_redondear_(
  numero
) {

  return Math.round(
    Number(
      numero || 0
    ) *
    100
  ) /
  100;

}


function AVR_clienteKey_(
  telefono,
  nombre,
  fallback
) {

  let tel =
    String(
      telefono || ''
    )
      .replace(
        /\D/g,
        ''
      );


  if (
    tel.length > 10
  ) {

    tel =
      tel.slice(
        -10
      );

  }


  if (tel) {

    return 'T:' +
      tel;

  }


  const nom =
    AVR_normalizar_(
      nombre
    );


  return nom
    ? 'N:' +
      nom
    : 'R:' +
      fallback;

}


function AVR_fechaDesdeKey_(
  key
) {

  const p =
    String(
      key || ''
    )
      .split(
        '-'
      );


  if (
    p.length !== 3
  ) {

    return new Date();

  }


  return new Date(
    Number(
      p[0]
    ),
    Number(
      p[1]
    ) -
    1,
    Number(
      p[2]
    ),
    12,
    0,
    0,
    0
  );

}


function AVR_parseFecha_(
  valor
) {

  if (
    valor instanceof Date &&
    !isNaN(
      valor.getTime()
    )
  ) {

    return valor;

  }


  if (
    typeof valor ===
    'number' &&
    valor > 20000 &&
    valor < 100000
  ) {

    const base =
      new Date(
        Date.UTC(
          1899,
          11,
          30
        )
      );


    return new Date(
      base.getTime() +
      valor *
      86400000
    );

  }


  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {

    return null;

  }


  const texto =
    String(
      valor
    ).trim();


  let m =
    texto.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
    );


  if (m) {

    return new Date(
      Number(
        m[1]
      ),
      Number(
        m[2]
      ) -
      1,
      Number(
        m[3]
      ),
      Number(
        m[4] || 12
      ),
      Number(
        m[5] || 0
      ),
      Number(
        m[6] || 0
      )
    );

  }


  m =
    texto.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
    );


  if (m) {

    return new Date(
      Number(
        m[3]
      ),
      Number(
        m[2]
      ) -
      1,
      Number(
        m[1]
      ),
      Number(
        m[4] || 12
      ),
      Number(
        m[5] || 0
      ),
      Number(
        m[6] || 0
      )
    );

  }


  const intento =
    new Date(
      texto
    );


  return isNaN(
    intento.getTime()
  )
    ? null
    : intento;

}


// ============================================================================
// DIAGNOSTICO
// ============================================================================

function ATLAS_VENTAS_RESUMEN_DIAGNOSTICO() {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const ventas =
      ss.getSheetByName(
        AVR_HOJA_VENTAS
      );


    const resumen =
      ss.getSheetByName(
        AVR_HOJA_RESUMEN
      );


    const triggerExiste =
      ScriptApp
        .getProjectTriggers()
        .some(
          trigger =>
            trigger.getHandlerFunction() ===
            AVR_FUNCION_TRIGGER
        );


    return {

      ok:
        true,

      ventasExiste:
        !!ventas,

      ventasFilas:
        ventas
          ? Math.max(
              0,
              ventas.getLastRow() -
              1
            )
          : 0,

      resumenExiste:
        !!resumen,

      resumenFilas:
        resumen
          ? Math.max(
              0,
              resumen.getLastRow() -
              1
            )
          : 0,

      triggerCada5Min:
        triggerExiste

    };


  } catch (error) {

    return {

      ok:
        false,

      mensaje:
        error.message

    };

  }

}
