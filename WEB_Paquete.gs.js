// ============================================================================
// WEB_Paquete.gs
// ATLAS REPARTO - PAQUETE LOCAL DE VENTAS
// ============================================================================
// SOLO LECTURA.
//
// OBJETIVO:
// Entregar al navegador, en una sola llamada, los registros necesarios para
// que Resumen, Ventas y Vendedores puedan filtrar por Ruta / Vendedor / Giro
// sin volver a llamar Apps Script.
//
// Usa WEB_DATA_getValues_() de WEB_Datos.gs.
// ============================================================================

const ATLASP_TZ = 'America/Merida';

function WEB_getAtlasVentasLocalPackage() {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );
    }


    const dataVentas =
      WEB_DATA_getValues_(
        ss,
        'VENTAS'
      );

    if (
      !dataVentas ||
      dataVentas.length < 2
    ) {

      return {
        ok: true,
        sinDatos: true,
        registros: [],
        catalogos: {
          rutas: [],
          vendedores: [],
          giros: []
        },
        periodos: {}
      };

    }


    const vendedores =
      ATLASP_mapaVendedores_(
        ss
      );

    const giros =
      ATLASP_mapaGiros_(
        ss
      );

    const c =
      ATLASP_columnas_(
        dataVentas[0]
      );


    if (
      c.importe === -1
    ) {
      throw new Error(
        'No se encontró la columna Importe en VENTAS.'
      );
    }


    const ahora =
      new Date();


    // ------------------------------------------------------------
    // PERIODOS EXACTOS TOMADOS DE LOS MOTORES YA EXISTENTES
    // ------------------------------------------------------------

    const tipos = [
      'hoy',
      'semana_anterior',
      'mes_anterior',
      'mes'
    ];

    const periodos = {};

    tipos.forEach(
      tipo => {

        if (
          typeof WEBR_getPeriodo_ ===
          'function'
        ) {

          periodos[tipo] =
            WEBR_getPeriodo_(
              tipo,
              ahora
            );

        }

        else if (
          typeof WEBV_getPeriodo_ ===
          'function'
        ) {

          periodos[tipo] =
            WEBV_getPeriodo_(
              tipo,
              ahora
            );

        }

      }
    );


    // Para cubrir:
    // - este mes vs mes anterior
    // - mes anterior vs dos meses antes
    // - semana anterior vs semana previa
    // cargamos desde el inicio más antiguo de esos comparativos.
    let fechaMinKey =
      Utilities.formatDate(
        new Date(
          ahora.getFullYear(),
          ahora.getMonth() - 2,
          1,
          12,
          0,
          0
        ),
        ATLASP_TZ,
        'yyyy-MM-dd'
      );


    Object
      .values(
        periodos
      )
      .forEach(
        p => {

          [
            p.actualInicioKey,
            p.comparacionActualInicioKey,
            p.comparacionAnteriorInicioKey
          ]
            .filter(Boolean)
            .forEach(
              key => {

                if (
                  key <
                  fechaMinKey
                ) {
                  fechaMinKey =
                    key;
                }

              }
            );

        }
      );


    const registros = [];

    const rutasSet =
      new Set();

    const girosSet =
      new Set();

    const vendedoresMap = {};


    for (
      let i = 1;
      i < dataVentas.length;
      i++
    ) {

      const fila =
        dataVentas[i];


      let fecha =
        ATLASP_parseFecha_(
          ATLASP_val_(
            fila,
            c.fecha
          )
        );


      let fechaHora =
        ATLASP_parseFecha_(
          ATLASP_val_(
            fila,
            c.fechaHora
          )
        );


      if (!fecha) {
        fecha =
          fechaHora;
      }


      if (!fecha) {
        continue;
      }


      if (!fechaHora) {
        fechaHora =
          fecha;
      }


      const fechaKey =
        Utilities.formatDate(
          fecha,
          ATLASP_TZ,
          'yyyy-MM-dd'
        );


      if (
        fechaKey <
        fechaMinKey
      ) {
        continue;
      }


      const vendedorId =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.vendedor
          )
        ) ||
        'Sin vendedor';


      const vendedorNombre =
        vendedores[
          vendedorId.toLowerCase()
        ] ||
        vendedorId;


      const giroCodigo =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.giro
          )
        );


      const giroNombre =
        giros[
          giroCodigo
        ] ||
        giros[
          giroCodigo.toLowerCase()
        ] ||
        giroCodigo ||
        'Sin giro';


      const telefono =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.telefono
          )
        );


      const cliente =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.cliente
          )
        );


      const ruta =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.ruta
          )
        ) ||
        'Sin ruta';


      const idVenta =
        ATLASP_texto_(
          ATLASP_val_(
            fila,
            c.idVenta
          )
        ) ||
        '__ROW_' +
        (i + 1);


      const clienteKey =
        ATLASP_clienteKey_(
          telefono,
          cliente,
          i + 1
        );


      const registro = {

        idVenta:
          idVenta,

        fechaKey:
          fechaKey,

        diaSemana:
          Number(
            Utilities.formatDate(
              fecha,
              ATLASP_TZ,
              'u'
            )
          ),

        hora:
          Number(
            Utilities.formatDate(
              fechaHora,
              ATLASP_TZ,
              'H'
            )
          ),

        telefono:
          telefono,

        cliente:
          cliente ||
          telefono ||
          'Sin nombre',

        clienteKey:
          clienteKey,

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

        agua:
          ATLASP_numero_(
            ATLASP_val_(
              fila,
              c.agua
            )
          ),

        superIce:
          ATLASP_numero_(
            ATLASP_val_(
              fila,
              c.superIce
            )
          ),

        importe:
          ATLASP_numero_(
            ATLASP_val_(
              fila,
              c.importe
            )
          ),

        comision:
          ATLASP_numero_(
            ATLASP_val_(
              fila,
              c.comision
            )
          )

      };


      registros.push(
        registro
      );


      rutasSet.add(
        ruta
      );


      girosSet.add(
        giroNombre
      );


      vendedoresMap[
        vendedorId
      ] =
        vendedorNombre;

    }


    const catalogos = {

      rutas:
        Array
          .from(
            rutasSet
          )
          .sort(
            (a, b) =>
              String(a)
                .localeCompare(
                  String(b),
                  'es',
                  {
                    numeric: true,
                    sensitivity: 'base'
                  }
                )
          ),

      giros:
        Array
          .from(
            girosSet
          )
          .sort(
            (a, b) =>
              String(a)
                .localeCompare(
                  String(b),
                  'es',
                  {
                    sensitivity: 'base'
                  }
                )
          ),

      vendedores:
        Object
          .keys(
            vendedoresMap
          )
          .map(
            id => ({
              valor:
                id,
              nombre:
                vendedoresMap[id]
            })
          )
          .sort(
            (a, b) =>
              a.nombre
                .localeCompare(
                  b.nombre,
                  'es',
                  {
                    sensitivity: 'base'
                  }
                )
          )

    };


    return {

      ok:
        true,

      sinDatos:
        false,

      generado:
        Utilities.formatDate(
          ahora,
          ATLASP_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      horaLocal:
        Number(
          Utilities.formatDate(
            ahora,
            ATLASP_TZ,
            'H'
          )
        ),

      fechaMinKey:
        fechaMinKey,

      registros:
        registros,

      catalogos:
        catalogos,

      periodos:
        periodos,

      meta: {

        registros:
          registros.length,

        fuente:
          'VENTAS · paquete local'

      }

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

  }

}


function ATLASP_mapaVendedores_(
  ss
) {

  const resultado = {};


  const data =
    WEB_DATA_getValues_(
      ss,
      'PARAMETROS'
    );


  if (
    !data ||
    data.length < 2
  ) {

    return resultado;

  }


  const idCol =
    ATLASP_findHeader_(
      data[0],
      [
        'id_vendedor',
        'id vendedor'
      ]
    );


  const nombreCol =
    ATLASP_findHeader_(
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
      ATLASP_texto_(
        data[i][idCol]
      );


    const nombre =
      ATLASP_texto_(
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


function ATLASP_mapaGiros_(
  ss
) {

  const resultado = {};


  const data =
    WEB_DATA_getValues_(
      ss,
      'GIROS'
    );


  if (
    !data ||
    data.length < 2
  ) {

    return resultado;

  }


  const idCol =
    ATLASP_findHeader_(
      data[0],
      [
        'idgiro',
        'id_giro'
      ]
    );


  const nombreCol =
    ATLASP_findHeader_(
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
      ATLASP_texto_(
        data[i][idCol]
      );


    const nombre =
      ATLASP_texto_(
        data[i][nombreCol]
      );


    if (id) {

      resultado[id] =
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


function ATLASP_columnas_(
  headers
) {

  return {

    idVenta:
      ATLASP_findHeader_(
        headers,
        [
          'id_venta',
          'id venta'
        ]
      ),

    telefono:
      ATLASP_findHeader_(
        headers,
        [
          'telefono_cliente',
          'telefono cliente',
          'telefono'
        ]
      ),

    cliente:
      ATLASP_findHeader_(
        headers,
        [
          'cliente',
          'nombre_cliente'
        ]
      ),

    ruta:
      ATLASP_findHeader_(
        headers,
        [
          'ruta'
        ]
      ),

    vendedor:
      ATLASP_findHeader_(
        headers,
        [
          'id_vendedor',
          'vendedor'
        ]
      ),

    fechaHora:
      ATLASP_findHeader_(
        headers,
        [
          'fecha_hora',
          'fecha hora'
        ]
      ),

    fecha:
      ATLASP_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    agua:
      ATLASP_findHeader_(
        headers,
        [
          'agua 20 litros',
          'agua_20_litros',
          'agua20litros'
        ]
      ),

    superIce:
      ATLASP_findHeader_(
        headers,
        [
          'superice',
          'super ice'
        ]
      ),

    importe:
      ATLASP_findHeader_(
        headers,
        [
          'importe'
        ]
      ),

    giro:
      ATLASP_findHeader_(
        headers,
        [
          'giro'
        ]
      ),

    comision:
      ATLASP_findHeader_(
        headers,
        [
          'comision',
          'comisión'
        ]
      )

  };
}


function ATLASP_clienteKey_(
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
    ATLASP_normalizar_(
      nombre
    );


  return nom
    ? 'N:' + nom
    : 'R:' + fallback;
}


function ATLASP_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      ATLASP_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        ATLASP_normalizar_(
          nombres[i]
        )
      );


    if (
      idx !== -1
    ) {

      return idx;

    }

  }


  return -1;
}


function ATLASP_normalizar_(
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


function ATLASP_val_(
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


function ATLASP_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();
}


function ATLASP_numero_(
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


function ATLASP_parseFecha_(
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
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      Number(m[4] || 12),
      Number(m[5] || 0),
      Number(m[6] || 0)
    );

  }


  m =
    texto.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
    );


  if (m) {

    return new Date(
      Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1]),
      Number(m[4] || 12),
      Number(m[5] || 0),
      Number(m[6] || 0)
    );

  }


  const intento =
    new Date(
      valor
    );


  return isNaN(
    intento.getTime()
  )
    ? null
    : intento;
}
