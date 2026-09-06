// ============================================================================
// WEB_Clientes.gs
// ATLAS REPARTO - ANÁLISIS DE CARTERA
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   CLIENTES
//   VENTAS
//   PARAMETROS
//
// Analiza:
//   - cartera
//   - actividad
//   - frecuencia individual
//   - clientes en atención
//   - clientes en riesgo
//   - clientes dormidos
//   - altas del mes
//   - reactivaciones
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBCL_TZ = 'America/Merida';

const WEBCL_HOJA_CLIENTES = 'CLIENTES';
const WEBCL_HOJA_VENTAS = 'VENTAS';
const WEBCL_HOJA_PARAMETROS = 'PARAMETROS';
const WEBCL_HOJA_GIROS = 'GIROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getClientesDashboard(filtros) {

  try {

    filtros = filtros || {};

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const shClientes =
      ss.getSheetByName(
        WEBCL_HOJA_CLIENTES
      );


    const shVentas =
      ss.getSheetByName(
        WEBCL_HOJA_VENTAS
      );


    if (!shClientes) {

      throw new Error(
        'No existe la hoja CLIENTES.'
      );

    }


    if (!shVentas) {

      throw new Error(
        'No existe la hoja VENTAS.'
      );

    }


    const vendedores =
      WEBCL_getMapaVendedores_(ss);


    const mapaGiros =
      WEBCL_getMapaGiros_(ss);


    const hoy =
      WEBCL_hoyLocal_();


    const inicioMes =
      new Date(
        hoy.getFullYear(),
        hoy.getMonth(),
        1,
        12,
        0,
        0
      );


    // ========================================================================
    // HISTORIAL DE VENTAS
    // ========================================================================

    const historialVentas =
      WEBCL_cargarHistorialVentas_(
        shVentas
      );


    // ========================================================================
    // CLIENTES
    // ========================================================================

    const dataClientes =
      shClientes
        .getDataRange()
        .getValues();


    if (
      !dataClientes ||
      dataClientes.length < 2
    ) {

      return {

        ok: true,

        sinDatos: true,

        mensaje:
          'La hoja CLIENTES no contiene información.'

      };

    }


    const headers =
      dataClientes[0];


    const c =
      WEBCL_columnasClientes_(
        headers
      );


    const clientes = [];


    for (
      let i = 1;
      i < dataClientes.length;
      i++
    ) {

      const fila =
        dataClientes[i];


      const telefono =
        WEBCL_texto_(
          WEBCL_val_(
            fila,
            c.telefono
          )
        );


      const nombre =
        WEBCL_texto_(
          WEBCL_val_(
            fila,
            c.nombre
          )
        );


      if (
        !telefono &&
        !nombre
      ) {
        continue;
      }


      const key =
        WEBCL_clienteKey_(
          telefono,
          nombre
        );


      const vendedorId =
        WEBCL_texto_(
          WEBCL_val_(
            fila,
            c.vendedor
          )
        );


      const vendedorFila =
        WEBCL_texto_(
          WEBCL_val_(
            fila,
            c.vendedorNombre
          )
        );


      const vendedorNombre =
        vendedorFila ||
        vendedores[
          vendedorId.toLowerCase()
        ] ||
        vendedorId ||
        'Sin vendedor';


      const giroCodigo =
        WEBCL_texto_(
          WEBCL_val_(
            fila,
            c.giro
          )
        );


      const giro =
        mapaGiros[
          giroCodigo
        ] ||
        mapaGiros[
          giroCodigo.toLowerCase()
        ] ||
        giroCodigo ||
        'Sin giro';


      const fechaRegistro =
        WEBCL_parseFecha_(
          WEBCL_val_(
            fila,
            c.fechaRegistro
          )
        );


      const activo =
        WEBCL_esActivo_(
          WEBCL_val_(
            fila,
            c.activo
          ),
          WEBCL_val_(
            fila,
            c.estatus
          )
        );


      const ventasCliente =
        historialVentas[key] || {
          fechas: [],
          movimientos: [],
          importeHistorico: 0,
          importeMes: 0,
          aguaMes: 0,
          iceMes: 0
        };


      const analisis =
        WEBCL_analizarCliente_(
          ventasCliente.fechas,
          fechaRegistro,
          hoy
        );


      const reactivado =
        WEBCL_esReactivadoMes_(
          ventasCliente.fechas,
          inicioMes
        );


      const nuevoMes =
        fechaRegistro
          ? WEBCL_mismoMes_(
              fechaRegistro,
              hoy
            )
          : false;


      clientes.push({

        key:
          key,

        telefono:
          telefono,

        nombre:
          nombre ||
          telefono ||
          'Sin nombre',

        activo:
          activo,

        vendedorId:
          vendedorId,

        vendedor:
          vendedorNombre,

        giroCodigo:
          giroCodigo,

        giro:
          giro,

        colonia:
          WEBCL_texto_(
            WEBCL_val_(
              fila,
              c.colonia
            )
          ) ||
          'Sin colonia',

        fechaRegistro:
          fechaRegistro
            ? WEBCL_fechaKey_(
                fechaRegistro
              )
            : '',

        nuevoMes:
          nuevoMes,

        reactivadoMes:
          reactivado,

        ultimaCompra:
          analisis.ultimaCompra,

        diasSinComprar:
          analisis.diasSinComprar,

        ciclo:
          analisis.ciclo,

        atraso:
          analisis.atraso,

        factor:
          analisis.factor,

        comprasHistoricas:
          ventasCliente.fechas.length,

        importeHistorico:
          WEBCL_redondear_(
            ventasCliente.importeHistorico
          ),

        importeMes:
          WEBCL_redondear_(
            ventasCliente.importeMes
          ),

        aguaMes:
          WEBCL_redondear_(
            ventasCliente.aguaMes
          ),

        iceMes:
          WEBCL_redondear_(
            ventasCliente.iceMes
          ),

        estadoSalud:
          activo
            ? analisis.estado
            : 'Inactivo'

      });

    }


    // ========================================================================
    // FILTROS
    // ========================================================================

    const filtrados =
      clientes.filter(
        x =>
          WEBCL_cumpleFiltros_(
            x,
            filtros
          )
      );


    const activos =
      filtrados.filter(
        x => x.activo
      );


    // ========================================================================
    // KPIs
    // ========================================================================

    const alDia =
      activos.filter(
        x =>
          x.estadoSalud ===
          'Al día'
      );


    const atencion =
      activos.filter(
        x =>
          x.estadoSalud ===
          'Atención'
      );


    const riesgo =
      activos.filter(
        x =>
          x.estadoSalud ===
          'Riesgo'
      );


    const dormidos =
      activos.filter(
        x =>
          x.estadoSalud ===
          'Dormido'
      );


    const aprendiendo =
      activos.filter(
        x =>
          x.estadoSalud ===
          'Aprendiendo'
      );


    const nuevosMes =
      filtrados.filter(
        x =>
          x.nuevoMes
      );


    const reactivadosMes =
      activos.filter(
        x =>
          x.reactivadoMes
      );


    // ========================================================================
    // TABLA DE ATENCIÓN
    // ========================================================================

    const prioridad = {
      Dormido: 4,
      Riesgo: 3,
      Atención: 2
    };


    const requierenAtencion =
      activos

        .filter(
          x =>
            [
              'Atención',
              'Riesgo',
              'Dormido'
            ].includes(
              x.estadoSalud
            )
        )

        .sort(
          (a, b) => {

            const pa =
              prioridad[
                a.estadoSalud
              ] || 0;


            const pb =
              prioridad[
                b.estadoSalud
              ] || 0;


            if (
              pb !== pa
            ) {
              return pb - pa;
            }


            return (
              b.diasSinComprar -
              a.diasSinComprar
            );

          }
        )

        .slice(
          0,
          50
        );


    // ========================================================================
    // NUEVOS
    // ========================================================================

    const nuevos =
      nuevosMes

        .sort(
          (a, b) =>
            String(
              b.fechaRegistro
            ).localeCompare(
              String(
                a.fechaRegistro
              )
            )
        )

        .slice(
          0,
          30
        );


    // ========================================================================
    // GIROS
    // ========================================================================

    const giros =
      WEBCL_agruparGiros_(
        activos
      );


    // ========================================================================
    // VENDEDORES CON RIESGO
    // ========================================================================

    const riesgoVendedores =
      WEBCL_riesgoPorVendedor_(
        activos
      );


    // ========================================================================
    // INSIGHTS
    // ========================================================================

    const insights =
      WEBCL_crearInsights_(
        activos,
        alDia,
        atencion,
        riesgo,
        dormidos,
        nuevosMes,
        reactivadosMes,
        riesgoVendedores
      );


    // ========================================================================
    // CATÁLOGOS
    // ========================================================================

    const catalogos =
      WEBCL_catalogos_(
        clientes
      );


    return {

      ok: true,

      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBCL_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      kpis: {

        cartera:
          filtrados.length,

        activos:
          activos.length,

        alDia:
          alDia.length,

        atencion:
          atencion.length,

        riesgo:
          riesgo.length,

        dormidos:
          dormidos.length,

        nuevos:
          nuevosMes.length,

        reactivados:
          reactivadosMes.length,

        aprendiendo:
          aprendiendo.length

      },

      salud: {

        alDia:
          alDia.length,

        atencion:
          atencion.length,

        riesgo:
          riesgo.length,

        dormidos:
          dormidos.length,

        aprendiendo:
          aprendiendo.length

      },

      clientes:
        filtrados,

      requiereAtencion:
        requierenAtencion,

      nuevos:
        nuevos,

      giros:
        giros,

      riesgoVendedores:
        riesgoVendedores,

      insights:
        insights,

      catalogos:
        catalogos

    };


  } catch (error) {

    console.error(error);


    return {

      ok: false,

      mensaje:
        error.message,

      stack:
        error.stack

    };

  }

}


// ============================================================================
// HISTORIAL DE VENTAS
// ============================================================================

function WEBCL_cargarHistorialVentas_(
  sheet
) {

  const data =
    sheet
      .getDataRange()
      .getValues();


  if (
    !data ||
    data.length < 2
  ) {
    return {};
  }


  const headers =
    data[0];


  const telCol =
    WEBCL_findHeader_(
      headers,
      [
        'telefono_cliente',
        'telefono cliente',
        'telefono'
      ]
    );


  const clienteCol =
    WEBCL_findHeader_(
      headers,
      [
        'cliente'
      ]
    );


  const fechaCol =
    WEBCL_findHeader_(
      headers,
      [
        'fecha'
      ]
    );


  const fechaHoraCol =
    WEBCL_findHeader_(
      headers,
      [
        'fecha_hora',
        'fecha hora'
      ]
    );


  const importeCol =
    WEBCL_findHeader_(
      headers,
      [
        'importe'
      ]
    );


  const aguaCol =
    WEBCL_findHeader_(
      headers,
      [
        'agua 20 litros',
        'agua_20_litros'
      ]
    );


  const iceCol =
    WEBCL_findHeader_(
      headers,
      [
        'superice',
        'super ice'
      ]
    );


  const hoy =
    WEBCL_hoyLocal_();


  const mapa = {};


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const fila =
      data[i];


    const telefono =
      WEBCL_texto_(
        WEBCL_val_(
          fila,
          telCol
        )
      );


    const cliente =
      WEBCL_texto_(
        WEBCL_val_(
          fila,
          clienteCol
        )
      );


    const key =
      WEBCL_clienteKey_(
        telefono,
        cliente
      );


    if (!key) {
      continue;
    }


    let fecha =
      WEBCL_parseFecha_(
        WEBCL_val_(
          fila,
          fechaCol
        )
      );


    if (!fecha) {

      fecha =
        WEBCL_parseFecha_(
          WEBCL_val_(
            fila,
            fechaHoraCol
          )
        );

    }


    if (!fecha) {
      continue;
    }


    if (!mapa[key]) {

      mapa[key] = {

        fechasMap: {},

        fechas: [],

        movimientos: [],

        importeHistorico: 0,

        importeMes: 0,

        aguaMes: 0,

        iceMes: 0

      };

    }


    const fechaKey =
      WEBCL_fechaKey_(
        fecha
      );


    mapa[key].fechasMap[
      fechaKey
    ] = true;


    const importe =
      WEBCL_numero_(
        WEBCL_val_(
          fila,
          importeCol
        )
      );


    const agua =
      WEBCL_numero_(
        WEBCL_val_(
          fila,
          aguaCol
        )
      );


    const ice =
      WEBCL_numero_(
        WEBCL_val_(
          fila,
          iceCol
        )
      );


    mapa[key].importeHistorico +=
      importe;


    if (
      WEBCL_mismoMes_(
        fecha,
        hoy
      )
    ) {

      mapa[key].importeMes +=
        importe;

      mapa[key].aguaMes +=
        agua;

      mapa[key].iceMes +=
        ice;

    }

  }


  Object.keys(
    mapa
  ).forEach(
    key => {

      mapa[key].fechas =
        Object.keys(
          mapa[key].fechasMap
        )

          .sort()

          .map(
            fecha =>
              WEBCL_keyToDate_(
                fecha
              )
          );

    }
  );


  return mapa;

}


// ============================================================================
// ANÁLISIS INDIVIDUAL
// ============================================================================

function WEBCL_analizarCliente_(
  fechas,
  fechaRegistro,
  hoy
) {

  fechas =
    Array.isArray(fechas)
      ? fechas.slice().sort(
          (a, b) =>
            a - b
        )
      : [];


  if (
    !fechas.length
  ) {

    const diasAlta =
      fechaRegistro
        ? WEBCL_diasEntre_(
            fechaRegistro,
            hoy
          )
        : null;


    let estado =
      'Aprendiendo';


    if (
      diasAlta !== null &&
      diasAlta > 30
    ) {

      estado =
        'Dormido';

    }

    else if (
      diasAlta !== null &&
      diasAlta > 15
    ) {

      estado =
        'Atención';

    }


    return {

      ultimaCompra:
        '',

      diasSinComprar:
        diasAlta || 0,

      ciclo:
        null,

      atraso:
        null,

      factor:
        null,

      estado:
        estado

    };

  }


  const ultima =
    fechas[
      fechas.length - 1
    ];


  const diasSinComprar =
    Math.max(
      0,
      WEBCL_diasEntre_(
        ultima,
        hoy
      )
    );


  if (
    fechas.length < 3
  ) {

    let estado =
      'Aprendiendo';


    if (
      diasSinComprar > 60
    ) {

      estado =
        'Dormido';

    }

    else if (
      diasSinComprar > 30
    ) {

      estado =
        'Atención';

    }


    return {

      ultimaCompra:
        WEBCL_fechaKey_(
          ultima
        ),

      diasSinComprar:
        diasSinComprar,

      ciclo:
        null,

      atraso:
        null,

      factor:
        null,

      estado:
        estado

    };

  }


  const intervalos = [];


  for (
    let i = 1;
    i < fechas.length;
    i++
  ) {

    const dias =
      WEBCL_diasEntre_(
        fechas[i - 1],
        fechas[i]
      );


    if (
      dias > 0
    ) {

      intervalos.push(
        dias
      );

    }

  }


  if (
    !intervalos.length
  ) {

    return {

      ultimaCompra:
        WEBCL_fechaKey_(
          ultima
        ),

      diasSinComprar:
        diasSinComprar,

      ciclo:
        null,

      atraso:
        null,

      factor:
        null,

      estado:
        'Aprendiendo'

    };

  }


  const ciclo =
    Math.max(
      1,
      WEBCL_mediana_(
        intervalos
      )
    );


  const atraso =
    Math.max(
      0,
      diasSinComprar -
      ciclo
    );


  const factor =
    ciclo
      ? diasSinComprar /
        ciclo
      : null;


  const limiteNormal =
    ciclo * 1.25 +
    2;


  const limiteAtencion =
    ciclo * 1.75 +
    3;


  const limiteDormido =
    Math.max(
      60,
      ciclo * 2.5
    );


  let estado;


  if (
    diasSinComprar <=
    limiteNormal
  ) {

    estado =
      'Al día';

  }

  else if (
    diasSinComprar <=
    limiteAtencion
  ) {

    estado =
      'Atención';

  }

  else if (
    diasSinComprar <=
    limiteDormido
  ) {

    estado =
      'Riesgo';

  }

  else {

    estado =
      'Dormido';

  }


  return {

    ultimaCompra:
      WEBCL_fechaKey_(
        ultima
      ),

    diasSinComprar:
      diasSinComprar,

    ciclo:
      WEBCL_redondear_(
        ciclo
      ),

    atraso:
      WEBCL_redondear_(
        atraso
      ),

    factor:
      WEBCL_redondear_(
        factor
      ),

    estado:
      estado

  };

}


// ============================================================================
// REACTIVACIÓN
// ============================================================================

function WEBCL_esReactivadoMes_(
  fechas,
  inicioMes
) {

  if (
    !Array.isArray(fechas) ||
    fechas.length < 2
  ) {

    return false;

  }


  const ordenadas =
    fechas.slice().sort(
      (a, b) =>
        a - b
    );


  const ultima =
    ordenadas[
      ordenadas.length - 1
    ];


  const anterior =
    ordenadas[
      ordenadas.length - 2
    ];


  if (
    ultima < inicioMes
  ) {

    return false;

  }


  return (
    WEBCL_diasEntre_(
      anterior,
      ultima
    ) >= 30
  );

}


// ============================================================================
// GIROS
// ============================================================================

function WEBCL_agruparGiros_(
  clientes
) {

  const mapa = {};


  clientes.forEach(
    cliente => {

      const giro =
        cliente.giro ||
        'Sin giro';


      if (
        !mapa[giro]
      ) {

        mapa[giro] = {

          giro:
            giro,

          clientes:
            0,

          riesgo:
            0

        };

      }


      mapa[giro].clientes++;


      if (
        [
          'Atención',
          'Riesgo',
          'Dormido'
        ].includes(
          cliente.estadoSalud
        )
      ) {

        mapa[giro].riesgo++;

      }

    }
  );


  return Object.values(
    mapa
  )

    .map(
      x => ({

        giro:
          x.giro,

        clientes:
          x.clientes,

        riesgo:
          x.riesgo,

        porcentajeRiesgo:
          x.clientes
            ? WEBCL_redondear_(
                x.riesgo /
                x.clientes *
                100
              )
            : 0

      })
    )

    .sort(
      (a, b) =>
        b.clientes -
        a.clientes
    );

}


// ============================================================================
// RIESGO POR VENDEDOR
// ============================================================================

function WEBCL_riesgoPorVendedor_(
  clientes
) {

  const mapa = {};


  clientes.forEach(
    cliente => {

      const vendedor =
        cliente.vendedor ||
        'Sin vendedor';


      if (
        !mapa[vendedor]
      ) {

        mapa[vendedor] = {

          vendedor:
            vendedor,

          clientes:
            0,

          riesgo:
            0

        };

      }


      mapa[vendedor].clientes++;


      if (
        [
          'Atención',
          'Riesgo',
          'Dormido'
        ].includes(
          cliente.estadoSalud
        )
      ) {

        mapa[vendedor].riesgo++;

      }

    }
  );


  return Object.values(
    mapa
  )

    .map(
      x => ({

        vendedor:
          x.vendedor,

        clientes:
          x.clientes,

        riesgo:
          x.riesgo,

        porcentaje:
          x.clientes
            ? WEBCL_redondear_(
                x.riesgo /
                x.clientes *
                100
              )
            : 0

      })
    )

    .sort(
      (a, b) =>
        b.riesgo -
        a.riesgo
    );

}


// ============================================================================
// INSIGHTS
// ============================================================================

function WEBCL_crearInsights_(
  activos,
  alDia,
  atencion,
  riesgo,
  dormidos,
  nuevos,
  reactivados,
  riesgoVendedores
) {

  const resultado = [];


  const carteraEvaluable =
    alDia.length +
    atencion.length +
    riesgo.length +
    dormidos.length;


  if (
    carteraEvaluable
  ) {

    const pctProblema =
      (
        atencion.length +
        riesgo.length +
        dormidos.length
      ) /
      carteraEvaluable *
      100;


    resultado.push({

      tipo:
        pctProblema >= 30
          ? 'advertencia'
          : 'positivo',

      titulo:
        'Salud de la cartera',

      texto:
        WEBCL_numeroTexto_(
          pctProblema
        ) +
        '% de los clientes evaluables está fuera de su ritmo normal de compra.'

    });

  }


  if (
    riesgo.length ||
    dormidos.length
  ) {

    resultado.push({

      tipo:
        'advertencia',

      titulo:
        'Clientes que requieren seguimiento',

      texto:
        riesgo.length +
        ' clientes están en riesgo y ' +
        dormidos.length +
        ' aparecen como dormidos según su historial individual.'

    });

  }


  if (
    nuevos.length
  ) {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Altas del mes',

      texto:
        'Este mes se han registrado ' +
        nuevos.length +
        ' clientes nuevos.'

    });

  }


  if (
    reactivados.length
  ) {

    resultado.push({

      tipo:
        'positivo',

      titulo:
        'Clientes recuperados',

      texto:
        reactivados.length +
        ' clientes regresaron este mes después de permanecer 30 días o más sin comprar.'

    });

  }


  if (
    riesgoVendedores.length &&
    riesgoVendedores[0].riesgo
  ) {

    const top =
      riesgoVendedores[0];


    resultado.push({

      tipo:
        'info',

      titulo:
        'Mayor cartera por atender',

      texto:
        top.vendedor +
        ' concentra ' +
        top.riesgo +
        ' clientes fuera de ritmo en su cartera.'

    });

  }


  return resultado.slice(
    0,
    5
  );

}


// ============================================================================
// CATÁLOGOS
// ============================================================================

function WEBCL_catalogos_(
  clientes
) {

  const vendedores = {};

  const giros =
    new Set();


  clientes.forEach(
    cliente => {

      if (
        cliente.vendedorId
      ) {

        vendedores[
          cliente.vendedorId
        ] =
          cliente.vendedor;

      }


      if (
        cliente.giro
      ) {

        giros.add(
          cliente.giro
        );

      }

    }
  );


  return {

    vendedores:
      Object.keys(
        vendedores
      )

        .map(
          id => ({

            valor:
              id,

            nombre:
              vendedores[id]

          })
        )

        .sort(
          (a, b) =>
            a.nombre.localeCompare(
              b.nombre,
              'es'
            )
        ),

    giros:
      Array.from(
        giros
      ).sort()

  };

}


// ============================================================================
// FILTROS
// ============================================================================

function WEBCL_cumpleFiltros_(
  cliente,
  filtros
) {

  if (
    filtros.vendedor &&
    String(
      cliente.vendedorId
    ) !==
    String(
      filtros.vendedor
    )
  ) {

    return false;

  }


  if (
    filtros.giro &&
    String(
      cliente.giro
    ) !==
    String(
      filtros.giro
    )
  ) {

    return false;

  }


  return true;

}


// ============================================================================
// COLUMNAS CLIENTES
// ============================================================================

function WEBCL_columnasClientes_(
  headers
) {

  return {

    telefono:
      WEBCL_findHeader_(
        headers,
        [
          'telefono',
          'teléfono'
        ]
      ),

    nombre:
      WEBCL_findHeader_(
        headers,
        [
          'nombre'
        ]
      ),

    estatus:
      WEBCL_findHeader_(
        headers,
        [
          'estatus'
        ]
      ),

    activo:
      WEBCL_findHeader_(
        headers,
        [
          'activo'
        ]
      ),

    vendedor:
      WEBCL_findHeader_(
        headers,
        [
          'id_vendedor'
        ]
      ),

    vendedorNombre:
      WEBCL_findHeader_(
        headers,
        [
          'vendedor_nombre'
        ]
      ),

    giro:
      WEBCL_findHeader_(
        headers,
        [
          'giro'
        ]
      ),

    colonia:
      WEBCL_findHeader_(
        headers,
        [
          'colonia'
        ]
      ),

    fechaRegistro:
      WEBCL_findHeader_(
        headers,
        [
          'fecha_registro',
          'fecha registro'
        ]
      )

  };

}


// ============================================================================
// CATÁLOGO DE GIROS
// ============================================================================

function WEBCL_getMapaGiros_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBCL_HOJA_GIROS
    );


  if (!sh) {
    return resultado;
  }


  const data =
    sh
      .getDataRange()
      .getValues();


  if (
    !data ||
    data.length < 2
  ) {
    return resultado;
  }


  const headers =
    data[0];


  const idCol =
    WEBCL_findHeader_(
      headers,
      [
        'idgiro',
        'id_giro',
        'codigo',
        'código'
      ]
    );


  const nombreCol =
    WEBCL_findHeader_(
      headers,
      [
        'giro',
        'nombre_giro',
        'nombre giro'
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
      WEBCL_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBCL_texto_(
        data[i][nombreCol]
      );


    if (id) {

      resultado[id] =
        nombre || id;


      resultado[
        id.toLowerCase()
      ] =
        nombre || id;

    }

  }


  return resultado;

}


// ============================================================================
// VENDEDORES
// ============================================================================

function WEBCL_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBCL_HOJA_PARAMETROS
    );


  if (!sh) {
    return resultado;
  }


  const data =
    sh
      .getDataRange()
      .getValues();


  if (
    !data ||
    data.length < 2
  ) {
    return resultado;
  }


  const headers =
    data[0];


  const idCol =
    WEBCL_findHeader_(
      headers,
      [
        'id_vendedor'
      ]
    );


  const nombreCol =
    WEBCL_findHeader_(
      headers,
      [
        'nombre'
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
      WEBCL_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBCL_texto_(
        data[i][nombreCol]
      );


    if (id) {

      resultado[
        id.toLowerCase()
      ] =
        nombre || id;

    }

  }


  return resultado;

}


// ============================================================================
// ACTIVO
// ============================================================================

function WEBCL_esActivo_(
  activo,
  estatus
) {

  const valorActivo =
    WEBCL_texto_(
      activo
    ).toLowerCase();


  if (
    valorActivo
  ) {

    if (
      [
        'false',
        '0',
        'no',
        'inactivo'
      ].includes(
        valorActivo
      )
    ) {

      return false;

    }


    if (
      [
        'true',
        '1',
        'si',
        'sí',
        'activo'
      ].includes(
        valorActivo
      )
    ) {

      return true;

    }

  }


  const valorEstatus =
    WEBCL_texto_(
      estatus
    ).toLowerCase();


  if (
    [
      'inactivo',
      'baja',
      'cancelado'
    ].includes(
      valorEstatus
    )
  ) {

    return false;

  }


  return true;

}


// ============================================================================
// CLIENT KEY
// ============================================================================

function WEBCL_clienteKey_(
  telefono,
  nombre
) {

  const telefonoKey =
    WEBCL_normalizarTelefono_(
      telefono
    );


  if (
    telefonoKey
  ) {

    return (
      'T:' +
      telefonoKey
    );

  }


  const nombreKey =
    WEBCL_normalizar_(
      nombre
    );


  return nombreKey
    ? 'N:' +
      nombreKey
    : '';

}


function WEBCL_normalizarTelefono_(
  valor
) {

  let numero =
    String(
      valor || ''
    )
      .replace(
        /\D/g,
        ''
      );


  if (
    numero.length > 10
  ) {

    numero =
      numero.slice(
        -10
      );

  }


  return numero;

}


// ============================================================================
// FECHAS
// ============================================================================

function WEBCL_hoyLocal_() {

  const key =
    Utilities.formatDate(
      new Date(),
      WEBCL_TZ,
      'yyyy-MM-dd'
    );


  return WEBCL_keyToDate_(
    key
  );

}


function WEBCL_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBCL_TZ,
    'yyyy-MM-dd'
  );

}


function WEBCL_keyToDate_(
  key
) {

  const p =
    String(
      key
    ).split(
      '-'
    );


  return new Date(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2]),
    12,
    0,
    0
  );

}


function WEBCL_diasEntre_(
  desde,
  hasta
) {

  const a =
    WEBCL_keyToDate_(
      WEBCL_fechaKey_(
        desde
      )
    );


  const b =
    WEBCL_keyToDate_(
      WEBCL_fechaKey_(
        hasta
      )
    );


  return Math.round(
    (
      b.getTime() -
      a.getTime()
    ) /
    86400000
  );

}


function WEBCL_mismoMes_(
  fecha,
  referencia
) {

  return (
    fecha.getFullYear() ===
      referencia.getFullYear() &&
    fecha.getMonth() ===
      referencia.getMonth()
  );

}


function WEBCL_parseFecha_(
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
      /^(\d{4})-(\d{1,2})-(\d{1,2})/
    );


  if (m) {

    return new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      12,
      0,
      0
    );

  }


  m =
    texto.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
    );


  if (m) {

    return new Date(
      Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1]),
      12,
      0,
      0
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


// ============================================================================
// MEDIANA
// ============================================================================

function WEBCL_mediana_(
  valores
) {

  if (
    !valores.length
  ) {
    return 0;
  }


  const ordenados =
    valores
      .slice()
      .sort(
        (a, b) =>
          a - b
      );


  const mitad =
    Math.floor(
      ordenados.length /
      2
    );


  if (
    ordenados.length %
    2
  ) {

    return ordenados[
      mitad
    ];

  }


  return (
    ordenados[
      mitad - 1
    ] +
    ordenados[
      mitad
    ]
  ) /
  2;

}


// ============================================================================
// HEADERS / VALORES
// ============================================================================

function WEBCL_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBCL_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const index =
      normalizados.indexOf(
        WEBCL_normalizar_(
          nombres[i]
        )
      );


    if (
      index !== -1
    ) {

      return index;

    }

  }


  return -1;

}


function WEBCL_normalizar_(
  valor
) {

  return String(
    valor || ''
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


function WEBCL_val_(
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


function WEBCL_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBCL_numero_(
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


function WEBCL_redondear_(
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


function WEBCL_numeroTexto_(
  numero
) {

  return Number(
    numero || 0
  ).toLocaleString(
    'es-MX',
    {
      maximumFractionDigits:
        1
    }
  );

}
