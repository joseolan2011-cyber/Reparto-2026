// ============================================================================
// WEB_Ventas.gs
// ATLAS REPARTO - ANÁLISIS DE VENTAS
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   VENTAS
//   PARAMETROS
//   GIROS
//
// PERIODOS:
//   - Hoy
//   - Semana anterior completa (lunes a sábado)
//   - Mes anterior completo
//   - Este mes
//   - Este año
//
// REGLA DE CIERRE:
//   - Antes de las 19:00, hoy se muestra pero no entra al comparativo.
//   - Desde las 19:00, hoy se considera cerrado.
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBV_TZ = 'America/Merida';
const WEBV_HORA_CORTE = 19;

const WEBV_HOJA_VENTAS = 'VENTAS';
const WEBV_HOJA_PARAMETROS = 'PARAMETROS';
const WEBV_HOJA_GIROS = 'GIROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getVentasDashboard(filtros) {

  try {

    filtros = filtros || {};


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const sh =
      ss.getSheetByName(
        WEBV_HOJA_VENTAS
      );


    if (!sh) {

      throw new Error(
        'No existe la hoja VENTAS.'
      );

    }


    const data =
      sh
        .getDataRange()
        .getValues();


    if (
      !data ||
      data.length < 2
    ) {

      return {

        ok: true,
        sinDatos: true,
        mensaje: 'La hoja VENTAS no contiene registros.'

      };

    }


    const mapaVendedores =
      WEBV_getMapaVendedores_(ss);


    const mapaGiros =
      WEBV_getMapaGiros_(ss);


    const columnas =
      WEBV_getColumnas_(
        data[0]
      );


    if (
      columnas.importe === -1
    ) {

      throw new Error(
        'No se encontró la columna Importe en VENTAS.'
      );

    }


    const registros = [];


    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const fila =
        data[i];


      let fecha =
        WEBV_parseFecha_(
          WEBV_val_(
            fila,
            columnas.fecha
          )
        );


      let fechaHora =
        WEBV_parseFecha_(
          WEBV_val_(
            fila,
            columnas.fechaHora
          )
        );


      if (!fecha) {
        fecha = fechaHora;
      }


      if (!fecha) {
        continue;
      }


      if (!fechaHora) {
        fechaHora = fecha;
      }


      const vendedorId =
        WEBV_texto_(
          WEBV_val_(
            fila,
            columnas.vendedor
          )
        );


      const giroCodigo =
        WEBV_texto_(
          WEBV_val_(
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
        WEBV_texto_(
          WEBV_val_(
            fila,
            columnas.telefono
          )
        );


      const cliente =
        WEBV_texto_(
          WEBV_val_(
            fila,
            columnas.cliente
          )
        );


      const idVenta =
        WEBV_texto_(
          WEBV_val_(
            fila,
            columnas.idVenta
          )
        ) ||
        '__ROW_' +
        (i + 1);


      registros.push({

        idVenta:
          idVenta,

        fecha:
          fecha,

        fechaHora:
          fechaHora,

        fechaKey:
          WEBV_fechaKey_(
            fecha
          ),

        telefono:
          telefono,

        cliente:
          cliente ||
          telefono ||
          'Sin nombre',

        clienteKey:
          WEBV_clienteKey_(
            telefono,
            cliente,
            i + 1
          ),

        ruta:
          WEBV_texto_(
            WEBV_val_(
              fila,
              columnas.ruta
            )
          ) ||
          'Sin ruta',

        vendedorId:
          vendedorId ||
          'Sin vendedor',

        vendedor:
          mapaVendedores[
            vendedorId.toLowerCase()
          ] ||
          vendedorId ||
          'Sin vendedor',

        giroCodigo:
          giroCodigo,

        giro:
          giroNombre,

        agua:
          WEBV_numero_(
            WEBV_val_(
              fila,
              columnas.agua
            )
          ),

        superIce:
          WEBV_numero_(
            WEBV_val_(
              fila,
              columnas.superIce
            )
          ),

        importe:
          WEBV_numero_(
            WEBV_val_(
              fila,
              columnas.importe
            )
          ),

        comision:
          WEBV_numero_(
            WEBV_val_(
              fila,
              columnas.comision
            )
          )

      });

    }


    const periodo =
      WEBV_getPeriodo_(
        filtros.periodo || 'mes',
        new Date()
      );


    const filtrados =
      registros.filter(
        r =>
          WEBV_cumpleFiltros_(
            r,
            filtros
          )
      );


    const visibles =
      filtrados.filter(
        r =>
          WEBV_entreFechas_(
            r.fechaKey,
            periodo.actualInicioKey,
            periodo.actualFinKey
          ) &&
          (
            !periodo.excluirDomingos ||
            r.fecha.getDay() !== 0
          )
      );


    let actualComparable = [];
    let anteriorComparable = [];


    if (
      periodo.comparacionHabilitada
    ) {

      actualComparable =
        filtrados.filter(
          r =>
            WEBV_entreFechas_(
              r.fechaKey,
              periodo.comparacionActualInicioKey,
              periodo.comparacionActualFinKey
            ) &&
            (
              !periodo.excluirDomingos ||
              r.fecha.getDay() !== 0
            )
        );


      anteriorComparable =
        filtrados.filter(
          r =>
            WEBV_entreFechas_(
              r.fechaKey,
              periodo.comparacionAnteriorInicioKey,
              periodo.comparacionAnteriorFinKey
            ) &&
            (
              !periodo.excluirDomingos ||
              r.fecha.getDay() !== 0
            )
        );

    }


    const totalVisible =
      WEBV_totales_(
        visibles
      );


    const totalActualComp =
      WEBV_totales_(
        actualComparable
      );


    const totalAnteriorComp =
      WEBV_totales_(
        anteriorComparable
      );


    const kpis = {

      venta:
        WEBV_kpi_(
          totalVisible.venta,
          totalActualComp.venta,
          totalAnteriorComp.venta,
          periodo.comparacionHabilitada
        ),

      tickets:
        WEBV_kpi_(
          totalVisible.tickets,
          totalActualComp.tickets,
          totalAnteriorComp.tickets,
          periodo.comparacionHabilitada
        ),

      ticketPromedio:
        WEBV_kpi_(
          totalVisible.ticketPromedio,
          totalActualComp.ticketPromedio,
          totalAnteriorComp.ticketPromedio,
          periodo.comparacionHabilitada
        ),

      clientes:
        WEBV_kpi_(
          totalVisible.clientes,
          totalActualComp.clientes,
          totalAnteriorComp.clientes,
          periodo.comparacionHabilitada
        ),

      agua:
        WEBV_kpi_(
          totalVisible.agua,
          totalActualComp.agua,
          totalAnteriorComp.agua,
          periodo.comparacionHabilitada
        ),

      superIce:
        WEBV_kpi_(
          totalVisible.superIce,
          totalActualComp.superIce,
          totalAnteriorComp.superIce,
          periodo.comparacionHabilitada
        ),

      comision:
        WEBV_kpi_(
          totalVisible.comision,
          totalActualComp.comision,
          totalAnteriorComp.comision,
          periodo.comparacionHabilitada
        ),

      unidadesTicket:
        WEBV_kpi_(
          totalVisible.unidadesTicket,
          totalActualComp.unidadesTicket,
          totalAnteriorComp.unidadesTicket,
          periodo.comparacionHabilitada
        )

    };


    const tendencia =
      WEBV_tendencia_(
        visibles,
        periodo
      );


    const mezcla = {

      agua:
        WEBV_redondear_(
          totalVisible.agua
        ),

      superIce:
        WEBV_redondear_(
          totalVisible.superIce
        )

    };


    const giros =
      WEBV_rankingGiros_(
        visibles,
        totalVisible.venta
      );


    const clientes =
      WEBV_rankingClientes_(
        visibles,
        totalVisible.venta
      );


    const insights =
      WEBV_insights_(
        periodo,
        kpis,
        tendencia,
        mezcla,
        giros,
        clientes,
        totalVisible
      );


    const catalogos =
      WEBV_catalogos_(
        registros
      );


    return {

      ok: true,
      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBV_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      kpis:
        kpis,

      tendencia:
        tendencia,

      mezcla:
        mezcla,

      giros:
        giros,

      clientes:
        clientes,

      insights:
        insights,

      catalogos:
        catalogos,

      meta: {

        registrosPeriodo:
          visibles.length,

        corteAlcanzado:
          periodo.corteAlcanzado

      }

    };


  } catch (error) {

    console.error(error);


    return {

      ok: false,
      mensaje: error.message,
      stack: error.stack

    };

  }

}


// ============================================================================
// TOTALES
// ============================================================================

function WEBV_totales_(
  registros
) {

  let venta = 0;
  let agua = 0;
  let superIce = 0;
  let comision = 0;


  const tickets =
    new Set();


  const clientes =
    new Set();


  registros.forEach(
    r => {

      venta +=
        r.importe;

      agua +=
        r.agua;

      superIce +=
        r.superIce;

      comision +=
        r.comision;


      tickets.add(
        r.idVenta
      );


      clientes.add(
        r.clienteKey
      );

    }
  );


  const numTickets =
    tickets.size;


  const unidades =
    agua +
    superIce;


  return {

    venta:
      WEBV_redondear_(
        venta
      ),

    agua:
      WEBV_redondear_(
        agua
      ),

    superIce:
      WEBV_redondear_(
        superIce
      ),

    comision:
      WEBV_redondear_(
        comision
      ),

    tickets:
      numTickets,

    clientes:
      clientes.size,

    ticketPromedio:
      numTickets
        ? WEBV_redondear_(
            venta /
            numTickets
          )
        : 0,

    unidadesTicket:
      numTickets
        ? WEBV_redondear_(
            unidades /
            numTickets
          )
        : 0

  };

}


// ============================================================================
// KPI
// ============================================================================

function WEBV_kpi_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  habilitada
) {

  return {

    actual:
      WEBV_redondear_(
        actualVisible
      ),

    variacion:
      habilitada
        ? WEBV_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBV_variacion_(
  actual,
  anterior
) {

  actual =
    Number(
      actual || 0
    );


  anterior =
    Number(
      anterior || 0
    );


  if (
    anterior === 0 &&
    actual === 0
  ) {
    return 0;
  }


  if (
    anterior === 0
  ) {
    return null;
  }


  return WEBV_redondear_(
    (
      (
        actual -
        anterior
      ) /
      Math.abs(
        anterior
      )
    ) *
    100
  );

}


// ============================================================================
// TENDENCIA
// ============================================================================

function WEBV_tendencia_(
  registros,
  periodo
) {

  const mapa = {};


  registros.forEach(
    r => {

      let key;
      let label;


      if (
        periodo.agrupacion ===
        'hora'
      ) {

        const hora =
          Utilities.formatDate(
            r.fechaHora,
            WEBV_TZ,
            'HH'
          );


        key =
          hora;


        label =
          hora +
          ':00';

      }

      else if (
        periodo.agrupacion ===
        'mes'
      ) {

        key =
          r.fechaKey.substring(
            0,
            7
          );


        label =
          WEBV_nombreMesCorto_(
            Number(
              key.substring(
                5,
                7
              )
            ) - 1
          );

      }

      else {

        key =
          r.fechaKey;


        const p =
          key.split(
            '-'
          );


        label =
          p[2] +
          '/' +
          p[1];

      }


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          key:
            key,

          label:
            label,

          venta:
            0,

          agua:
            0,

          superIce:
            0,

          tickets:
            new Set()

        };

      }


      mapa[key].venta +=
        r.importe;


      mapa[key].agua +=
        r.agua;


      mapa[key].superIce +=
        r.superIce;


      mapa[key]
        .tickets
        .add(
          r.idVenta
        );

    }
  );


  return Object
    .keys(mapa)
    .sort()
    .map(
      key => ({

        key:
          key,

        label:
          mapa[key].label,

        venta:
          WEBV_redondear_(
            mapa[key].venta
          ),

        agua:
          WEBV_redondear_(
            mapa[key].agua
          ),

        superIce:
          WEBV_redondear_(
            mapa[key].superIce
          ),

        tickets:
          mapa[key]
            .tickets
            .size

      })
    );

}


// ============================================================================
// RANKING DE GIROS
// ============================================================================

function WEBV_rankingGiros_(
  registros,
  ventaTotal
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        r.giro ||
        'Sin giro';


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          giro:
            key,

          venta:
            0,

          agua:
            0,

          superIce:
            0,

          tickets:
            new Set(),

          clientes:
            new Set()

        };

      }


      mapa[key].venta +=
        r.importe;


      mapa[key].agua +=
        r.agua;


      mapa[key].superIce +=
        r.superIce;


      mapa[key]
        .tickets
        .add(
          r.idVenta
        );


      mapa[key]
        .clientes
        .add(
          r.clienteKey
        );

    }
  );


  return Object
    .values(mapa)
    .map(
      x => ({

        giro:
          x.giro,

        venta:
          WEBV_redondear_(
            x.venta
          ),

        agua:
          WEBV_redondear_(
            x.agua
          ),

        superIce:
          WEBV_redondear_(
            x.superIce
          ),

        tickets:
          x.tickets.size,

        clientes:
          x.clientes.size,

        ticketPromedio:
          x.tickets.size
            ? WEBV_redondear_(
                x.venta /
                x.tickets.size
              )
            : 0,

        participacion:
          ventaTotal
            ? WEBV_redondear_(
                x.venta /
                ventaTotal *
                100
              )
            : 0

      })
    )
    .sort(
      (a, b) =>
        b.venta -
        a.venta
    );

}


// ============================================================================
// RANKING DE CLIENTES
// ============================================================================

function WEBV_rankingClientes_(
  registros,
  ventaTotal
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        r.clienteKey;


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          key:
            key,

          nombre:
            r.cliente,

          telefono:
            r.telefono,

          vendedor:
            r.vendedor,

          giro:
            r.giro,

          venta:
            0,

          agua:
            0,

          superIce:
            0,

          tickets:
            new Set(),

          ultimaCompra:
            r.fechaKey

        };

      }


      mapa[key].venta +=
        r.importe;


      mapa[key].agua +=
        r.agua;


      mapa[key].superIce +=
        r.superIce;


      mapa[key]
        .tickets
        .add(
          r.idVenta
        );


      if (
        r.fechaKey >
        mapa[key].ultimaCompra
      ) {

        mapa[key].ultimaCompra =
          r.fechaKey;

      }

    }
  );


  return Object
    .values(mapa)
    .map(
      x => ({

        key:
          x.key,

        nombre:
          x.nombre,

        telefono:
          x.telefono,

        vendedor:
          x.vendedor,

        giro:
          x.giro,

        venta:
          WEBV_redondear_(
            x.venta
          ),

        agua:
          WEBV_redondear_(
            x.agua
          ),

        superIce:
          WEBV_redondear_(
            x.superIce
          ),

        tickets:
          x.tickets.size,

        ticketPromedio:
          x.tickets.size
            ? WEBV_redondear_(
                x.venta /
                x.tickets.size
              )
            : 0,

        participacion:
          ventaTotal
            ? WEBV_redondear_(
                x.venta /
                ventaTotal *
                100
              )
            : 0,

        ultimaCompra:
          x.ultimaCompra

      })
    )
    .sort(
      (a, b) =>
        b.venta -
        a.venta
    )
    .slice(
      0,
      50
    );

}


// ============================================================================
// INSIGHTS
// ============================================================================

function WEBV_insights_(
  periodo,
  kpis,
  tendencia,
  mezcla,
  giros,
  clientes,
  totales
) {

  const resultado = [];


  if (
    periodo.comparacionHabilitada
  ) {

    const variacion =
      kpis.venta.variacion;


    if (
      variacion !== null
    ) {

      if (
        variacion >= 5
      ) {

        resultado.push({

          tipo:
            'positivo',

          titulo:
            'Venta por encima del comparativo',

          texto:
            'La venta del periodo comparable creció ' +
            WEBV_pctTexto_(
              variacion
            ) +
            ' frente a ' +
            periodo.etiquetaAnterior +
            '.'

        });

      }

      else if (
        variacion <= -5
      ) {

        resultado.push({

          tipo:
            'advertencia',

          titulo:
            'Venta por debajo del comparativo',

          texto:
            'La venta del periodo comparable está ' +
            WEBV_pctTexto_(
              Math.abs(
                variacion
              )
            ) +
            ' por debajo de ' +
            periodo.etiquetaAnterior +
            '.'

        });

      }

    }

  }

  else {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Jornada abierta',

      texto:
        'La venta de hoy se muestra en tiempo real. El comparativo diario se habilita después de las 7:00 p. m.'

    });

  }


  if (
    tendencia.length
  ) {

    const mejor =
      tendencia
        .slice()
        .sort(
          (a, b) =>
            b.venta -
            a.venta
        )[0];


    resultado.push({

      tipo:
        'positivo',

      titulo:
        'Pico de venta del periodo',

      texto:
        mejor.label +
        ' es el punto con mayor venta: ' +
        WEBV_monedaTexto_(
          mejor.venta
        ) +
        '.'

    });

  }


  if (
    giros.length
  ) {

    const top =
      giros[0];


    resultado.push({

      tipo:
        'info',

      titulo:
        'Giro con mayor peso',

      texto:
        top.giro +
        ' representa ' +
        WEBV_pctTexto_(
          top.participacion
        ) +
        ' de la venta del periodo.'

    });

  }


  if (
    clientes.length &&
    totales.venta
  ) {

    const top5 =
      clientes
        .slice(
          0,
          5
        )
        .reduce(
          (suma, x) =>
            suma +
            x.venta,
          0
        );


    const concentracion =
      top5 /
      totales.venta *
      100;


    resultado.push({

      tipo:
        concentracion >= 40
          ? 'advertencia'
          : 'info',

      titulo:
        'Concentración en clientes',

      texto:
        'Los 5 clientes con mayor compra concentran ' +
        WEBV_pctTexto_(
          concentracion
        ) +
        ' de la venta seleccionada.'

    });

  }


  const unidades =
    mezcla.agua +
    mezcla.superIce;


  if (
    unidades > 0
  ) {

    const pctAgua =
      mezcla.agua /
      unidades *
      100;


    resultado.push({

      tipo:
        'info',

      titulo:
        'Mezcla de unidades',

      texto:
        'Agua 20 L representa ' +
        WEBV_pctTexto_(
          pctAgua
        ) +
        ' de las unidades registradas en el periodo.'

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

function WEBV_catalogos_(
  registros
) {

  const rutas =
    new Set();


  const giros =
    new Set();


  const vendedores = {};


  registros.forEach(
    r => {

      if (
        r.ruta
      ) {
        rutas.add(
          r.ruta
        );
      }


      if (
        r.giro
      ) {
        giros.add(
          r.giro
        );
      }


      if (
        r.vendedorId
      ) {

        vendedores[
          r.vendedorId
        ] =
          r.vendedor;

      }

    }
  );


  return {

    rutas:
      Array
        .from(rutas)
        .sort(
          (a, b) =>
            String(a).localeCompare(
              String(b),
              'es',
              {
                numeric: true
              }
            )
        ),

    giros:
      Array
        .from(giros)
        .sort(
          (a, b) =>
            String(a).localeCompare(
              String(b),
              'es'
            )
        ),

    vendedores:
      Object
        .keys(vendedores)
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
        )

  };

}


// ============================================================================
// FILTROS
// ============================================================================

function WEBV_cumpleFiltros_(
  registro,
  filtros
) {

  if (
    filtros.ruta &&
    String(
      registro.ruta
    ) !==
    String(
      filtros.ruta
    )
  ) {

    return false;

  }


  if (
    filtros.vendedor &&
    String(
      registro.vendedorId
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
      registro.giro
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
// PERIODOS
// ============================================================================

function WEBV_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo || 'mes'
    ).toLowerCase();


  const hoy =
    WEBV_hoyLocal_(
      ahora
    );


  const corte =
    WEBV_corteAlcanzado_(
      ahora
    );


  let actualInicio;
  let actualFin;

  let compActualInicio = null;
  let compActualFin = null;

  let compAnteriorInicio = null;
  let compAnteriorFin = null;

  let comparacionHabilitada = true;

  let etiqueta = '';
  let etiquetaAnterior = '';

  let agrupacion = 'dia';
  let excluirDomingos = false;


  if (
    tipo === 'hoy'
  ) {

    actualInicio =
      hoy;

    actualFin =
      hoy;

    etiqueta =
      'Hoy';

    agrupacion =
      'hora';


    if (
      corte
    ) {

      compActualInicio =
        hoy;

      compActualFin =
        hoy;


      const anterior =
        WEBV_diaOperativoAnterior_(
          hoy
        );


      compAnteriorInicio =
        anterior;

      compAnteriorFin =
        anterior;


      etiquetaAnterior =
        'último día operativo';

    }

    else {

      comparacionHabilitada =
        false;

      etiquetaAnterior =
        'Comparación disponible después de las 7:00 p. m.';

    }

  }

  else if (
    tipo === 'semana_anterior'
  ) {

    excluirDomingos =
      true;


    const lunesActual =
      WEBV_inicioSemanaLunes_(
        hoy
      );


    actualInicio =
      WEBV_addDias_(
        lunesActual,
        -7
      );


    actualFin =
      WEBV_addDias_(
        actualInicio,
        5
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBV_addDias_(
        actualInicio,
        -7
      );


    compAnteriorFin =
      WEBV_addDias_(
        compAnteriorInicio,
        5
      );


    etiqueta =
      'Semana anterior';


    etiquetaAnterior =
      'semana previa';

  }

  else if (
    tipo === 'mes_anterior'
  ) {

    const primerMesActual =
      WEBV_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualInicio =
      WEBV_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 1,
        1
      );


    actualFin =
      WEBV_addDias_(
        primerMesActual,
        -1
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBV_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 2,
        1
      );


    compAnteriorFin =
      WEBV_addDias_(
        actualInicio,
        -1
      );


    etiqueta =
      WEBV_nombreMes_(
        actualInicio.getMonth()
      ) +
      ' ' +
      actualInicio.getFullYear();


    etiquetaAnterior =
      WEBV_nombreMes_(
        compAnteriorInicio.getMonth()
      ) +
      ' ' +
      compAnteriorInicio.getFullYear();

  }

  else if (
    tipo === 'anio'
  ) {

    actualInicio =
      WEBV_fechaLocal_(
        hoy.getFullYear(),
        0,
        1
      );


    actualFin =
      hoy;


    etiqueta =
      'Año ' +
      hoy.getFullYear();


    agrupacion =
      'mes';


    const ultimoComparable =
      corte
        ? hoy
        : WEBV_addDias_(
            hoy,
            -1
          );


    if (
      ultimoComparable <
      actualInicio
    ) {

      comparacionHabilitada =
        false;

      etiquetaAnterior =
        'mismo periodo del año anterior';

    }

    else {

      compActualInicio =
        actualInicio;

      compActualFin =
        ultimoComparable;


      compAnteriorInicio =
        WEBV_fechaLocal_(
          hoy.getFullYear() - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBV_fechaClampeada_(
          hoy.getFullYear() - 1,
          ultimoComparable.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        'mismo periodo de ' +
        (hoy.getFullYear() - 1) +
        (
          corte
            ? ''
            : ' · hoy excluido'
        );

    }

  }

  else {

    tipo =
      'mes';


    actualInicio =
      WEBV_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualFin =
      hoy;


    etiqueta =
      WEBV_nombreMes_(
        hoy.getMonth()
      ) +
      ' ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBV_addDias_(
            hoy,
            -1
          );


    if (
      ultimoComparable <
      actualInicio
    ) {

      comparacionHabilitada =
        false;

      etiquetaAnterior =
        'Mes anterior · disponible después de las 7:00 p. m.';

    }

    else {

      compActualInicio =
        actualInicio;

      compActualFin =
        ultimoComparable;


      const mesAnterior =
        WEBV_fechaLocal_(
          hoy.getFullYear(),
          hoy.getMonth() - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBV_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        WEBV_nombreMes_(
          mesAnterior.getMonth()
        ) +
        ' ' +
        mesAnterior.getFullYear() +
        ' al ' +
        compAnteriorFin.getDate() +
        (
          corte
            ? ''
            : ' · hoy excluido'
        );

    }

  }


  return {

    tipo:
      tipo,

    etiqueta:
      etiqueta,

    etiquetaAnterior:
      etiquetaAnterior,

    agrupacion:
      agrupacion,

    excluirDomingos:
      excluirDomingos,

    corteAlcanzado:
      corte,

    comparacionHabilitada:
      comparacionHabilitada,

    actualInicioKey:
      WEBV_fechaKey_(
        actualInicio
      ),

    actualFinKey:
      WEBV_fechaKey_(
        actualFin
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBV_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBV_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBV_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBV_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// COLUMNAS
// ============================================================================

function WEBV_getColumnas_(
  headers
) {

  return {

    idVenta:
      WEBV_findHeader_(
        headers,
        [
          'id_venta',
          'id venta'
        ]
      ),

    telefono:
      WEBV_findHeader_(
        headers,
        [
          'telefono_cliente',
          'telefono cliente'
        ]
      ),

    cliente:
      WEBV_findHeader_(
        headers,
        [
          'cliente'
        ]
      ),

    ruta:
      WEBV_findHeader_(
        headers,
        [
          'ruta'
        ]
      ),

    vendedor:
      WEBV_findHeader_(
        headers,
        [
          'id_vendedor'
        ]
      ),

    fechaHora:
      WEBV_findHeader_(
        headers,
        [
          'fecha_hora',
          'fecha hora'
        ]
      ),

    fecha:
      WEBV_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    agua:
      WEBV_findHeader_(
        headers,
        [
          'agua 20 litros',
          'agua_20_litros'
        ]
      ),

    superIce:
      WEBV_findHeader_(
        headers,
        [
          'superice',
          'super ice'
        ]
      ),

    importe:
      WEBV_findHeader_(
        headers,
        [
          'importe'
        ]
      ),

    giro:
      WEBV_findHeader_(
        headers,
        [
          'giro'
        ]
      ),

    comision:
      WEBV_findHeader_(
        headers,
        [
          'comision',
          'comisión'
        ]
      )

  };

}


// ============================================================================
// PARAMETROS
// ============================================================================

function WEBV_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBV_HOJA_PARAMETROS
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


  const idCol =
    WEBV_findHeader_(
      data[0],
      [
        'id_vendedor'
      ]
    );


  const nombreCol =
    WEBV_findHeader_(
      data[0],
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
      WEBV_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBV_texto_(
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
// GIROS
// ============================================================================

function WEBV_getMapaGiros_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBV_HOJA_GIROS
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


  const idCol =
    WEBV_findHeader_(
      data[0],
      [
        'idgiro',
        'id_giro'
      ]
    );


  const nombreCol =
    WEBV_findHeader_(
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
      WEBV_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBV_texto_(
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
// FECHAS
// ============================================================================

function WEBV_hoyLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      WEBV_TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


  return WEBV_fechaLocal_(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2])
  );

}


function WEBV_corteAlcanzado_(
  fecha
) {

  return Number(
    Utilities.formatDate(
      fecha,
      WEBV_TZ,
      'H'
    )
  ) >= WEBV_HORA_CORTE;

}


function WEBV_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  return WEBV_addDias_(
    fecha,
    dia === 0
      ? -6
      : 1 - dia
  );

}


function WEBV_diaOperativoAnterior_(
  fecha
) {

  let r =
    WEBV_addDias_(
      fecha,
      -1
    );


  while (
    r.getDay() === 0
  ) {

    r =
      WEBV_addDias_(
        r,
        -1
      );

  }


  return r;

}


function WEBV_fechaLocal_(
  year,
  month,
  day
) {

  return new Date(
    year,
    month,
    day,
    12,
    0,
    0,
    0
  );

}


function WEBV_addDias_(
  fecha,
  dias
) {

  const r =
    new Date(
      fecha.getTime()
    );


  r.setDate(
    r.getDate() +
    dias
  );


  return WEBV_fechaLocal_(
    r.getFullYear(),
    r.getMonth(),
    r.getDate()
  );

}


function WEBV_fechaClampeada_(
  year,
  month,
  day
) {

  const ultimo =
    new Date(
      year,
      month + 1,
      0
    ).getDate();


  return WEBV_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimo
    )
  );

}


function WEBV_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBV_TZ,
    'yyyy-MM-dd'
  );

}


function WEBV_entreFechas_(
  key,
  inicio,
  fin
) {

  return (
    inicio &&
    fin &&
    key >= inicio &&
    key <= fin
  );

}


function WEBV_parseFecha_(
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


// ============================================================================
// UTILIDADES
// ============================================================================

function WEBV_clienteKey_(
  telefono,
  nombre,
  fallback
) {

  let tel =
    String(
      telefono || ''
    ).replace(
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
    return 'T:' + tel;
  }


  const nom =
    WEBV_normalizar_(
      nombre
    );


  return nom
    ? 'N:' + nom
    : 'R:' + fallback;

}


function WEBV_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBV_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        WEBV_normalizar_(
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


function WEBV_normalizar_(
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


function WEBV_val_(
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


function WEBV_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBV_numero_(
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


  const n =
    Number(
      limpio
    );


  return isFinite(n)
    ? n
    : 0;

}


function WEBV_redondear_(
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


function WEBV_nombreMes_(
  mes
) {

  const meses = [
    'Enero',
    'Febrero',
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre'
  ];


  return meses[
    mes
  ] || '';

}


function WEBV_nombreMesCorto_(
  mes
) {

  const meses = [
    'Ene',
    'Feb',
    'Mar',
    'Abr',
    'May',
    'Jun',
    'Jul',
    'Ago',
    'Sep',
    'Oct',
    'Nov',
    'Dic'
  ];


  return meses[
    mes
  ] || '';

}


function WEBV_pctTexto_(
  valor
) {

  return Number(
    valor || 0
  ).toLocaleString(
    'es-MX',
    {
      maximumFractionDigits: 1
    }
  ) + '%';

}


function WEBV_monedaTexto_(
  valor
) {

  return '$' +
    Number(
      valor || 0
    ).toLocaleString(
      'es-MX',
      {
        maximumFractionDigits: 0
      }
    );

}
