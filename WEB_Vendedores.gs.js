// ============================================================================
// WEB_Vendedores.gs
// ATLAS REPARTO - ANÁLISIS DE VENDEDORES
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   VENTAS
//   PARAMETROS
//   GIROS
//
// Analiza:
//   - venta por vendedor
//   - crecimiento / caída
//   - clientes compradores
//   - tickets
//   - ticket promedio
//   - Agua / SuperIce
//   - comisiones
//   - participación de venta
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBVD_TZ = 'America/Merida';
const WEBVD_HORA_CORTE = 19;

const WEBVD_HOJA_VENTAS = 'VENTAS';
const WEBVD_HOJA_PARAMETROS = 'PARAMETROS';
const WEBVD_HOJA_GIROS = 'GIROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getVendedoresDashboard(filtros) {

  try {

    filtros = filtros || {};


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const shVentas =
      ss.getSheetByName(
        WEBVD_HOJA_VENTAS
      );


    if (!shVentas) {

      throw new Error(
        'No existe la hoja VENTAS.'
      );

    }


    const data =
      shVentas
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


    const vendedores =
      WEBVD_getMapaVendedores_(ss);


    const giros =
      WEBVD_getMapaGiros_(ss);


    const columnas =
      WEBVD_getColumnas_(
        data[0]
      );


    const registros = [];


    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const fila =
        data[i];


      let fecha =
        WEBVD_parseFecha_(
          WEBVD_val_(
            fila,
            columnas.fecha
          )
        );


      let fechaHora =
        WEBVD_parseFecha_(
          WEBVD_val_(
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
        WEBVD_texto_(
          WEBVD_val_(
            fila,
            columnas.vendedor
          )
        ) ||
        'Sin vendedor';


      const vendedorNombre =
        vendedores[
          vendedorId.toLowerCase()
        ] ||
        vendedorId;


      const giroCodigo =
        WEBVD_texto_(
          WEBVD_val_(
            fila,
            columnas.giro
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
        WEBVD_texto_(
          WEBVD_val_(
            fila,
            columnas.telefono
          )
        );


      const cliente =
        WEBVD_texto_(
          WEBVD_val_(
            fila,
            columnas.cliente
          )
        );


      const idVenta =
        WEBVD_texto_(
          WEBVD_val_(
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
          WEBVD_fechaKey_(
            fecha
          ),

        vendedorId:
          vendedorId,

        vendedor:
          vendedorNombre,

        ruta:
          WEBVD_texto_(
            WEBVD_val_(
              fila,
              columnas.ruta
            )
          ) ||
          'Sin ruta',

        giro:
          giroNombre,

        clienteKey:
          WEBVD_clienteKey_(
            telefono,
            cliente,
            i + 1
          ),

        cliente:
          cliente ||
          telefono ||
          'Sin nombre',

        agua:
          WEBVD_numero_(
            WEBVD_val_(
              fila,
              columnas.agua
            )
          ),

        superIce:
          WEBVD_numero_(
            WEBVD_val_(
              fila,
              columnas.superIce
            )
          ),

        importe:
          WEBVD_numero_(
            WEBVD_val_(
              fila,
              columnas.importe
            )
          ),

        comision:
          WEBVD_numero_(
            WEBVD_val_(
              fila,
              columnas.comision
            )
          )

      });

    }


    const periodo =
      WEBVD_getPeriodo_(
        filtros.periodo || 'mes',
        new Date()
      );


    const filtrados =
      registros.filter(
        r =>
          WEBVD_cumpleFiltros_(
            r,
            filtros
          )
      );


    const actuales =
      filtrados.filter(
        r =>
          WEBVD_entreFechas_(
            r.fechaKey,
            periodo.actualInicioKey,
            periodo.actualFinKey
          ) &&
          (
            !periodo.excluirDomingos ||
            r.fecha.getDay() !== 0
          )
      );


    let actualesComp = [];
    let anterioresComp = [];


    if (
      periodo.comparacionHabilitada
    ) {

      actualesComp =
        filtrados.filter(
          r =>
            WEBVD_entreFechas_(
              r.fechaKey,
              periodo.comparacionActualInicioKey,
              periodo.comparacionActualFinKey
            ) &&
            (
              !periodo.excluirDomingos ||
              r.fecha.getDay() !== 0
            )
        );


      anterioresComp =
        filtrados.filter(
          r =>
            WEBVD_entreFechas_(
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


    const tabla =
      WEBVD_crearTablaVendedores_(
        actuales,
        actualesComp,
        anterioresComp,
        periodo.comparacionHabilitada
      );


    const totalesActuales =
      WEBVD_totales_(
        actuales
      );


    const totalesActualesComp =
      WEBVD_totales_(
        actualesComp
      );


    const totalesAnterioresComp =
      WEBVD_totales_(
        anterioresComp
      );


    const vendedoresConVenta =
      tabla.filter(
        x =>
          x.venta > 0
      ).length;


    const promedioVenta =
      vendedoresConVenta
        ? WEBVD_redondear_(
            totalesActuales.venta /
            vendedoresConVenta
          )
        : 0;


    const kpis = {

      vendedores:
        vendedoresConVenta,

      venta:
        WEBVD_kpi_(
          totalesActuales.venta,
          totalesActualesComp.venta,
          totalesAnterioresComp.venta,
          periodo.comparacionHabilitada
        ),

      ventaPromedio:
        promedioVenta,

      tickets:
        WEBVD_kpi_(
          totalesActuales.tickets,
          totalesActualesComp.tickets,
          totalesAnterioresComp.tickets,
          periodo.comparacionHabilitada
        ),

      clientes:
        WEBVD_kpi_(
          totalesActuales.clientes,
          totalesActualesComp.clientes,
          totalesAnterioresComp.clientes,
          periodo.comparacionHabilitada
        ),

      ticketPromedio:
        WEBVD_kpi_(
          totalesActuales.ticketPromedio,
          totalesActualesComp.ticketPromedio,
          totalesAnterioresComp.ticketPromedio,
          periodo.comparacionHabilitada
        ),

      agua:
        WEBVD_kpi_(
          totalesActuales.agua,
          totalesActualesComp.agua,
          totalesAnterioresComp.agua,
          periodo.comparacionHabilitada
        ),

      comision:
        WEBVD_kpi_(
          totalesActuales.comision,
          totalesActualesComp.comision,
          totalesAnterioresComp.comision,
          periodo.comparacionHabilitada
        )

    };


    const grafica =
      tabla
        .slice(
          0,
          10
        )
        .map(
          x => ({

            vendedor:
              x.vendedor,

            venta:
              x.venta,

            clientes:
              x.clientes,

            tickets:
              x.tickets

          })
        );


    const insights =
      WEBVD_insights_(
        tabla,
        kpis,
        periodo,
        totalesActuales
      );


    const catalogos =
      WEBVD_catalogos_(
        registros
      );


    return {

      ok: true,
      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBVD_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      kpis:
        kpis,

      vendedores:
        tabla,

      grafica:
        grafica,

      insights:
        insights,

      catalogos:
        catalogos

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
// TABLA POR VENDEDOR
// ============================================================================

function WEBVD_crearTablaVendedores_(
  actuales,
  actualesComp,
  anterioresComp,
  comparacionHabilitada
) {

  const mapActual =
    WEBVD_agruparPorVendedor_(
      actuales
    );


  const mapActualComp =
    WEBVD_agruparPorVendedor_(
      actualesComp
    );


  const mapAnteriorComp =
    WEBVD_agruparPorVendedor_(
      anterioresComp
    );


  const ventaTotal =
    Object
      .values(
        mapActual
      )
      .reduce(
        (suma, x) =>
          suma +
          x.venta,
        0
      );


  const resultado = [];


  Object
    .keys(
      mapActual
    )
    .forEach(
      key => {

        const actual =
          mapActual[key];


        const actualComp =
          mapActualComp[key] || {
            venta: 0
          };


        const anteriorComp =
          mapAnteriorComp[key] || {
            venta: 0
          };


        resultado.push({

          vendedorId:
            key,

          vendedor:
            actual.vendedor,

          venta:
            WEBVD_redondear_(
              actual.venta
            ),

          agua:
            WEBVD_redondear_(
              actual.agua
            ),

          superIce:
            WEBVD_redondear_(
              actual.superIce
            ),

          comision:
            WEBVD_redondear_(
              actual.comision
            ),

          tickets:
            actual
              .tickets
              .size,

          clientes:
            actual
              .clientes
              .size,

          ticketPromedio:
            actual.tickets.size
              ? WEBVD_redondear_(
                  actual.venta /
                  actual.tickets.size
                )
              : 0,

          ventaCliente:
            actual.clientes.size
              ? WEBVD_redondear_(
                  actual.venta /
                  actual.clientes.size
                )
              : 0,

          participacion:
            ventaTotal
              ? WEBVD_redondear_(
                  actual.venta /
                  ventaTotal *
                  100
                )
              : 0,

          variacion:
            comparacionHabilitada
              ? WEBVD_variacion_(
                  actualComp.venta,
                  anteriorComp.venta
                )
              : null

        });

      }
    );


  resultado.sort(
    (a, b) =>
      b.venta -
      a.venta
  );


  return resultado;

}


function WEBVD_agruparPorVendedor_(
  registros
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        r.vendedorId;


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          vendedor:
            r.vendedor,

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


      mapa[key].venta +=
        r.importe;


      mapa[key].agua +=
        r.agua;


      mapa[key].superIce +=
        r.superIce;


      mapa[key].comision +=
        r.comision;


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


  return mapa;

}


// ============================================================================
// TOTALES
// ============================================================================

function WEBVD_totales_(
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


  return {

    venta:
      WEBVD_redondear_(
        venta
      ),

    agua:
      WEBVD_redondear_(
        agua
      ),

    superIce:
      WEBVD_redondear_(
        superIce
      ),

    comision:
      WEBVD_redondear_(
        comision
      ),

    tickets:
      tickets.size,

    clientes:
      clientes.size,

    ticketPromedio:
      tickets.size
        ? WEBVD_redondear_(
            venta /
            tickets.size
          )
        : 0

  };

}


// ============================================================================
// KPI
// ============================================================================

function WEBVD_kpi_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  habilitada
) {

  return {

    actual:
      WEBVD_redondear_(
        actualVisible
      ),

    variacion:
      habilitada
        ? WEBVD_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBVD_variacion_(
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


  return WEBVD_redondear_(
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
// INSIGHTS
// ============================================================================

function WEBVD_insights_(
  tabla,
  kpis,
  periodo,
  totales
) {

  const resultado = [];


  if (
    tabla.length
  ) {

    const top =
      tabla[0];


    resultado.push({

      tipo:
        'positivo',

      titulo:
        'Mayor aportación',

      texto:
        top.vendedor +
        ' genera ' +
        WEBVD_pctTexto_(
          top.participacion
        ) +
        ' de la venta del periodo con ' +
        WEBVD_monedaTexto_(
          top.venta
        ) +
        '.'

    });

  }


  if (
    periodo.comparacionHabilitada
  ) {

    const crecimientos =
      tabla
        .filter(
          x =>
            x.variacion !== null
        )
        .slice()
        .sort(
          (a, b) =>
            b.variacion -
            a.variacion
        );


    if (
      crecimientos.length &&
      crecimientos[0].variacion > 5
    ) {

      const top =
        crecimientos[0];


      resultado.push({

        tipo:
          'positivo',

        titulo:
          'Mayor crecimiento',

        texto:
          top.vendedor +
          ' presenta un crecimiento de ' +
          WEBVD_pctTexto_(
            top.variacion
          ) +
          ' frente al periodo comparable.'

      });

    }


    const caidas =
      tabla
        .filter(
          x =>
            x.variacion !== null &&
            x.variacion < 0
        )
        .slice()
        .sort(
          (a, b) =>
            a.variacion -
            b.variacion
        );


    if (
      caidas.length &&
      caidas[0].variacion <= -10
    ) {

      const bajo =
        caidas[0];


      resultado.push({

        tipo:
          'advertencia',

        titulo:
          'Caída a revisar',

        texto:
          bajo.vendedor +
          ' está ' +
          WEBVD_pctTexto_(
            Math.abs(
              bajo.variacion
            )
          ) +
          ' por debajo de su periodo comparable.'

      });

    }

  }

  else {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Jornada todavía abierta',

      texto:
        'Atlas mostrará la comparación diaria de vendedores después de las 7:00 p. m.'

    });

  }


  if (
    tabla.length >= 2 &&
    totales.venta
  ) {

    const top2 =
      tabla
        .slice(
          0,
          2
        )
        .reduce(
          (suma, x) =>
            suma +
            x.venta,
          0
        );


    const pct =
      top2 /
      totales.venta *
      100;


    resultado.push({

      tipo:
        pct >= 70
          ? 'advertencia'
          : 'info',

      titulo:
        'Concentración del equipo',

      texto:
        'Los 2 vendedores con mayor venta concentran ' +
        WEBVD_pctTexto_(
          pct
        ) +
        ' del total del periodo.'

    });

  }


  if (
    tabla.length
  ) {

    const mejorTicket =
      tabla
        .slice()
        .sort(
          (a, b) =>
            b.ticketPromedio -
            a.ticketPromedio
        )[0];


    resultado.push({

      tipo:
        'info',

      titulo:
        'Mayor ticket promedio',

      texto:
        mejorTicket.vendedor +
        ' tiene el ticket promedio más alto del periodo: ' +
        WEBVD_monedaTexto_(
          mejorTicket.ticketPromedio
        ) +
        '.'

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

function WEBVD_catalogos_(
  registros
) {

  const rutas =
    new Set();


  const giros =
    new Set();


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
        )

  };

}


// ============================================================================
// FILTROS
// ============================================================================

function WEBVD_cumpleFiltros_(
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

function WEBVD_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo || 'mes'
    ).toLowerCase();


  const hoy =
    WEBVD_hoyLocal_(
      ahora
    );


  const corte =
    WEBVD_corteAlcanzado_(
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


    if (
      corte
    ) {

      compActualInicio =
        hoy;

      compActualFin =
        hoy;


      const anterior =
        WEBVD_diaOperativoAnterior_(
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
      WEBVD_inicioSemanaLunes_(
        hoy
      );


    actualInicio =
      WEBVD_addDias_(
        lunesActual,
        -7
      );


    actualFin =
      WEBVD_addDias_(
        actualInicio,
        5
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBVD_addDias_(
        actualInicio,
        -7
      );


    compAnteriorFin =
      WEBVD_addDias_(
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
      WEBVD_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualInicio =
      WEBVD_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 1,
        1
      );


    actualFin =
      WEBVD_addDias_(
        primerMesActual,
        -1
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBVD_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 2,
        1
      );


    compAnteriorFin =
      WEBVD_addDias_(
        actualInicio,
        -1
      );


    etiqueta =
      WEBVD_nombreMes_(
        actualInicio.getMonth()
      ) +
      ' ' +
      actualInicio.getFullYear();


    etiquetaAnterior =
      WEBVD_nombreMes_(
        compAnteriorInicio.getMonth()
      ) +
      ' ' +
      compAnteriorInicio.getFullYear();

  }

  else if (
    tipo === 'anio'
  ) {

    actualInicio =
      WEBVD_fechaLocal_(
        hoy.getFullYear(),
        0,
        1
      );


    actualFin =
      hoy;


    etiqueta =
      'Año ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBVD_addDias_(
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
        WEBVD_fechaLocal_(
          hoy.getFullYear() - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBVD_fechaClampeada_(
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
      WEBVD_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualFin =
      hoy;


    etiqueta =
      WEBVD_nombreMes_(
        hoy.getMonth()
      ) +
      ' ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBVD_addDias_(
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
        WEBVD_fechaLocal_(
          hoy.getFullYear(),
          hoy.getMonth() - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBVD_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        WEBVD_nombreMes_(
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

    excluirDomingos:
      excluirDomingos,

    corteAlcanzado:
      corte,

    comparacionHabilitada:
      comparacionHabilitada,

    actualInicioKey:
      WEBVD_fechaKey_(
        actualInicio
      ),

    actualFinKey:
      WEBVD_fechaKey_(
        actualFin
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBVD_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBVD_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBVD_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBVD_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// COLUMNAS
// ============================================================================

function WEBVD_getColumnas_(
  headers
) {

  return {

    idVenta:
      WEBVD_findHeader_(
        headers,
        [
          'id_venta'
        ]
      ),

    telefono:
      WEBVD_findHeader_(
        headers,
        [
          'telefono_cliente'
        ]
      ),

    cliente:
      WEBVD_findHeader_(
        headers,
        [
          'cliente'
        ]
      ),

    ruta:
      WEBVD_findHeader_(
        headers,
        [
          'ruta'
        ]
      ),

    vendedor:
      WEBVD_findHeader_(
        headers,
        [
          'id_vendedor'
        ]
      ),

    fechaHora:
      WEBVD_findHeader_(
        headers,
        [
          'fecha_hora'
        ]
      ),

    fecha:
      WEBVD_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    agua:
      WEBVD_findHeader_(
        headers,
        [
          'agua 20 litros',
          'agua_20_litros'
        ]
      ),

    superIce:
      WEBVD_findHeader_(
        headers,
        [
          'superice'
        ]
      ),

    importe:
      WEBVD_findHeader_(
        headers,
        [
          'importe'
        ]
      ),

    giro:
      WEBVD_findHeader_(
        headers,
        [
          'giro'
        ]
      ),

    comision:
      WEBVD_findHeader_(
        headers,
        [
          'comision'
        ]
      )

  };

}


// ============================================================================
// MAPA VENDEDORES
// ============================================================================

function WEBVD_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBVD_HOJA_PARAMETROS
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
    WEBVD_findHeader_(
      data[0],
      [
        'id_vendedor'
      ]
    );


  const nombreCol =
    WEBVD_findHeader_(
      data[0],
      [
        'nombre'
      ]
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const id =
      WEBVD_texto_(
        WEBVD_val_(
          data[i],
          idCol
        )
      );


    const nombre =
      WEBVD_texto_(
        WEBVD_val_(
          data[i],
          nombreCol
        )
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
// MAPA GIROS
// ============================================================================

function WEBVD_getMapaGiros_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBVD_HOJA_GIROS
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
    WEBVD_findHeader_(
      data[0],
      [
        'idgiro',
        'id_giro'
      ]
    );


  const nombreCol =
    WEBVD_findHeader_(
      data[0],
      [
        'giro'
      ]
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const id =
      WEBVD_texto_(
        WEBVD_val_(
          data[i],
          idCol
        )
      );


    const nombre =
      WEBVD_texto_(
        WEBVD_val_(
          data[i],
          nombreCol
        )
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


// ============================================================================
// FECHAS
// ============================================================================

function WEBVD_hoyLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      WEBVD_TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


  return WEBVD_fechaLocal_(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2])
  );

}


function WEBVD_corteAlcanzado_(
  fecha
) {

  return Number(
    Utilities.formatDate(
      fecha,
      WEBVD_TZ,
      'H'
    )
  ) >= WEBVD_HORA_CORTE;

}


function WEBVD_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  return WEBVD_addDias_(
    fecha,
    dia === 0
      ? -6
      : 1 - dia
  );

}


function WEBVD_diaOperativoAnterior_(
  fecha
) {

  let r =
    WEBVD_addDias_(
      fecha,
      -1
    );


  while (
    r.getDay() === 0
  ) {

    r =
      WEBVD_addDias_(
        r,
        -1
      );

  }


  return r;

}


function WEBVD_fechaLocal_(
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


function WEBVD_addDias_(
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


  return WEBVD_fechaLocal_(
    r.getFullYear(),
    r.getMonth(),
    r.getDate()
  );

}


function WEBVD_fechaClampeada_(
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


  return WEBVD_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimo
    )
  );

}


function WEBVD_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBVD_TZ,
    'yyyy-MM-dd'
  );

}


function WEBVD_entreFechas_(
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


function WEBVD_parseFecha_(
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

function WEBVD_clienteKey_(
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
    return 'T:' + tel;
  }


  const nom =
    WEBVD_normalizar_(
      nombre
    );


  return nom
    ? 'N:' + nom
    : 'R:' + fallback;

}


function WEBVD_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBVD_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        WEBVD_normalizar_(
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


function WEBVD_normalizar_(
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


function WEBVD_val_(
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


function WEBVD_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBVD_numero_(
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


function WEBVD_redondear_(
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


function WEBVD_nombreMes_(
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


function WEBVD_pctTexto_(
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


function WEBVD_monedaTexto_(
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
