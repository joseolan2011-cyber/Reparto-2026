// ============================================================================
// WEB_Resumen.gs
// ATLAS REPARTO - MOTOR ANALÍTICO
// ============================================================================
// SOLO LECTURA.
//
// PERIODOS DISPONIBLES:
//   - Hoy
//   - Semana anterior completa (lunes a sábado)
//   - Mes anterior completo
//   - Este mes
//   - Este año
//
// REGLA DE CIERRE:
//   - Antes de las 19:00, HOY se considera abierto.
//   - Desde las 19:00, HOY se considera cerrado.
//
// NO MODIFICA DATOS.
// NO TOCA TELEGRAM.
// NO TOCA MAKE.
// NO TOCA doPost(e).
// ============================================================================


const WEBR_TZ = 'America/Merida';

const WEBR_HORA_CORTE = 19;

const WEBR_HOJA_VENTAS = 'VENTAS';

const WEBR_HOJA_PARAMETROS = 'PARAMETROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getDashboard(filtros) {

  try {

    filtros = filtros || {};

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {

      throw new Error(
        'No fue posible identificar el Spreadsheet activo.'
      );

    }

    if (
      !ss.getSheetByName(
        WEBR_HOJA_VENTAS
      )
    ) {

      throw new Error(
        'No existe la hoja "' +
        WEBR_HOJA_VENTAS +
        '".'
      );

    }


    const mapaVendedores =
      WEBR_getMapaVendedores_(
        ss
      );


    // ==========================================================
    // CAMBIO IMPORTANTE:
    // YA NO LEE VENTAS DIRECTAMENTE DESDE SHEETS.
    // USA LA CAPA COMPARTIDA DE WEB_Datos.gs
    // ==========================================================

    const datos =
      WEB_DATA_getValues_(
        ss,
        WEBR_HOJA_VENTAS
      );


    if (
      !datos ||
      datos.length < 2
    ) {

      return {
        ok:
          true,
        sinDatos:
          true,
        mensaje:
          'La hoja VENTAS no contiene registros.'
      };

    }


    const encabezados =
      datos[0];


    const columnas =
      WEBR_getColumnasVentas_(
        encabezados
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


    const registros = [];


    for (
      let i = 1;
      i < datos.length;
      i++
    ) {

      const fila =
        datos[i];


      let fecha = null;


      if (
        columnas.fecha !== -1
      ) {

        fecha =
          WEBR_parseFecha_(
            WEBR_val_(
              fila,
              columnas.fecha
            )
          );

      }


      let fechaHora = null;


      if (
        columnas.fechaHora !== -1
      ) {

        fechaHora =
          WEBR_parseFecha_(
            WEBR_val_(
              fila,
              columnas.fechaHora
            )
          );


        if (!fecha) {

          fecha =
            fechaHora;

        }

      }


      if (!fecha) {

        continue;

      }


      const idVendedor =
        WEBR_texto_(
          WEBR_val_(
            fila,
            columnas.vendedor
          )
        );


      const vendedorNombre =
        mapaVendedores[
          idVendedor.toLowerCase()
        ] ||
        idVendedor ||
        'Sin vendedor';


      const telefono =
        WEBR_texto_(
          WEBR_val_(
            fila,
            columnas.telefono
          )
        );


      const cliente =
        WEBR_texto_(
          WEBR_val_(
            fila,
            columnas.cliente
          )
        );


      const idVenta =
        WEBR_texto_(
          WEBR_val_(
            fila,
            columnas.idVenta
          )
        );


      registros.push({

        rowNumber:
          i + 1,

        idVenta:
          idVenta ||
          '__ROW_' +
          (i + 1),

        fecha:
          fecha,

        fechaHora:
          fechaHora ||
          fecha,

        fechaKey:
          Utilities.formatDate(
            fecha,
            WEBR_TZ,
            'yyyy-MM-dd'
          ),

        telefono:
          telefono,

        cliente:
          cliente,

        clienteKey:
          telefono ||
          cliente.toLowerCase() ||
          '__CLIENTE_' +
          (i + 1),

        ruta:
          WEBR_texto_(
            WEBR_val_(
              fila,
              columnas.ruta
            )
          ) ||
          'Sin ruta',

        vendedorId:
          idVendedor ||
          'Sin vendedor',

        vendedorNombre:
          vendedorNombre,

        giro:
          WEBR_texto_(
            WEBR_val_(
              fila,
              columnas.giro
            )
          ) ||
          'Sin giro',

        agua:
          WEBR_numero_(
            WEBR_val_(
              fila,
              columnas.agua
            )
          ),

        superIce:
          WEBR_numero_(
            WEBR_val_(
              fila,
              columnas.superIce
            )
          ),

        importe:
          WEBR_numero_(
            WEBR_val_(
              fila,
              columnas.importe
            )
          ),

        comision:
          WEBR_numero_(
            WEBR_val_(
              fila,
              columnas.comision
            )
          )

      });

    }


    const catalogos =
      WEBR_crearCatalogos_(
        registros,
        mapaVendedores
      );


    const periodo =
      WEBR_getPeriodo_(
        filtros.periodo ||
        'mes',
        new Date()
      );


    const registrosFiltrados =
      registros.filter(
        registro =>
          WEBR_cumpleFiltros_(
            registro,
            filtros
          )
      );


    const actuales =
      registrosFiltrados.filter(
        registro =>
          WEBR_entreFechas_(
            registro.fechaKey,
            periodo.actualInicioKey,
            periodo.actualFinKey
          ) &&
          WEBR_fechaPermitidaEnPeriodo_(
            registro.fecha,
            periodo
          )
      );


    let actualesComparacion = [];
    let anterioresComparacion = [];


    if (
      periodo.comparacionHabilitada
    ) {

      actualesComparacion =
        registrosFiltrados.filter(
          registro =>
            WEBR_entreFechas_(
              registro.fechaKey,
              periodo.comparacionActualInicioKey,
              periodo.comparacionActualFinKey
            ) &&
            WEBR_fechaPermitidaComparacion_(
              registro.fecha,
              periodo
            )
        );


      anterioresComparacion =
        registrosFiltrados.filter(
          registro =>
            WEBR_entreFechas_(
              registro.fechaKey,
              periodo.comparacionAnteriorInicioKey,
              periodo.comparacionAnteriorFinKey
            ) &&
            WEBR_fechaPermitidaComparacion_(
              registro.fecha,
              periodo
            )
        );

    }


    const resumenActual =
      WEBR_agruparTotales_(
        actuales
      );


    const resumenActualComparacion =
      WEBR_agruparTotales_(
        actualesComparacion
      );


    const resumenAnteriorComparacion =
      WEBR_agruparTotales_(
        anterioresComparacion
      );


    const kpis = {

      venta:
        WEBR_crearKPI_(
          resumenActual.venta,
          resumenActualComparacion.venta,
          resumenAnteriorComparacion.venta,
          periodo.comparacionHabilitada
        ),

      agua:
        WEBR_crearKPI_(
          resumenActual.agua,
          resumenActualComparacion.agua,
          resumenAnteriorComparacion.agua,
          periodo.comparacionHabilitada
        ),

      superIce:
        WEBR_crearKPI_(
          resumenActual.superIce,
          resumenActualComparacion.superIce,
          resumenAnteriorComparacion.superIce,
          periodo.comparacionHabilitada
        ),

      tickets:
        WEBR_crearKPI_(
          resumenActual.tickets,
          resumenActualComparacion.tickets,
          resumenAnteriorComparacion.tickets,
          periodo.comparacionHabilitada
        ),

      ticketPromedio:
        WEBR_crearKPI_(
          resumenActual.ticketPromedio,
          resumenActualComparacion.ticketPromedio,
          resumenAnteriorComparacion.ticketPromedio,
          periodo.comparacionHabilitada
        ),

      clientes:
        WEBR_crearKPI_(
          resumenActual.clientes,
          resumenActualComparacion.clientes,
          resumenAnteriorComparacion.clientes,
          periodo.comparacionHabilitada
        ),

      comision:
        WEBR_crearKPI_(
          resumenActual.comision,
          resumenActualComparacion.comision,
          resumenAnteriorComparacion.comision,
          periodo.comparacionHabilitada
        ),

      ventaCliente:
        WEBR_crearKPI_(
          resumenActual.ventaCliente,
          resumenActualComparacion.ventaCliente,
          resumenAnteriorComparacion.ventaCliente,
          periodo.comparacionHabilitada
        )

    };


    const tendencia =
      WEBR_crearTendencia_(
        actuales,
        periodo
      );


    const vendedores =
      WEBR_crearRanking_(
        actuales,
        actualesComparacion,
        anterioresComparacion,
        'vendedorId',
        'vendedorNombre',
        resumenActual.venta,
        periodo.comparacionHabilitada
      );


    const rutas =
      WEBR_crearRanking_(
        actuales,
        actualesComparacion,
        anterioresComparacion,
        'ruta',
        'ruta',
        resumenActual.venta,
        periodo.comparacionHabilitada
      );


    const insights =
      WEBR_crearInsights_(
        kpis,
        vendedores,
        rutas,
        periodo
      );


    return {

      ok:
        true,

      sinDatos:
        false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBR_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      filtrosAplicados: {

        periodo:
          filtros.periodo ||
          'mes',

        ruta:
          filtros.ruta ||
          '',

        vendedor:
          filtros.vendedor ||
          '',

        giro:
          filtros.giro ||
          ''

      },

      catalogos:
        catalogos,

      kpis:
        kpis,

      tendencia:
        tendencia,

      vendedores:
        vendedores,

      rutas:
        rutas,

      insights:
        insights,

      meta: {

        registrosTotales:
          registros.length,

        registrosPeriodo:
          actuales.length,

        registrosComparacionActual:
          actualesComparacion.length,

        registrosComparacionAnterior:
          anterioresComparacion.length,

        corteAlcanzado:
          periodo.corteAlcanzado,

        horaCorte:
          '19:00',

        fuenteDatos:
          'WEB_Datos cache compartida'

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



// ============================================================================
// PERIODOS
// ============================================================================

function WEBR_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo ||
      'mes'
    ).toLowerCase();


  const hoy =
    WEBR_hoyLocal_(
      ahora
    );


  const corteAlcanzado =
    WEBR_corteAlcanzado_(
      ahora
    );


  let inicioActual;
  let finActual;

  let compActualInicio = null;
  let compActualFin = null;

  let compAnteriorInicio = null;
  let compAnteriorFin = null;

  let comparacionHabilitada =
    true;

  let etiqueta = '';

  let etiquetaAnterior = '';

  let agrupacion =
    'dia';

  let excluirDomingos =
    false;


  // ==========================================================================
  // HOY
  // ==========================================================================

  if (
    tipo === 'hoy'
  ) {

    inicioActual =
      hoy;

    finActual =
      hoy;

    etiqueta =
      'Hoy';

    agrupacion =
      'hora';


    if (
      corteAlcanzado
    ) {

      compActualInicio =
        hoy;

      compActualFin =
        hoy;


      const anterior =
        WEBR_diaOperativoAnterior_(
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
        'comparación disponible después de las 7:00 p. m.';

    }

  }


  // ==========================================================================
  // SEMANA ANTERIOR COMPLETA
  // LUNES A SÁBADO
  // ==========================================================================

  else if (
    tipo === 'semana_anterior'
  ) {

    excluirDomingos =
      true;


    const lunesSemanaActual =
      WEBR_inicioSemanaLunes_(
        hoy
      );


    const lunesSemanaAnterior =
      WEBR_addDias_(
        lunesSemanaActual,
        -7
      );


    const sabadoSemanaAnterior =
      WEBR_addDias_(
        lunesSemanaAnterior,
        5
      );


    inicioActual =
      lunesSemanaAnterior;

    finActual =
      sabadoSemanaAnterior;


    const lunesDosSemanasAntes =
      WEBR_addDias_(
        lunesSemanaAnterior,
        -7
      );


    const sabadoDosSemanasAntes =
      WEBR_addDias_(
        lunesDosSemanasAntes,
        5
      );


    compActualInicio =
      lunesSemanaAnterior;

    compActualFin =
      sabadoSemanaAnterior;


    compAnteriorInicio =
      lunesDosSemanasAntes;

    compAnteriorFin =
      sabadoDosSemanasAntes;


    etiqueta =
      'Semana anterior';


    etiquetaAnterior =
      'semana previa';


    agrupacion =
      'dia';

  }


  // ==========================================================================
  // MES ANTERIOR COMPLETO
  // ==========================================================================

  else if (
    tipo === 'mes_anterior'
  ) {

    const year =
      hoy.getFullYear();


    const month =
      hoy.getMonth();


    const primerDiaMesActual =
      WEBR_fechaLocal_(
        year,
        month,
        1
      );


    const primerDiaMesAnterior =
      WEBR_fechaLocal_(
        year,
        month - 1,
        1
      );


    const ultimoDiaMesAnterior =
      WEBR_addDias_(
        primerDiaMesActual,
        -1
      );


    inicioActual =
      primerDiaMesAnterior;

    finActual =
      ultimoDiaMesAnterior;


    const primerDiaDosMesesAntes =
      WEBR_fechaLocal_(
        year,
        month - 2,
        1
      );


    const ultimoDiaDosMesesAntes =
      WEBR_addDias_(
        primerDiaMesAnterior,
        -1
      );


    compActualInicio =
      primerDiaMesAnterior;

    compActualFin =
      ultimoDiaMesAnterior;


    compAnteriorInicio =
      primerDiaDosMesesAntes;

    compAnteriorFin =
      ultimoDiaDosMesesAntes;


    etiqueta =
      WEBR_nombreMes_(
        primerDiaMesAnterior.getMonth()
      ) +
      ' ' +
      primerDiaMesAnterior.getFullYear();


    etiquetaAnterior =
      WEBR_nombreMes_(
        primerDiaDosMesesAntes.getMonth()
      ) +
      ' ' +
      primerDiaDosMesesAntes.getFullYear();


    agrupacion =
      'dia';

  }


  // ==========================================================================
  // ESTE AÑO
  // ==========================================================================

  else if (
    tipo === 'anio'
  ) {

    const year =
      hoy.getFullYear();


    inicioActual =
      WEBR_fechaLocal_(
        year,
        0,
        1
      );


    finActual =
      hoy;


    etiqueta =
      'Año ' +
      year;


    agrupacion =
      'mes';


    const ultimoDiaComparable =
      corteAlcanzado
        ? hoy
        : WEBR_addDias_(
            hoy,
            -1
          );


    if (
      ultimoDiaComparable <
      inicioActual
    ) {

      comparacionHabilitada =
        false;


      etiquetaAnterior =
        'mismo periodo del año anterior';

    }

    else {

      compActualInicio =
        inicioActual;

      compActualFin =
        ultimoDiaComparable;


      compAnteriorInicio =
        WEBR_fechaLocal_(
          year - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBR_fechaClampeada_(
          year - 1,
          ultimoDiaComparable.getMonth(),
          ultimoDiaComparable.getDate()
        );


      etiquetaAnterior =
        'mismo periodo de ' +
        (year - 1);


      if (
        !corteAlcanzado
      ) {

        etiquetaAnterior +=
          ' · hoy excluido';

      }

    }

  }


  // ==========================================================================
  // ESTE MES
  // ==========================================================================

  else {

    tipo =
      'mes';


    const year =
      hoy.getFullYear();


    const month =
      hoy.getMonth();


    inicioActual =
      WEBR_fechaLocal_(
        year,
        month,
        1
      );


    finActual =
      hoy;


    etiqueta =
      WEBR_nombreMes_(
        month
      ) +
      ' ' +
      year;


    const ultimoDiaComparable =
      corteAlcanzado
        ? hoy
        : WEBR_addDias_(
            hoy,
            -1
          );


    if (
      ultimoDiaComparable <
      inicioActual
    ) {

      comparacionHabilitada =
        false;


      etiquetaAnterior =
        'mes anterior · disponible después de las 7:00 p. m.';

    }

    else {

      compActualInicio =
        inicioActual;

      compActualFin =
        ultimoDiaComparable;


      const mesAnterior =
        WEBR_fechaLocal_(
          year,
          month - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBR_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoDiaComparable.getDate()
        );


      etiquetaAnterior =
        WEBR_nombreMes_(
          mesAnterior.getMonth()
        ) +
        ' ' +
        mesAnterior.getFullYear() +
        ' al ' +
        compAnteriorFin.getDate();


      if (
        !corteAlcanzado
      ) {

        etiquetaAnterior +=
          ' · hoy excluido';

      }

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
      corteAlcanzado,

    comparacionHabilitada:
      comparacionHabilitada,

    actualInicioKey:
      WEBR_fechaKey_(
        inicioActual
      ),

    actualFinKey:
      WEBR_fechaKey_(
        finActual
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBR_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBR_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBR_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBR_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// REGLAS DEL PERIODO
// ============================================================================

function WEBR_fechaPermitidaEnPeriodo_(
  fecha,
  periodo
) {

  if (
    periodo.excluirDomingos &&
    fecha.getDay() === 0
  ) {

    return false;

  }


  return true;

}


function WEBR_fechaPermitidaComparacion_(
  fecha,
  periodo
) {

  if (
    periodo.excluirDomingos &&
    fecha.getDay() === 0
  ) {

    return false;

  }


  return true;

}


// ============================================================================
// INICIO DE SEMANA - LUNES
// ============================================================================

function WEBR_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  const desplazamiento =
    dia === 0
      ? -6
      : 1 - dia;


  return WEBR_addDias_(
    fecha,
    desplazamiento
  );

}


// ============================================================================
// COLUMNAS
// ============================================================================

function WEBR_getColumnasVentas_(
  encabezados
) {

  return {

    idVenta:
      WEBR_findHeader_(
        encabezados,
        [
          'id_venta',
          'id venta'
        ]
      ),

    telefono:
      WEBR_findHeader_(
        encabezados,
        [
          'telefono_cliente',
          'telefono cliente',
          'telefono'
        ]
      ),

    cliente:
      WEBR_findHeader_(
        encabezados,
        [
          'cliente',
          'nombre_cliente'
        ]
      ),

    ruta:
      WEBR_findHeader_(
        encabezados,
        [
          'ruta'
        ]
      ),

    vendedor:
      WEBR_findHeader_(
        encabezados,
        [
          'id_vendedor',
          'vendedor'
        ]
      ),

    fechaHora:
      WEBR_findHeader_(
        encabezados,
        [
          'fecha_hora',
          'fecha hora'
        ]
      ),

    fecha:
      WEBR_findHeader_(
        encabezados,
        [
          'fecha'
        ]
      ),

    agua:
      WEBR_findHeader_(
        encabezados,
        [
          'agua 20 litros',
          'agua_20_litros',
          'agua20litros'
        ]
      ),

    superIce:
      WEBR_findHeader_(
        encabezados,
        [
          'superice',
          'super ice'
        ]
      ),

    importe:
      WEBR_findHeader_(
        encabezados,
        [
          'importe'
        ]
      ),

    giro:
      WEBR_findHeader_(
        encabezados,
        [
          'giro'
        ]
      ),

    comision:
      WEBR_findHeader_(
        encabezados,
        [
          'comision',
          'comisión'
        ]
      )

  };

}


// ============================================================================
// VENDEDORES
// ============================================================================

function WEBR_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const datos =
    WEB_DATA_getValues_(
      ss,
      WEBR_HOJA_PARAMETROS
    );


  if (
    !datos ||
    datos.length < 2
  ) {

    return resultado;

  }


  const headers =
    datos[0];


  const colId =
    WEBR_findHeader_(
      headers,
      [
        'id_vendedor',
        'id vendedor'
      ]
    );


  const colNombre =
    WEBR_findHeader_(
      headers,
      [
        'nombre',
        'vendedor_nombre',
        'vendedor nombre'
      ]
    );


  if (
    colId === -1 ||
    colNombre === -1
  ) {

    return resultado;

  }


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const id =
      WEBR_texto_(
        datos[i][colId]
      );


    const nombre =
      WEBR_texto_(
        datos[i][colNombre]
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
// CATÁLOGOS
// ============================================================================

function WEBR_crearCatalogos_(
  registros,
  mapaVendedores
) {

  const rutas =
    new Set();


  const giros =
    new Set();


  const vendedoresMap = {};


  registros.forEach(
    r => {

      if (r.ruta) {
        rutas.add(r.ruta);
      }


      if (r.giro) {
        giros.add(r.giro);
      }


      if (r.vendedorId) {

        vendedoresMap[
          r.vendedorId
        ] =
          r.vendedorNombre;

      }

    }
  );


  return {

    rutas:
      Array
        .from(rutas)
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
        .from(giros)
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
        .keys(vendedoresMap)
        .map(
          id => ({

            valor:
              id,

            nombre:
              vendedoresMap[id] ||
              mapaVendedores[
                String(id).toLowerCase()
              ] ||
              id

          })
        )
        .sort(
          (a, b) =>
            a.nombre.localeCompare(
              b.nombre,
              'es',
              {
                sensitivity: 'base'
              }
            )
        )

  };

}


// ============================================================================
// FILTROS
// ============================================================================

function WEBR_cumpleFiltros_(
  registro,
  filtros
) {

  if (
    filtros.ruta &&
    filtros.ruta !== '__ALL__' &&
    String(registro.ruta) !==
      String(filtros.ruta)
  ) {

    return false;

  }


  if (
    filtros.vendedor &&
    filtros.vendedor !== '__ALL__' &&
    String(registro.vendedorId) !==
      String(filtros.vendedor)
  ) {

    return false;

  }


  if (
    filtros.giro &&
    filtros.giro !== '__ALL__' &&
    String(registro.giro) !==
      String(filtros.giro)
  ) {

    return false;

  }


  return true;

}


// ============================================================================
// TOTALES
// ============================================================================

function WEBR_agruparTotales_(
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


  const cantidadTickets =
    tickets.size;


  const cantidadClientes =
    clientes.size;


  return {

    venta:
      WEBR_redondear_(
        venta
      ),

    agua:
      WEBR_redondear_(
        agua
      ),

    superIce:
      WEBR_redondear_(
        superIce
      ),

    comision:
      WEBR_redondear_(
        comision
      ),

    tickets:
      cantidadTickets,

    clientes:
      cantidadClientes,

    ticketPromedio:
      cantidadTickets
        ? WEBR_redondear_(
            venta /
            cantidadTickets
          )
        : 0,

    ventaCliente:
      cantidadClientes
        ? WEBR_redondear_(
            venta /
            cantidadClientes
          )
        : 0

  };

}


// ============================================================================
// KPI
// ============================================================================

function WEBR_crearKPI_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  comparacionHabilitada
) {

  return {

    actual:
      WEBR_redondear_(
        actualVisible
      ),

    actualComparable:
      comparacionHabilitada
        ? WEBR_redondear_(
            actualComparable
          )
        : null,

    anterior:
      comparacionHabilitada
        ? WEBR_redondear_(
            anteriorComparable
          )
        : null,

    variacion:
      comparacionHabilitada
        ? WEBR_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBR_variacion_(
  actual,
  anterior
) {

  actual =
    Number(
      actual ||
      0
    );


  anterior =
    Number(
      anterior ||
      0
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


  return WEBR_redondear_(
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

function WEBR_crearTendencia_(
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
          Number(
            Utilities.formatDate(
              r.fechaHora,
              WEBR_TZ,
              'H'
            )
          );


        key =
          String(
            hora
          ).padStart(
            2,
            '0'
          );


        label =
          key +
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


        const mes =
          Number(
            key.substring(
              5,
              7
            )
          ) -
          1;


        label =
          WEBR_nombreMesCorto_(
            mes
          );

      }

      else {

        key =
          r.fechaKey;


        const partes =
          key.split(
            '-'
          );


        label =
          partes[2] +
          '/' +
          partes[1];

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
          mapa[key].key,

        label:
          mapa[key].label,

        venta:
          WEBR_redondear_(
            mapa[key].venta
          ),

        agua:
          WEBR_redondear_(
            mapa[key].agua
          ),

        superIce:
          WEBR_redondear_(
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
// RANKINGS
// ============================================================================

function WEBR_crearRanking_(
  actualesVisibles,
  actualesComparacion,
  anterioresComparacion,
  campoKey,
  campoNombre,
  ventaTotalVisible,
  comparacionHabilitada
) {

  const visibleMap =
    WEBR_agruparPorCampo_(
      actualesVisibles,
      campoKey,
      campoNombre
    );


  const actualCompMap =
    WEBR_agruparPorCampo_(
      actualesComparacion,
      campoKey,
      campoNombre
    );


  const anteriorCompMap =
    WEBR_agruparPorCampo_(
      anterioresComparacion,
      campoKey,
      campoNombre
    );


  const resultado = [];


  Object
    .keys(
      visibleMap
    )
    .forEach(
      key => {

        const visible =
          visibleMap[key];


        const actualComp =
          actualCompMap[key] ||
          {
            venta: 0
          };


        const anteriorComp =
          anteriorCompMap[key] ||
          {
            venta: 0
          };


        resultado.push({

          key:
            key,

          nombre:
            visible.nombre,

          venta:
            WEBR_redondear_(
              visible.venta
            ),

          agua:
            WEBR_redondear_(
              visible.agua
            ),

          superIce:
            WEBR_redondear_(
              visible.superIce
            ),

          tickets:
            visible
              .tickets
              .size,

          clientes:
            visible
              .clientes
              .size,

          ticketPromedio:
            visible.tickets.size
              ? WEBR_redondear_(
                  visible.venta /
                  visible.tickets.size
                )
              : 0,

          participacion:
            ventaTotalVisible
              ? WEBR_redondear_(
                  (
                    visible.venta /
                    ventaTotalVisible
                  ) *
                  100
                )
              : 0,

          variacion:
            comparacionHabilitada
              ? WEBR_variacion_(
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


  return resultado.slice(
    0,
    10
  );

}


function WEBR_agruparPorCampo_(
  registros,
  campoKey,
  campoNombre
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        String(
          r[campoKey] ||
          'Sin dato'
        );


      const nombre =
        String(
          r[campoNombre] ||
          key
        );


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          nombre:
            nombre,

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
// INSIGHTS
// ============================================================================

function WEBR_crearInsights_(
  kpis,
  vendedores,
  rutas,
  periodo
) {

  const insights = [];


  if (
    !periodo.comparacionHabilitada
  ) {

    insights.push({

      tipo:
        'info',

      titulo:
        'Jornada todavía abierta',

      texto:
        'Los datos actuales se muestran normalmente, pero Atlas habilitará las comparaciones después de las 7:00 p. m.'

    });

  }

  else {

    const varVenta =
      kpis.venta.variacion;


    if (
      varVenta === null
    ) {

      insights.push({

        tipo:
          'info',

        titulo:
          'Sin referencia comparable',

        texto:
          'El periodo anterior no contiene información suficiente para calcular una variación confiable.'

      });

    }

    else if (
      varVenta >= 5
    ) {

      insights.push({

        tipo:
          'positivo',

        titulo:
          'La venta está creciendo',

        texto:
          'La venta presenta un crecimiento de ' +
          WEBR_pctTexto_(
            varVenta
          ) +
          ' frente a ' +
          periodo.etiquetaAnterior +
          '.'

      });

    }

    else if (
      varVenta <= -5
    ) {

      insights.push({

        tipo:
          'negativo',

        titulo:
          'La venta requiere atención',

        texto:
          'La venta se encuentra ' +
          WEBR_pctTexto_(
            Math.abs(
              varVenta
            )
          ) +
          ' por debajo de ' +
          periodo.etiquetaAnterior +
          '.'

      });

    }

    else {

      insights.push({

        tipo:
          'neutral',

        titulo:
          'Venta estable',

        texto:
          'La variación frente a ' +
          periodo.etiquetaAnterior +
          ' es de ' +
          WEBR_pctTexto_(
            varVenta
          ) +
          '.'

      });

    }

  }


  if (
    vendedores.length
  ) {

    const top =
      vendedores[0];


    insights.push({

      tipo:
        'positivo',

      titulo:
        'Mayor aportación comercial',

      texto:
        top.nombre +
        ' genera ' +
        WEBR_pctTexto_(
          top.participacion
        ) +
        ' de la venta del periodo, con ' +
        WEBR_monedaTexto_(
          top.venta
        ) +
        '.'

    });

  }


  if (
    rutas.length
  ) {

    const topRuta =
      rutas[0];


    insights.push({

      tipo:
        'info',

      titulo:
        'Ruta con mayor venta',

      texto:
        'La ruta ' +
        topRuta.nombre +
        ' concentra ' +
        WEBR_pctTexto_(
          topRuta.participacion
        ) +
        ' de la venta del periodo.'

    });

  }


  if (
    periodo.comparacionHabilitada
  ) {

    const varClientes =
      kpis.clientes.variacion;


    if (
      varClientes !== null &&
      Math.abs(
        varClientes
      ) >= 5
    ) {

      insights.push({

        tipo:
          varClientes >= 0
            ? 'positivo'
            : 'advertencia',

        titulo:
          varClientes >= 0
            ? 'Más clientes compradores'
            : 'Menos clientes compradores',

        texto:
          'La cantidad de clientes compradores cambió ' +
          WEBR_pctTexto_(
            varClientes
          ) +
          ' frente al periodo anterior.'

      });

    }


    const vendedoresCaida =
      vendedores
        .filter(
          v =>
            v.variacion !== null &&
            v.variacion <= -10
        )
        .sort(
          (a, b) =>
            a.variacion -
            b.variacion
        );


    if (
      vendedoresCaida.length
    ) {

      const vendedor =
        vendedoresCaida[0];


      insights.push({

        tipo:
          'advertencia',

        titulo:
          'Vendedor con caída relevante',

        texto:
          vendedor.nombre +
          ' presenta una variación de ' +
          WEBR_pctTexto_(
            vendedor.variacion
          ) +
          ' frente al periodo comparable.'

      });

    }

  }


  return insights.slice(
    0,
    5
  );

}


// ============================================================================
// FECHAS
// ============================================================================

function WEBR_hoyLocal_(
  fecha
) {

  const texto =
    Utilities.formatDate(
      fecha,
      WEBR_TZ,
      'yyyy-MM-dd'
    );


  const partes =
    texto.split(
      '-'
    );


  return WEBR_fechaLocal_(
    Number(
      partes[0]
    ),
    Number(
      partes[1]
    ) - 1,
    Number(
      partes[2]
    )
  );

}


function WEBR_corteAlcanzado_(
  fecha
) {

  const hora =
    Number(
      Utilities.formatDate(
        fecha,
        WEBR_TZ,
        'H'
      )
    );


  return (
    hora >=
    WEBR_HORA_CORTE
  );

}


function WEBR_diaOperativoAnterior_(
  fecha
) {

  let resultado =
    WEBR_addDias_(
      fecha,
      -1
    );


  while (
    resultado.getDay() === 0
  ) {

    resultado =
      WEBR_addDias_(
        resultado,
        -1
      );

  }


  return resultado;

}


function WEBR_parseFecha_(
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
    typeof valor === 'number' &&
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
    valor === undefined
  ) {

    return null;

  }


  const texto =
    String(
      valor
    ).trim();


  if (!texto) {
    return null;
  }


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

    let p1 =
      Number(m[1]);


    let p2 =
      Number(m[2]);


    const year =
      Number(m[3]);


    let dia;

    let mes;


    if (
      p2 > 12
    ) {

      mes =
        p1;

      dia =
        p2;

    }

    else {

      dia =
        p1;

      mes =
        p2;

    }


    return new Date(

      year,

      mes - 1,

      dia,

      Number(m[4] || 12),

      Number(m[5] || 0),

      Number(m[6] || 0)

    );

  }


  const intento =
    new Date(
      texto
    );


  if (
    !isNaN(
      intento.getTime()
    )
  ) {

    return intento;

  }


  return null;

}


function WEBR_fechaLocal_(
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


function WEBR_addDias_(
  fecha,
  dias
) {

  const resultado =
    new Date(
      fecha.getTime()
    );


  resultado.setDate(
    resultado.getDate() +
    dias
  );


  return WEBR_fechaLocal_(
    resultado.getFullYear(),
    resultado.getMonth(),
    resultado.getDate()
  );

}


function WEBR_fechaClampeada_(
  year,
  month,
  day
) {

  const ultimoDia =
    new Date(
      year,
      month + 1,
      0
    ).getDate();


  return WEBR_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimoDia
    )
  );

}


function WEBR_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBR_TZ,
    'yyyy-MM-dd'
  );

}


function WEBR_entreFechas_(
  key,
  inicio,
  fin
) {

  if (
    !inicio ||
    !fin
  ) {

    return false;

  }


  return (
    key >= inicio &&
    key <= fin
  );

}


// ============================================================================
// UTILIDADES
// ============================================================================

function WEBR_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBR_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const buscado =
      WEBR_normalizar_(
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


function WEBR_normalizar_(
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


function WEBR_val_(
  fila,
  indice
) {

  if (
    indice === -1 ||
    indice === undefined ||
    indice === null
  ) {

    return '';

  }


  return fila[
    indice
  ];

}


function WEBR_texto_(
  valor
) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';

  }


  return String(
    valor
  ).trim();

}


function WEBR_numero_(
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


  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {

    return 0;

  }


  const limpio =
    String(
      valor
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


function WEBR_redondear_(
  numero
) {

  return Math.round(
    Number(
      numero ||
      0
    ) *
    100
  ) /
  100;

}


function WEBR_nombreMes_(
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


function WEBR_nombreMesCorto_(
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


function WEBR_pctTexto_(
  numero
) {

  if (
    numero === null ||
    numero === undefined
  ) {

    return 'sin referencia';

  }


  const n =
    WEBR_redondear_(
      numero
    );


  return (

    (
      n > 0
        ? '+'
        : ''
    ) +

    n.toLocaleString(
      'es-MX',
      {
        maximumFractionDigits:
          1
      }
    ) +

    '%'

  );

}


function WEBR_monedaTexto_(
  numero
) {

  return '$' +
    Number(
      numero ||
      0
    )
      .toLocaleString(
        'es-MX',
        {
          maximumFractionDigits:
            0
        }
      );

}