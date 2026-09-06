// ============================================================================
// WEB_Pedidos.gs
// ATLAS REPARTO - ANÁLISIS DE PEDIDOS Y SERVICIO
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   PEDIDOS
//   PARAMETROS
//
// IGNORA:
//   PEDIDOS MANUALES
//   Copia de PEDIDOS
//
// Analiza:
//   - pedidos recibidos
//   - entregados / cancelados / pendientes
//   - retrasos
//   - puntualidad
//   - tiempo de aceptación
//   - tiempo de entrega
//   - procedencia del pedido
//   - desempeño por vendedor
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBP_TZ = 'America/Merida';
const WEBP_HORA_CORTE = 19;

const WEBP_HOJA_PEDIDOS = 'PEDIDOS';
const WEBP_HOJA_PARAMETROS = 'PARAMETROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getPedidosDashboard(filtros) {

  try {

    filtros = filtros || {};


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const mapaVendedores =
      WEBP_getMapaVendedores_(ss);


    const shPedidos =
      ss.getSheetByName(
        WEBP_HOJA_PEDIDOS
      );


    if (!shPedidos) {

      throw new Error(
        'No existe la hoja PEDIDOS.'
      );

    }


    const registros =
      WEBP_leerPedidosHoja_(
        shPedidos,
        mapaVendedores
      );



    if (
      !registros.length
    ) {

      return {

        ok: true,
        sinDatos: true,
        mensaje: 'No hay pedidos disponibles para analizar.'

      };

    }


    const periodo =
      WEBP_getPeriodo_(
        filtros.periodo || 'mes',
        new Date()
      );


    const filtrados =
      registros.filter(
        r =>
          WEBP_cumpleFiltros_(
            r,
            filtros
          )
      );


    const actuales =
      filtrados.filter(
        r =>
          WEBP_entreFechas_(
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
            WEBP_entreFechas_(
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
            WEBP_entreFechas_(
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


    const resumen =
      WEBP_resumen_(
        actuales
      );


    const resumenCompActual =
      WEBP_resumen_(
        actualesComp
      );


    const resumenCompAnterior =
      WEBP_resumen_(
        anterioresComp
      );


    const kpis = {

      pedidos:
        WEBP_kpi_(
          resumen.pedidos,
          resumenCompActual.pedidos,
          resumenCompAnterior.pedidos,
          periodo.comparacionHabilitada
        ),

      entregados:
        WEBP_kpi_(
          resumen.entregados,
          resumenCompActual.entregados,
          resumenCompAnterior.entregados,
          periodo.comparacionHabilitada
        ),

      cancelados:
        WEBP_kpi_(
          resumen.cancelados,
          resumenCompActual.cancelados,
          resumenCompAnterior.cancelados,
          periodo.comparacionHabilitada
        ),

      retrasados:
        WEBP_kpi_(
          resumen.retrasados,
          resumenCompActual.retrasados,
          resumenCompAnterior.retrasados,
          periodo.comparacionHabilitada
        ),

      entregaPct: {
        actual:
          resumen.entregaPct,

        delta:
          periodo.comparacionHabilitada
            ? WEBP_redondear_(
                resumenCompActual.entregaPct -
                resumenCompAnterior.entregaPct
              )
            : null
      },

      puntualidadPct: {
        actual:
          resumen.puntualidadPct,

        delta:
          periodo.comparacionHabilitada
            ? WEBP_redondear_(
                resumenCompActual.puntualidadPct -
                resumenCompAnterior.puntualidadPct
              )
            : null
      },

      aceptacionMin:
        WEBP_kpi_(
          resumen.aceptacionMin,
          resumenCompActual.aceptacionMin,
          resumenCompAnterior.aceptacionMin,
          periodo.comparacionHabilitada
        ),

      entregaMin:
        WEBP_kpi_(
          resumen.entregaMin,
          resumenCompActual.entregaMin,
          resumenCompAnterior.entregaMin,
          periodo.comparacionHabilitada
        )

    };


    const estados =
      WEBP_agruparEstados_(
        actuales
      );


    const procedencias =
      WEBP_agruparProcedencias_(
        actuales
      );


    const vendedores =
      WEBP_desempenoVendedores_(
        actuales
      );


    const tendencia =
      WEBP_tendencia_(
        actuales,
        periodo
      );


    const criticos =
      WEBP_pedidosCriticos_(
        actuales
      );


    const insights =
      WEBP_insights_(
        resumen,
        kpis,
        procedencias,
        vendedores,
        periodo
      );


    const catalogos =
      WEBP_catalogos_(
        registros
      );


    return {

      ok: true,
      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBP_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      kpis:
        kpis,

      estados:
        estados,

      procedencias:
        procedencias,

      vendedores:
        vendedores,

      tendencia:
        tendencia,

      criticos:
        criticos,

      insights:
        insights,

      catalogos:
        catalogos,

      meta: {

        registrosPeriodo:
          actuales.length,

        registrosTotales:
          registros.length

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
// LECTURA DE PEDIDOS
// ============================================================================

function WEBP_leerPedidosHoja_(
  sh,
  mapaVendedores
) {

  const data =
    sh
      .getDataRange()
      .getValues();


  if (
    !data ||
    data.length < 2
  ) {
    return [];
  }


  const c =
    WEBP_columnas_(
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


    const idPedido =
      WEBP_texto_(
        WEBP_val_(
          fila,
          c.idPedido
        )
      ) ||
      'PEDIDO_ROW_' +
      (i + 1);


    let fecha =
      WEBP_parseFecha_(
        WEBP_val_(
          fila,
          c.fecha
        )
      );


    const recibido =
      WEBP_parseFecha_(
        WEBP_val_(
          fila,
          c.recibido
        )
      );


    const aceptado =
      WEBP_parseFecha_(
        WEBP_val_(
          fila,
          c.aceptado
        )
      );


    const entregaReal =
      WEBP_parseFecha_(
        WEBP_val_(
          fila,
          c.entregaReal
        )
      );


    const estimada =
      WEBP_parseFecha_(
        WEBP_val_(
          fila,
          c.estimada
        )
      );


    if (!fecha) {

      fecha =
        recibido ||
        aceptado ||
        entregaReal;

    }


    if (!fecha) {
      continue;
    }


    const estado =
      WEBP_texto_(
        WEBP_val_(
          fila,
          c.estado
        )
      ) ||
      'Sin estado';


    const pedidoCanceladoRaw =
      WEBP_texto_(
        WEBP_val_(
          fila,
          c.pedidoCancelado
        )
      );


    const cancelado =
      WEBP_esCancelado_(
        estado,
        pedidoCanceladoRaw
      );


    const entregado =
      WEBP_esEntregado_(
        estado,
        entregaReal
      );


    const retrasado =
      WEBP_esRetrasado_(
        estado,
        entregado,
        cancelado,
        estimada,
        entregaReal
      );


    const vendedorId =
      WEBP_texto_(
        WEBP_val_(
          fila,
          c.vendedor
        )
      );


    const vendedorFila =
      WEBP_texto_(
        WEBP_val_(
          fila,
          c.vendedorNombre
        )
      );


    const vendedor =
      vendedorFila ||
      mapaVendedores[
        vendedorId.toLowerCase()
      ] ||
      vendedorId ||
      'Sin vendedor';


    const tiempoAceptacion =
      WEBP_minutosEntre_(
        recibido,
        aceptado
      );


    let tiempoEntrega =
      WEBP_minutosEntre_(
        recibido,
        entregaReal
      );


    if (
      tiempoEntrega === null
    ) {

      tiempoEntrega =
        WEBP_parseDuracionMin_(
          WEBP_val_(
            fila,
            c.tiempoEntrega
          )
        );

    }


    const puntual =
      entregado &&
      estimada &&
      entregaReal
        ? entregaReal.getTime() <=
          estimada.getTime()
        : null;


    registros.push({

      idPedido:
        idPedido,

      fecha:
        fecha,

      fechaKey:
        WEBP_fechaKey_(
          fecha
        ),

      telefono:
        WEBP_texto_(
          WEBP_val_(
            fila,
            c.telefono
          )
        ),

      cliente:
        WEBP_texto_(
          WEBP_val_(
            fila,
            c.nombre
          )
        ) ||
        'Sin nombre',

      vendedorId:
        vendedorId,

      vendedor:
        vendedor,

      estado:
        estado,

      procedencia:
        WEBP_texto_(
          WEBP_val_(
            fila,
            c.procedencia
          )
        ) ||
        'Sin procedencia',

      recibido:
        recibido,

      aceptado:
        aceptado,

      estimada:
        estimada,

      entregaReal:
        entregaReal,

      tiempoAceptacion:
        tiempoAceptacion,

      tiempoEntrega:
        tiempoEntrega,

      entregado:
        entregado,

      cancelado:
        cancelado,

      retrasado:
        retrasado,

      puntual:
        puntual

    });

  }


  return registros;

}


// ============================================================================
// RESUMEN
// ============================================================================

function WEBP_resumen_(
  registros
) {

  const pedidos =
    registros.length;


  const entregados =
    registros.filter(
      x => x.entregado
    ).length;


  const cancelados =
    registros.filter(
      x => x.cancelado
    ).length;


  const retrasados =
    registros.filter(
      x => x.retrasado
    ).length;


  const baseEntrega =
    Math.max(
      0,
      pedidos -
      cancelados
    );


  const entregaPct =
    baseEntrega
      ? WEBP_redondear_(
          entregados /
          baseEntrega *
          100
        )
      : 0;


  const puntualidadEvaluables =
    registros.filter(
      x =>
        x.entregado &&
        x.puntual !== null
    );


  const puntuales =
    puntualidadEvaluables.filter(
      x => x.puntual
    ).length;


  const puntualidadPct =
    puntualidadEvaluables.length
      ? WEBP_redondear_(
          puntuales /
          puntualidadEvaluables.length *
          100
        )
      : 0;


  const aceptaciones =
    registros
      .map(
        x => x.tiempoAceptacion
      )
      .filter(
        x =>
          x !== null &&
          x >= 0 &&
          x <= 1440
      );


  const entregas =
    registros
      .map(
        x => x.tiempoEntrega
      )
      .filter(
        x =>
          x !== null &&
          x >= 0 &&
          x <= 2880
      );


  return {

    pedidos:
      pedidos,

    entregados:
      entregados,

    cancelados:
      cancelados,

    retrasados:
      retrasados,

    entregaPct:
      entregaPct,

    puntualidadPct:
      puntualidadPct,

    aceptacionMin:
      WEBP_promedio_(
        aceptaciones
      ),

    entregaMin:
      WEBP_promedio_(
        entregas
      )

  };

}


// ============================================================================
// ESTADOS
// ============================================================================

function WEBP_agruparEstados_(
  registros
) {

  const mapa = {};


  registros.forEach(
    r => {

      const estado =
        r.estado ||
        'Sin estado';


      if (
        !mapa[estado]
      ) {
        mapa[estado] = 0;
      }


      mapa[estado]++;

    }
  );


  const total =
    registros.length;


  return Object
    .keys(mapa)
    .map(
      estado => ({

        estado:
          estado,

        cantidad:
          mapa[estado],

        porcentaje:
          total
            ? WEBP_redondear_(
                mapa[estado] /
                total *
                100
              )
            : 0

      })
    )
    .sort(
      (a, b) =>
        b.cantidad -
        a.cantidad
    );

}


// ============================================================================
// PROCEDENCIAS
// ============================================================================

function WEBP_agruparProcedencias_(
  registros
) {

  const mapa = {};


  registros.forEach(
    r => {

      const procedencia =
        r.procedencia ||
        'Sin procedencia';


      if (
        !mapa[
          procedencia
        ]
      ) {

        mapa[
          procedencia
        ] = 0;

      }


      mapa[
        procedencia
      ]++;

    }
  );


  const total =
    registros.length;


  return Object
    .keys(mapa)
    .map(
      procedencia => ({

        procedencia:
          procedencia,

        cantidad:
          mapa[
            procedencia
          ],

        porcentaje:
          total
            ? WEBP_redondear_(
                mapa[
                  procedencia
                ] /
                total *
                100
              )
            : 0

      })
    )
    .sort(
      (a, b) =>
        b.cantidad -
        a.cantidad
    );

}


// ============================================================================
// VENDEDORES
// ============================================================================

function WEBP_desempenoVendedores_(
  registros
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        r.vendedorId ||
        r.vendedor ||
        'Sin vendedor';


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          vendedor:
            r.vendedor,

          pedidos:
            0,

          entregados:
            0,

          cancelados:
            0,

          retrasados:
            0,

          puntualEvaluable:
            0,

          puntuales:
            0,

          aceptaciones:
            [],

          entregas:
            []

        };

      }


      const item =
        mapa[key];


      item.pedidos++;


      if (
        r.entregado
      ) {
        item.entregados++;
      }


      if (
        r.cancelado
      ) {
        item.cancelados++;
      }


      if (
        r.retrasado
      ) {
        item.retrasados++;
      }


      if (
        r.puntual !== null
      ) {

        item.puntualEvaluable++;


        if (
          r.puntual
        ) {
          item.puntuales++;
        }

      }


      if (
        r.tiempoAceptacion !== null &&
        r.tiempoAceptacion >= 0 &&
        r.tiempoAceptacion <= 1440
      ) {

        item.aceptaciones.push(
          r.tiempoAceptacion
        );

      }


      if (
        r.tiempoEntrega !== null &&
        r.tiempoEntrega >= 0 &&
        r.tiempoEntrega <= 2880
      ) {

        item.entregas.push(
          r.tiempoEntrega
        );

      }

    }
  );


  return Object
    .values(mapa)
    .map(
      x => {

        const baseEntrega =
          Math.max(
            0,
            x.pedidos -
            x.cancelados
          );


        return {

          vendedor:
            x.vendedor,

          pedidos:
            x.pedidos,

          entregados:
            x.entregados,

          cancelados:
            x.cancelados,

          retrasados:
            x.retrasados,

          entregaPct:
            baseEntrega
              ? WEBP_redondear_(
                  x.entregados /
                  baseEntrega *
                  100
                )
              : 0,

          puntualidadPct:
            x.puntualEvaluable
              ? WEBP_redondear_(
                  x.puntuales /
                  x.puntualEvaluable *
                  100
                )
              : null,

          aceptacionMin:
            WEBP_promedio_(
              x.aceptaciones
            ),

          entregaMin:
            WEBP_promedio_(
              x.entregas
            )

        };

      }
    )
    .sort(
      (a, b) =>
        b.pedidos -
        a.pedidos
    );

}


// ============================================================================
// TENDENCIA
// ============================================================================

function WEBP_tendencia_(
  registros,
  periodo
) {

  const mapa = {};


  registros.forEach(
    r => {

      let key =
        r.fechaKey;


      let label;


      if (
        periodo.agrupacion ===
        'hora'
      ) {

        const hora =
          Utilities.formatDate(
            r.fecha,
            WEBP_TZ,
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
          WEBP_nombreMesCorto_(
            Number(
              key.substring(
                5,
                7
              )
            ) - 1
          );

      }

      else {

        const p =
          r.fechaKey.split(
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

          pedidos:
            0,

          entregados:
            0,

          cancelados:
            0,

          retrasados:
            0

        };

      }


      mapa[key].pedidos++;


      if (
        r.entregado
      ) {
        mapa[key].entregados++;
      }


      if (
        r.cancelado
      ) {
        mapa[key].cancelados++;
      }


      if (
        r.retrasado
      ) {
        mapa[key].retrasados++;
      }

    }
  );


  return Object
    .keys(mapa)
    .sort()
    .map(
      key =>
        mapa[key]
    );

}


// ============================================================================
// PEDIDOS CRÍTICOS / RECIENTES
// ============================================================================

function WEBP_pedidosCriticos_(
  registros
) {

  const prioridad =
    r => {

      if (
        r.retrasado &&
        !r.entregado &&
        !r.cancelado
      ) {
        return 5;
      }


      if (
        r.estado.toLowerCase() ===
        'sin vendedor'
      ) {
        return 4;
      }


      if (
        r.cancelado
      ) {
        return 3;
      }


      if (
        r.retrasado
      ) {
        return 2;
      }


      if (
        !r.entregado
      ) {
        return 1;
      }


      return 0;

    };


  return registros
    .filter(
      r =>
        prioridad(r) > 0
    )
    .sort(
      (a, b) => {

        const pa =
          prioridad(a);


        const pb =
          prioridad(b);


        if (
          pb !== pa
        ) {
          return pb - pa;
        }


        return (
          b.fecha.getTime() -
          a.fecha.getTime()
        );

      }
    )
    .slice(
      0,
      40
    )
    .map(
      r => ({

        idPedido:
          r.idPedido,

        cliente:
          r.cliente,

        telefono:
          r.telefono,

        vendedor:
          r.vendedor,

        estado:
          r.estado,

        fecha:
          r.fechaKey,

        estimada:
          r.estimada
            ? Utilities.formatDate(
                r.estimada,
                WEBP_TZ,
                'dd/MM HH:mm'
              )
            : '',

        entregaReal:
          r.entregaReal
            ? Utilities.formatDate(
                r.entregaReal,
                WEBP_TZ,
                'dd/MM HH:mm'
              )
            : '',

        tiempoAceptacion:
          r.tiempoAceptacion,

        tiempoEntrega:
          r.tiempoEntrega,

        procedencia:
          r.procedencia

      })
    );

}


// ============================================================================
// INSIGHTS
// ============================================================================

function WEBP_insights_(
  resumen,
  kpis,
  procedencias,
  vendedores,
  periodo
) {

  const resultado = [];


  resultado.push({

    tipo:
      resumen.entregaPct >= 90
        ? 'positivo'
        : resumen.entregaPct >= 75
          ? 'info'
          : 'advertencia',

    titulo:
      'Tasa de entrega',

    texto:
      'Se ha entregado ' +
      WEBP_pctTexto_(
        resumen.entregaPct
      ) +
      ' de los pedidos no cancelados del periodo.'

  });


  if (
    resumen.puntualidadPct ||
    resumen.entregados
  ) {

    resultado.push({

      tipo:
        resumen.puntualidadPct >= 85
          ? 'positivo'
          : resumen.puntualidadPct >= 70
            ? 'info'
            : 'advertencia',

      titulo:
        'Puntualidad',

      texto:
        WEBP_pctTexto_(
          resumen.puntualidadPct
        ) +
        ' de las entregas evaluables se completaron dentro de la hora estimada.'

    });

  }


  if (
    resumen.aceptacionMin > 0
  ) {

    resultado.push({

      tipo:
        resumen.aceptacionMin <= 10
          ? 'positivo'
          : resumen.aceptacionMin <= 20
            ? 'info'
            : 'advertencia',

      titulo:
        'Tiempo de aceptación',

      texto:
        'El pedido tarda en promedio ' +
        WEBP_minTexto_(
          resumen.aceptacionMin
        ) +
        ' en ser aceptado.'

    });

  }


  if (
    procedencias.length
  ) {

    const principal =
      procedencias[0];


    resultado.push({

      tipo:
        'info',

      titulo:
        'Procedencia principal',

      texto:
        principal.procedencia +
        ' concentra ' +
        WEBP_pctTexto_(
          principal.porcentaje
        ) +
        ' de los pedidos del periodo.'

    });

  }


  const vendedoresEvaluables =
    vendedores.filter(
      x =>
        x.pedidos >= 3 &&
        x.puntualidadPct !== null
    );


  if (
    vendedoresEvaluables.length
  ) {

    const menor =
      vendedoresEvaluables
        .slice()
        .sort(
          (a, b) =>
            a.puntualidadPct -
            b.puntualidadPct
        )[0];


    if (
      menor.puntualidadPct < 80
    ) {

      resultado.push({

        tipo:
          'advertencia',

        titulo:
          'Puntualidad por vendedor',

        texto:
          menor.vendedor +
          ' presenta la menor puntualidad entre los vendedores evaluables: ' +
          WEBP_pctTexto_(
            menor.puntualidadPct
          ) +
          '.'

      });

    }

  }


  if (
    !periodo.comparacionHabilitada
  ) {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Jornada abierta',

      texto:
        'Las comparaciones del día actual se habilitan después de las 7:00 p. m.'

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

function WEBP_catalogos_(
  registros
) {

  const vendedores = {};


  registros.forEach(
    r => {

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

function WEBP_cumpleFiltros_(
  registro,
  filtros
) {

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


  return true;

}


// ============================================================================
// COLUMNAS
// ============================================================================

function WEBP_columnas_(
  headers
) {

  return {

    idPedido:
      WEBP_findHeader_(
        headers,
        [
          'id_pedido',
          'id pedido'
        ]
      ),

    telefono:
      WEBP_findHeader_(
        headers,
        [
          'telefono',
          'teléfono'
        ]
      ),

    nombre:
      WEBP_findHeader_(
        headers,
        [
          'nombre',
          'cliente',
          'nombre_cliente'
        ]
      ),

    fecha:
      WEBP_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    vendedor:
      WEBP_findHeader_(
        headers,
        [
          'id_vendedor',
          'vendedor'
        ]
      ),

    vendedorNombre:
      WEBP_findHeader_(
        headers,
        [
          'vendedor_nombre',
          'vendedor nombre'
        ]
      ),

    estado:
      WEBP_findHeader_(
        headers,
        [
          'estado',
          'estatus'
        ]
      ),

    estimada:
      WEBP_findHeader_(
        headers,
        [
          'hora_esti_entrega',
          'hora esti entrega',
          'hora_estimada',
          'hora estimada'
        ]
      ),

    recibido:
      WEBP_findHeader_(
        headers,
        [
          'fecha y hora_recibido',
          'fecha y hora recibido',
          'hora_recibido',
          'fecha_hora_recibido'
        ]
      ),

    aceptado:
      WEBP_findHeader_(
        headers,
        [
          'fechahora aceptado',
          'fecha hora aceptado',
          'fecha_hora_aceptado',
          'hora_aceptado'
        ]
      ),

    entregaReal:
      WEBP_findHeader_(
        headers,
        [
          'hora_entrega_real',
          'hora entrega real',
          'fecha_hora_entrega_real'
        ]
      ),

    tiempoEntrega:
      WEBP_findHeader_(
        headers,
        [
          'tiempo entrega real',
          'tiempo_entrega_real'
        ]
      ),

    procedencia:
      WEBP_findHeader_(
        headers,
        [
          'procedencia',
          'origen'
        ]
      ),

    pedidoCancelado:
      WEBP_findHeader_(
        headers,
        [
          'pedido_cancelado',
          'pedido cancelado',
          'cancelado'
        ]
      )

  };

}


// ============================================================================
// ESTADOS
// ============================================================================

function WEBP_esEntregado_(
  estado,
  entregaReal
) {

  const e =
    WEBP_normalizar_(
      estado
    );


  return (
    e === 'entregado' ||
    !!entregaReal
  );

}


function WEBP_esCancelado_(
  estado,
  valorCancelado
) {

  const e =
    WEBP_normalizar_(
      estado
    );


  if (
    e === 'cancelado' ||
    e === 'cancelada'
  ) {
    return true;
  }


  const c =
    WEBP_normalizar_(
      valorCancelado
    );


  return [
    'true',
    'si',
    '1',
    'cancelado',
    'cancelada'
  ].includes(
    c
  );

}


function WEBP_esRetrasado_(
  estado,
  entregado,
  cancelado,
  estimada,
  entregaReal
) {

  const e =
    WEBP_normalizar_(
      estado
    );


  if (
    e === 'retrasado' ||
    e === 'retrasada'
  ) {
    return true;
  }


  if (
    cancelado
  ) {
    return false;
  }


  if (
    entregado &&
    estimada &&
    entregaReal
  ) {

    return (
      entregaReal.getTime() >
      estimada.getTime()
    );

  }


  return false;

}


// ============================================================================
// TIEMPOS
// ============================================================================

function WEBP_minutosEntre_(
  inicio,
  fin
) {

  if (
    !inicio ||
    !fin
  ) {
    return null;
  }


  const minutos =
    (
      fin.getTime() -
      inicio.getTime()
    ) /
    60000;


  if (
    !isFinite(
      minutos
    ) ||
    minutos < 0
  ) {
    return null;
  }


  return WEBP_redondear_(
    minutos
  );

}


function WEBP_parseDuracionMin_(
  valor
) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {
    return null;
  }


  if (
    typeof valor ===
    'number'
  ) {

    if (
      valor >= 0 &&
      valor < 2
    ) {

      return WEBP_redondear_(
        valor *
        24 *
        60
      );

    }


    return WEBP_redondear_(
      valor
    );

  }


  const texto =
    String(
      valor
    ).trim();


  let m =
    texto.match(
      /^(\d+):(\d{1,2})(?::(\d{1,2}))?$/
    );


  if (m) {

    const horas =
      Number(
        m[1]
      );


    const minutos =
      Number(
        m[2]
      );


    const segundos =
      Number(
        m[3] || 0
      );


    return WEBP_redondear_(
      horas *
      60 +
      minutos +
      segundos /
      60
    );

  }


  m =
    texto.match(
      /(\d+(?:\.\d+)?)\s*(?:min|minutos?)/
    );


  if (m) {

    return WEBP_redondear_(
      Number(
        m[1]
      )
    );

  }


  const n =
    Number(
      texto
        .replace(
          ',',
          '.'
        )
    );


  return isFinite(
    n
  )
    ? WEBP_redondear_(
        n
      )
    : null;

}


function WEBP_promedio_(
  valores
) {

  if (
    !valores.length
  ) {
    return 0;
  }


  return WEBP_redondear_(
    valores.reduce(
      (a, b) =>
        a + b,
      0
    ) /
    valores.length
  );

}


// ============================================================================
// KPI
// ============================================================================

function WEBP_kpi_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  habilitada
) {

  return {

    actual:
      WEBP_redondear_(
        actualVisible
      ),

    variacion:
      habilitada
        ? WEBP_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBP_variacion_(
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


  return WEBP_redondear_(
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
// PERIODOS
// ============================================================================

function WEBP_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo || 'mes'
    ).toLowerCase();


  const hoy =
    WEBP_hoyLocal_(
      ahora
    );


  const corte =
    WEBP_corteAlcanzado_(
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
        WEBP_diaOperativoAnterior_(
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
      WEBP_inicioSemanaLunes_(
        hoy
      );


    actualInicio =
      WEBP_addDias_(
        lunesActual,
        -7
      );


    actualFin =
      WEBP_addDias_(
        actualInicio,
        5
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBP_addDias_(
        actualInicio,
        -7
      );


    compAnteriorFin =
      WEBP_addDias_(
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
      WEBP_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualInicio =
      WEBP_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 1,
        1
      );


    actualFin =
      WEBP_addDias_(
        primerMesActual,
        -1
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBP_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 2,
        1
      );


    compAnteriorFin =
      WEBP_addDias_(
        actualInicio,
        -1
      );


    etiqueta =
      WEBP_nombreMes_(
        actualInicio.getMonth()
      ) +
      ' ' +
      actualInicio.getFullYear();


    etiquetaAnterior =
      WEBP_nombreMes_(
        compAnteriorInicio.getMonth()
      ) +
      ' ' +
      compAnteriorInicio.getFullYear();

  }

  else if (
    tipo === 'anio'
  ) {

    actualInicio =
      WEBP_fechaLocal_(
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
        : WEBP_addDias_(
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
        WEBP_fechaLocal_(
          hoy.getFullYear() - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBP_fechaClampeada_(
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
      WEBP_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualFin =
      hoy;


    etiqueta =
      WEBP_nombreMes_(
        hoy.getMonth()
      ) +
      ' ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBP_addDias_(
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
        WEBP_fechaLocal_(
          hoy.getFullYear(),
          hoy.getMonth() - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBP_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        WEBP_nombreMes_(
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
      WEBP_fechaKey_(
        actualInicio
      ),

    actualFinKey:
      WEBP_fechaKey_(
        actualFin
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBP_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBP_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBP_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBP_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// PARAMETROS
// ============================================================================

function WEBP_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBP_HOJA_PARAMETROS
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
    WEBP_findHeader_(
      data[0],
      [
        'id_vendedor'
      ]
    );


  const nombreCol =
    WEBP_findHeader_(
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
      WEBP_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBP_texto_(
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
// FECHAS
// ============================================================================

function WEBP_hoyLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      WEBP_TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


  return WEBP_fechaLocal_(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2])
  );

}


function WEBP_corteAlcanzado_(
  fecha
) {

  return Number(
    Utilities.formatDate(
      fecha,
      WEBP_TZ,
      'H'
    )
  ) >= WEBP_HORA_CORTE;

}


function WEBP_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  return WEBP_addDias_(
    fecha,
    dia === 0
      ? -6
      : 1 - dia
  );

}


function WEBP_diaOperativoAnterior_(
  fecha
) {

  let r =
    WEBP_addDias_(
      fecha,
      -1
    );


  while (
    r.getDay() === 0
  ) {

    r =
      WEBP_addDias_(
        r,
        -1
      );

  }


  return r;

}


function WEBP_fechaLocal_(
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


function WEBP_addDias_(
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


  return WEBP_fechaLocal_(
    r.getFullYear(),
    r.getMonth(),
    r.getDate()
  );

}


function WEBP_fechaClampeada_(
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


  return WEBP_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimo
    )
  );

}


function WEBP_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBP_TZ,
    'yyyy-MM-dd'
  );

}


function WEBP_entreFechas_(
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


function WEBP_parseFecha_(
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

function WEBP_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBP_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        WEBP_normalizar_(
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


function WEBP_normalizar_(
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


function WEBP_val_(
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


function WEBP_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBP_redondear_(
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


function WEBP_nombreMes_(
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


function WEBP_nombreMesCorto_(
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


function WEBP_pctTexto_(
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


function WEBP_minTexto_(
  minutos
) {

  minutos =
    Number(
      minutos || 0
    );


  if (
    minutos < 60
  ) {

    return (
      WEBP_redondear_(
        minutos
      ) +
      ' min'
    );

  }


  const horas =
    Math.floor(
      minutos /
      60
    );


  const resto =
    Math.round(
      minutos %
      60
    );


  return (
    horas +
    ' h ' +
    resto +
    ' min'
  );

}
