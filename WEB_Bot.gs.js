// ============================================================================
// WEB_Bot.gs
// ATLAS REPARTO - ANÁLISIS DEL BOT / TREBLE
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   ENTRADASWEBOOK
//   PEDIDOS
//
// RELACIÓN:
//   ENTRADASWEBOOK[IDPedido] -> PEDIDOS[id_pedido]
//
// FECHA DE ACTIVIDAD:
//   ENTRADASWEBOOK[Entrada] = timestamp real del webhook
//   ENTRADASWEBOOK[Fecha]   = solo fallback
//
// Analiza:
//   - conversaciones / sesiones
//   - sesiones que terminan en pedido
//   - conversión a pedido
//   - pedidos entregados / cancelados
//   - elecciones GPT
//   - conversaciones sin pedido
//   - tendencia por periodo
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBBOT_TZ = 'America/Merida';
const WEBBOT_HORA_CORTE = 19;

const WEBBOT_HOJA_ENTRADAS = 'ENTRADASWEBOOK';
const WEBBOT_HOJA_PEDIDOS = 'PEDIDOS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getBotDashboard(filtros) {

  try {

    filtros = filtros || {};


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const shEntradas =
      ss.getSheetByName(
        WEBBOT_HOJA_ENTRADAS
      );


    if (!shEntradas) {

      throw new Error(
        'No existe la hoja ENTRADASWEBOOK.'
      );

    }


    const shPedidos =
      ss.getSheetByName(
        WEBBOT_HOJA_PEDIDOS
      );


    const mapaPedidos =
      shPedidos
        ? WEBBOT_mapaPedidos_(
            shPedidos
          )
        : {};


    const data =
      shEntradas
        .getDataRange()
        .getValues();


    if (
      !data ||
      data.length < 2
    ) {

      return {

        ok: true,
        sinDatos: true,
        mensaje: 'ENTRADASWEBOOK no contiene registros.'

      };

    }


    const c =
      WEBBOT_columnas_(
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


      // ENTRADASWEBOOK[Entrada] es la marca de tiempo real
      // de inserción del webhook (DD/MM/YYYY HH:mm:ss).
      // La columna Fecha es un dato adicional del proceso y no es
      // confiable como fecha principal de la sesión.
      let fecha =
        WEBBOT_parseFecha_(
          WEBBOT_val_(
            fila,
            c.entrada
          )
        );


      if (!fecha) {

        fecha =
          WEBBOT_parseFecha_(
            WEBBOT_val_(
              fila,
              c.fecha
            )
          );

      }


      if (!fecha) {
        continue;
      }


      const sessionId =
        WEBBOT_texto_(
          WEBBOT_val_(
            fila,
            c.sessionId
          )
        );


      const nvoId =
        WEBBOT_texto_(
          WEBBOT_val_(
            fila,
            c.nvoId
          )
        );


      const telefono =
        WEBBOT_texto_(
          WEBBOT_val_(
            fila,
            c.telefono
          )
        );


      const sessionKey =
        sessionId ||
        nvoId ||
        (
          telefono
            ? telefono +
              '_' +
              WEBBOT_fechaKey_(
                fecha
              )
            : 'ROW_' +
              (i + 1)
        );


      const idPedido =
        WEBBOT_texto_(
          WEBBOT_val_(
            fila,
            c.idPedido
          )
        );


      const pedido =
        idPedido
          ? mapaPedidos[
              idPedido
            ] || null
          : null;


      registros.push({

        fecha:
          fecha,

        fechaKey:
          WEBBOT_fechaKey_(
            fecha
          ),

        sessionId:
          sessionId,

        sessionKey:
          sessionKey,

        telefono:
          telefono,

        nombre:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.nombre
            )
          ) ||
          'Sin nombre',

        entrada:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.mensajeInicial
            )
          ),

        mensajeInicial:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.mensajeInicial
            )
          ),

        eleccionGPT:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.eleccionGPT
            )
          ) ||
          'Sin elección',

        salida:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.salida
            )
          ),

        vendedor:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.vendedor
            )
          ) ||
          'Sin vendedor',

        contador:
          WEBBOT_numero_(
            WEBBOT_val_(
              fila,
              c.contador
            )
          ),

        conversacion:
          WEBBOT_texto_(
            WEBBOT_val_(
              fila,
              c.conversacion
            )
          ),

        idPedido:
          idPedido,

        tienePedido:
          !!idPedido,

        pedidoEncontrado:
          !!pedido,

        pedidoEstado:
          pedido
            ? pedido.estado
            : '',

        pedidoEntregado:
          pedido
            ? pedido.entregado
            : false,

        pedidoCancelado:
          pedido
            ? pedido.cancelado
            : false,

        pedidoRetrasado:
          pedido
            ? pedido.retrasado
            : false

      });

    }


    const periodo =
      WEBBOT_getPeriodo_(
        filtros.periodo || 'mes',
        new Date()
      );


    // ------------------------------------------------------------------------
    // IMPORTANTE:
    // Primero filtramos las FILAS por periodo y después las agrupamos por
    // sesión. Un session_id puede reaparecer en distintos días o meses.
    // Si agrupamos todo el historial antes de filtrar, la sesión conserva
    // una fecha antigua y desaparece del periodo actual.
    // ------------------------------------------------------------------------

    const filasActuales =
      registros.filter(
        r =>
          WEBBOT_entreFechas_(
            r.fechaKey,
            periodo.actualInicioKey,
            periodo.actualFinKey
          ) &&
          (
            !periodo.excluirDomingos ||
            r.fecha.getDay() !== 0
          )
      );


    const actuales =
      WEBBOT_colapsarSesiones_(
        filasActuales
      );


    let actualesComp = [];
    let anterioresComp = [];


    if (
      periodo.comparacionHabilitada
    ) {

      const filasActualesComp =
        registros.filter(
          r =>
            WEBBOT_entreFechas_(
              r.fechaKey,
              periodo.comparacionActualInicioKey,
              periodo.comparacionActualFinKey
            ) &&
            (
              !periodo.excluirDomingos ||
              r.fecha.getDay() !== 0
            )
        );


      const filasAnterioresComp =
        registros.filter(
          r =>
            WEBBOT_entreFechas_(
              r.fechaKey,
              periodo.comparacionAnteriorInicioKey,
              periodo.comparacionAnteriorFinKey
            ) &&
            (
              !periodo.excluirDomingos ||
              r.fecha.getDay() !== 0
            )
        );


      actualesComp =
        WEBBOT_colapsarSesiones_(
          filasActualesComp
        );


      anterioresComp =
        WEBBOT_colapsarSesiones_(
          filasAnterioresComp
        );

    }


    // Para diagnóstico mostramos cuántas sesiones únicas hay en todo el
    // historial, sin usar esa agrupación para los cálculos por periodo.
    const sesionesHistoricas =
      WEBBOT_colapsarSesiones_(
        registros
      );


    const resumen =
      WEBBOT_resumen_(
        actuales
      );


    const resumenCompActual =
      WEBBOT_resumen_(
        actualesComp
      );


    const resumenCompAnterior =
      WEBBOT_resumen_(
        anterioresComp
      );


    const kpis = {

      sesiones:
        WEBBOT_kpi_(
          resumen.sesiones,
          resumenCompActual.sesiones,
          resumenCompAnterior.sesiones,
          periodo.comparacionHabilitada
        ),

      conPedido:
        WEBBOT_kpi_(
          resumen.conPedido,
          resumenCompActual.conPedido,
          resumenCompAnterior.conPedido,
          periodo.comparacionHabilitada
        ),

      sinPedido:
        WEBBOT_kpi_(
          resumen.sinPedido,
          resumenCompActual.sinPedido,
          resumenCompAnterior.sinPedido,
          periodo.comparacionHabilitada
        ),

      entregados:
        WEBBOT_kpi_(
          resumen.entregados,
          resumenCompActual.entregados,
          resumenCompAnterior.entregados,
          periodo.comparacionHabilitada
        ),

      cancelados:
        WEBBOT_kpi_(
          resumen.cancelados,
          resumenCompActual.cancelados,
          resumenCompAnterior.cancelados,
          periodo.comparacionHabilitada
        ),

      conversionPct: {
        actual:
          resumen.conversionPct,

        delta:
          periodo.comparacionHabilitada
            ? WEBBOT_redondear_(
                resumenCompActual.conversionPct -
                resumenCompAnterior.conversionPct
              )
            : null
      },

      entregaPedidoPct: {
        actual:
          resumen.entregaPedidoPct,

        delta:
          periodo.comparacionHabilitada
            ? WEBBOT_redondear_(
                resumenCompActual.entregaPedidoPct -
                resumenCompAnterior.entregaPedidoPct
              )
            : null
      },

      interaccionesPromedio:
        WEBBOT_kpi_(
          resumen.interaccionesPromedio,
          resumenCompActual.interaccionesPromedio,
          resumenCompAnterior.interaccionesPromedio,
          periodo.comparacionHabilitada
        )

    };


    const tendencia =
      WEBBOT_tendencia_(
        actuales,
        periodo
      );


    const elecciones =
      WEBBOT_elecciones_(
        actuales
      );


    const sinPedido =
      WEBBOT_sesionesSinPedido_(
        actuales
      );


    const vendedores =
      WEBBOT_vendedores_(
        actuales
      );


    const insights =
      WEBBOT_insights_(
        resumen,
        kpis,
        elecciones,
        vendedores,
        periodo
      );


    return {

      ok: true,
      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBBOT_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      kpis:
        kpis,

      tendencia:
        tendencia,

      elecciones:
        elecciones,

      sinPedido:
        sinPedido,

      vendedores:
        vendedores,

      insights:
        insights,

      meta: {

        filasEntradas:
          registros.length,

        sesionesTotales:
          sesionesHistoricas.length,

        filasPeriodo:
          filasActuales.length,

        sesionesPeriodo:
          actuales.length

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
// MAPA DE PEDIDOS
// ============================================================================

function WEBBOT_mapaPedidos_(
  sh
) {

  const data =
    sh
      .getDataRange()
      .getValues();


  if (
    !data ||
    data.length < 2
  ) {
    return {};
  }


  const c =
    WEBBOT_columnasPedidos_(
      data[0]
    );


  const mapa = {};


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const fila =
      data[i];


    const id =
      WEBBOT_texto_(
        WEBBOT_val_(
          fila,
          c.idPedido
        )
      );


    if (!id) {
      continue;
    }


    const estado =
      WEBBOT_texto_(
        WEBBOT_val_(
          fila,
          c.estado
        )
      ) ||
      'Sin estado';


    const entregaReal =
      WEBBOT_parseFecha_(
        WEBBOT_val_(
          fila,
          c.entregaReal
        )
      );


    const estimada =
      WEBBOT_parseFecha_(
        WEBBOT_val_(
          fila,
          c.estimada
        )
      );


    const cancelado =
      WEBBOT_esCancelado_(
        estado,
        WEBBOT_val_(
          fila,
          c.cancelado
        )
      );


    const entregado =
      WEBBOT_esEntregado_(
        estado,
        entregaReal
      );


    mapa[id] = {

      estado:
        estado,

      entregado:
        entregado,

      cancelado:
        cancelado,

      retrasado:
        WEBBOT_esRetrasado_(
          estado,
          entregado,
          cancelado,
          estimada,
          entregaReal
        )

    };

  }


  return mapa;

}


// ============================================================================
// COLAPSAR FILAS A SESIONES
// ============================================================================

function WEBBOT_colapsarSesiones_(
  registros
) {

  const mapa = {};


  registros.forEach(
    r => {

      const key =
        r.sessionKey;


      if (
        !mapa[key]
      ) {

        mapa[key] = {

          sessionKey:
            key,

          sessionId:
            r.sessionId,

          fecha:
            r.fecha,

          fechaKey:
            r.fechaKey,

          telefono:
            r.telefono,

          nombre:
            r.nombre,

          primeraEntrada:
            r.entrada ||
            r.mensajeInicial,

          ultimaSalida:
            r.salida,

          eleccionGPT:
            r.eleccionGPT,

          vendedor:
            r.vendedor,

          interacciones:
            0,

          idPedido:
            '',

          tienePedido:
            false,

          pedidoEncontrado:
            false,

          pedidoEstado:
            '',

          pedidoEntregado:
            false,

          pedidoCancelado:
            false,

          pedidoRetrasado:
            false

        };

      }


      const s =
        mapa[key];


      s.interacciones++;


      if (
        r.fecha <
        s.fecha
      ) {

        s.fecha =
          r.fecha;

        s.fechaKey =
          r.fechaKey;

      }


      if (
        r.salida
      ) {
        s.ultimaSalida =
          r.salida;
      }


      if (
        r.eleccionGPT &&
        r.eleccionGPT !==
        'Sin elección'
      ) {

        s.eleccionGPT =
          r.eleccionGPT;

      }


      if (
        r.vendedor &&
        r.vendedor !==
        'Sin vendedor'
      ) {

        s.vendedor =
          r.vendedor;

      }


      if (
        r.idPedido
      ) {

        s.idPedido =
          r.idPedido;

        s.tienePedido =
          true;

        s.pedidoEncontrado =
          r.pedidoEncontrado;

        s.pedidoEstado =
          r.pedidoEstado;

        s.pedidoEntregado =
          r.pedidoEntregado;

        s.pedidoCancelado =
          r.pedidoCancelado;

        s.pedidoRetrasado =
          r.pedidoRetrasado;

      }

    }
  );


  return Object
    .values(
      mapa
    );

}


// ============================================================================
// RESUMEN
// ============================================================================

function WEBBOT_resumen_(
  sesiones
) {

  const total =
    sesiones.length;


  const conPedido =
    sesiones.filter(
      x =>
        x.tienePedido
    ).length;


  const sinPedido =
    total -
    conPedido;


  const entregados =
    sesiones.filter(
      x =>
        x.pedidoEntregado
    ).length;


  const cancelados =
    sesiones.filter(
      x =>
        x.pedidoCancelado
    ).length;


  const conversionPct =
    total
      ? WEBBOT_redondear_(
          conPedido /
          total *
          100
        )
      : 0;


  const entregaPedidoPct =
    conPedido
      ? WEBBOT_redondear_(
          entregados /
          conPedido *
          100
        )
      : 0;


  const interacciones =
    sesiones.map(
      x =>
        x.interacciones
    );


  return {

    sesiones:
      total,

    conPedido:
      conPedido,

    sinPedido:
      sinPedido,

    entregados:
      entregados,

    cancelados:
      cancelados,

    conversionPct:
      conversionPct,

    entregaPedidoPct:
      entregaPedidoPct,

    interaccionesPromedio:
      WEBBOT_promedio_(
        interacciones
      )

  };

}


// ============================================================================
// TENDENCIA
// ============================================================================

function WEBBOT_tendencia_(
  sesiones,
  periodo
) {

  const mapa = {};


  sesiones.forEach(
    s => {

      let key =
        s.fechaKey;


      let label;


      if (
        periodo.agrupacion ===
        'hora'
      ) {

        const hora =
          Utilities.formatDate(
            s.fecha,
            WEBBOT_TZ,
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
          s.fechaKey.substring(
            0,
            7
          );


        label =
          WEBBOT_nombreMesCorto_(
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
          s.fechaKey.split(
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

          sesiones:
            0,

          pedidos:
            0,

          entregados:
            0

        };

      }


      mapa[key].sesiones++;


      if (
        s.tienePedido
      ) {
        mapa[key].pedidos++;
      }


      if (
        s.pedidoEntregado
      ) {
        mapa[key].entregados++;
      }

    }
  );


  return Object
    .keys(
      mapa
    )
    .sort()
    .map(
      key =>
        mapa[key]
    );

}


// ============================================================================
// ELECCIONES GPT
// ============================================================================

function WEBBOT_elecciones_(
  sesiones
) {

  const mapa = {};


  sesiones.forEach(
    s => {

      const eleccion =
        s.eleccionGPT ||
        'Sin elección';


      if (
        !mapa[eleccion]
      ) {

        mapa[eleccion] = {

          eleccion:
            eleccion,

          sesiones:
            0,

          pedidos:
            0

        };

      }


      mapa[eleccion].sesiones++;


      if (
        s.tienePedido
      ) {
        mapa[eleccion].pedidos++;
      }

    }
  );


  return Object
    .values(
      mapa
    )
    .map(
      x => ({

        eleccion:
          x.eleccion,

        sesiones:
          x.sesiones,

        pedidos:
          x.pedidos,

        conversionPct:
          x.sesiones
            ? WEBBOT_redondear_(
                x.pedidos /
                x.sesiones *
                100
              )
            : 0

      })
    )
    .sort(
      (a, b) =>
        b.sesiones -
        a.sesiones
    )
    .slice(
      0,
      15
    );

}


// ============================================================================
// SESIONES SIN PEDIDO
// ============================================================================

function WEBBOT_sesionesSinPedido_(
  sesiones
) {

  return sesiones
    .filter(
      s =>
        !s.tienePedido
    )
    .sort(
      (a, b) =>
        b.fecha.getTime() -
        a.fecha.getTime()
    )
    .slice(
      0,
      40
    )
    .map(
      s => ({

        fecha:
          s.fechaKey,

        sessionId:
          s.sessionId,

        telefono:
          s.telefono,

        nombre:
          s.nombre,

        entrada:
          s.primeraEntrada,

        salida:
          s.ultimaSalida,

        eleccionGPT:
          s.eleccionGPT,

        vendedor:
          s.vendedor,

        interacciones:
          s.interacciones

      })
    );

}


// ============================================================================
// VENDEDORES
// ============================================================================

function WEBBOT_vendedores_(
  sesiones
) {

  const mapa = {};


  sesiones.forEach(
    s => {

      const vendedor =
        s.vendedor ||
        'Sin vendedor';


      if (
        !mapa[vendedor]
      ) {

        mapa[vendedor] = {

          vendedor:
            vendedor,

          sesiones:
            0,

          pedidos:
            0,

          entregados:
            0

        };

      }


      mapa[vendedor].sesiones++;


      if (
        s.tienePedido
      ) {
        mapa[vendedor].pedidos++;
      }


      if (
        s.pedidoEntregado
      ) {
        mapa[vendedor].entregados++;
      }

    }
  );


  return Object
    .values(
      mapa
    )
    .map(
      x => ({

        vendedor:
          x.vendedor,

        sesiones:
          x.sesiones,

        pedidos:
          x.pedidos,

        entregados:
          x.entregados,

        conversionPct:
          x.sesiones
            ? WEBBOT_redondear_(
                x.pedidos /
                x.sesiones *
                100
              )
            : 0

      })
    )
    .sort(
      (a, b) =>
        b.sesiones -
        a.sesiones
    );

}


// ============================================================================
// INSIGHTS
// ============================================================================

function WEBBOT_insights_(
  resumen,
  kpis,
  elecciones,
  vendedores,
  periodo
) {

  const resultado = [];


  resultado.push({

    tipo:
      resumen.conversionPct >= 60
        ? 'positivo'
        : resumen.conversionPct >= 35
          ? 'info'
          : 'advertencia',

    titulo:
      'Conversión a pedido',

    texto:
      WEBBOT_pctTexto_(
        resumen.conversionPct
      ) +
      ' de las sesiones del bot terminaron con un ID de pedido.'

  });


  if (
    resumen.conPedido
  ) {

    resultado.push({

      tipo:
        resumen.entregaPedidoPct >= 85
          ? 'positivo'
          : resumen.entregaPedidoPct >= 70
            ? 'info'
            : 'advertencia',

      titulo:
        'Pedidos del bot entregados',

      texto:
        WEBBOT_pctTexto_(
          resumen.entregaPedidoPct
        ) +
        ' de los pedidos originados desde las sesiones analizadas aparecen como entregados en PEDIDOS.'

    });

  }


  if (
    resumen.sinPedido
  ) {

    resultado.push({

      tipo:
        resumen.sinPedido >
        resumen.conPedido
          ? 'advertencia'
          : 'info',

      titulo:
        'Sesiones sin pedido',

      texto:
        resumen.sinPedido +
        ' sesiones del periodo no terminaron con un ID de pedido.'

    });

  }


  if (
    elecciones.length
  ) {

    const conVolumen =
      elecciones.filter(
        x =>
          x.sesiones >= 3 &&
          x.eleccion !==
          'Sin elección'
      );


    if (
      conVolumen.length
    ) {

      const mejor =
        conVolumen
          .slice()
          .sort(
            (a, b) =>
              b.conversionPct -
              a.conversionPct
          )[0];


      resultado.push({

        tipo:
          'info',

        titulo:
          'Elección GPT con mejor conversión',

        texto:
          mejor.eleccion +
          ' convierte ' +
          WEBBOT_pctTexto_(
            mejor.conversionPct
          ) +
          ' de sus sesiones a pedido.'

      });

    }

  }


  if (
    periodo.comparacionHabilitada &&
    kpis.conversionPct.delta !== null
  ) {

    const delta =
      kpis.conversionPct.delta;


    if (
      Math.abs(
        delta
      ) >= 5
    ) {

      resultado.push({

        tipo:
          delta > 0
            ? 'positivo'
            : 'advertencia',

        titulo:
          delta > 0
            ? 'Mejora de conversión'
            : 'Caída de conversión',

        texto:
          'La conversión cambió ' +
          WEBBOT_pctTexto_(
            Math.abs(
              delta
            )
          ) +
          ' puntos porcentuales frente al periodo comparable.'

      });

    }

  }


  return resultado.slice(
    0,
    5
  );

}


// ============================================================================
// COLUMNAS ENTRADAS
// ============================================================================

function WEBBOT_columnas_(
  headers
) {

  return {

    sessionId:
      WEBBOT_findHeader_(
        headers,
        [
          'session_id'
        ]
      ),

    telefono:
      WEBBOT_findHeader_(
        headers,
        [
          'telefono'
        ]
      ),

    entrada:
      WEBBOT_findHeader_(
        headers,
        [
          'entrada'
        ]
      ),

    nombre:
      WEBBOT_findHeader_(
        headers,
        [
          'nombre'
        ]
      ),

    mensajeInicial:
      WEBBOT_findHeader_(
        headers,
        [
          'mensaje inicial',
          'mensaje_inicial'
        ]
      ),

    eleccionGPT:
      WEBBOT_findHeader_(
        headers,
        [
          'eleccion gpt',
          'eleccion_gpt'
        ]
      ),

    salida:
      WEBBOT_findHeader_(
        headers,
        [
          'salida'
        ]
      ),

    vendedor:
      WEBBOT_findHeader_(
        headers,
        [
          'vendedor'
        ]
      ),

    contador:
      WEBBOT_findHeader_(
        headers,
        [
          'contador'
        ]
      ),

    conversacion:
      WEBBOT_findHeader_(
        headers,
        [
          'conversacion'
        ]
      ),

    idPedido:
      WEBBOT_findHeader_(
        headers,
        [
          'idpedido',
          'id_pedido'
        ]
      ),

    fecha:
      WEBBOT_findHeader_(
        headers,
        [
          'fecha'
        ]
      ),

    nvoId:
      WEBBOT_findHeader_(
        headers,
        [
          'nvo_id',
          'nvo id'
        ]
      )

  };

}


// ============================================================================
// COLUMNAS PEDIDOS
// ============================================================================

function WEBBOT_columnasPedidos_(
  headers
) {

  return {

    idPedido:
      WEBBOT_findHeader_(
        headers,
        [
          'id_pedido'
        ]
      ),

    estado:
      WEBBOT_findHeader_(
        headers,
        [
          'estado',
          'estatus'
        ]
      ),

    estimada:
      WEBBOT_findHeader_(
        headers,
        [
          'hora_esti_entrega'
        ]
      ),

    entregaReal:
      WEBBOT_findHeader_(
        headers,
        [
          'hora_entrega_real'
        ]
      ),

    cancelado:
      WEBBOT_findHeader_(
        headers,
        [
          'pedido_cancelado',
          'cancelado'
        ]
      )

  };

}


// ============================================================================
// ESTADOS PEDIDO
// ============================================================================

function WEBBOT_esEntregado_(
  estado,
  entregaReal
) {

  return (
    WEBBOT_normalizar_(
      estado
    ) ===
    'entregado' ||
    !!entregaReal
  );

}


function WEBBOT_esCancelado_(
  estado,
  valor
) {

  const e =
    WEBBOT_normalizar_(
      estado
    );


  if (
    e === 'cancelado' ||
    e === 'cancelada'
  ) {
    return true;
  }


  const v =
    WEBBOT_normalizar_(
      valor
    );


  return [
    'true',
    'si',
    '1',
    'cancelado',
    'cancelada'
  ].includes(
    v
  );

}


function WEBBOT_esRetrasado_(
  estado,
  entregado,
  cancelado,
  estimada,
  entregaReal
) {

  const e =
    WEBBOT_normalizar_(
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
// KPI
// ============================================================================

function WEBBOT_kpi_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  habilitada
) {

  return {

    actual:
      WEBBOT_redondear_(
        actualVisible
      ),

    variacion:
      habilitada
        ? WEBBOT_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBBOT_variacion_(
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


  return WEBBOT_redondear_(
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

function WEBBOT_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo || 'mes'
    ).toLowerCase();


  const hoy =
    WEBBOT_hoyLocal_(
      ahora
    );


  const corte =
    WEBBOT_corteAlcanzado_(
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
        WEBBOT_diaOperativoAnterior_(
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
      WEBBOT_inicioSemanaLunes_(
        hoy
      );


    actualInicio =
      WEBBOT_addDias_(
        lunesActual,
        -7
      );


    actualFin =
      WEBBOT_addDias_(
        actualInicio,
        5
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBBOT_addDias_(
        actualInicio,
        -7
      );


    compAnteriorFin =
      WEBBOT_addDias_(
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
      WEBBOT_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualInicio =
      WEBBOT_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 1,
        1
      );


    actualFin =
      WEBBOT_addDias_(
        primerMesActual,
        -1
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBBOT_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 2,
        1
      );


    compAnteriorFin =
      WEBBOT_addDias_(
        actualInicio,
        -1
      );


    etiqueta =
      WEBBOT_nombreMes_(
        actualInicio.getMonth()
      ) +
      ' ' +
      actualInicio.getFullYear();


    etiquetaAnterior =
      WEBBOT_nombreMes_(
        compAnteriorInicio.getMonth()
      ) +
      ' ' +
      compAnteriorInicio.getFullYear();

  }

  else if (
    tipo === 'anio'
  ) {

    actualInicio =
      WEBBOT_fechaLocal_(
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
        : WEBBOT_addDias_(
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
        WEBBOT_fechaLocal_(
          hoy.getFullYear() - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBBOT_fechaClampeada_(
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
      WEBBOT_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualFin =
      hoy;


    etiqueta =
      WEBBOT_nombreMes_(
        hoy.getMonth()
      ) +
      ' ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBBOT_addDias_(
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
        WEBBOT_fechaLocal_(
          hoy.getFullYear(),
          hoy.getMonth() - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBBOT_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        WEBBOT_nombreMes_(
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
      WEBBOT_fechaKey_(
        actualInicio
      ),

    actualFinKey:
      WEBBOT_fechaKey_(
        actualFin
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBBOT_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBBOT_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBBOT_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBBOT_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// FECHAS
// ============================================================================

function WEBBOT_hoyLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      WEBBOT_TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


  return WEBBOT_fechaLocal_(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2])
  );

}


function WEBBOT_corteAlcanzado_(
  fecha
) {

  return Number(
    Utilities.formatDate(
      fecha,
      WEBBOT_TZ,
      'H'
    )
  ) >= WEBBOT_HORA_CORTE;

}


function WEBBOT_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  return WEBBOT_addDias_(
    fecha,
    dia === 0
      ? -6
      : 1 - dia
  );

}


function WEBBOT_diaOperativoAnterior_(
  fecha
) {

  let r =
    WEBBOT_addDias_(
      fecha,
      -1
    );


  while (
    r.getDay() === 0
  ) {

    r =
      WEBBOT_addDias_(
        r,
        -1
      );

  }


  return r;

}


function WEBBOT_fechaLocal_(
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


function WEBBOT_addDias_(
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


  return WEBBOT_fechaLocal_(
    r.getFullYear(),
    r.getMonth(),
    r.getDate()
  );

}


function WEBBOT_fechaClampeada_(
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


  return WEBBOT_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimo
    )
  );

}


function WEBBOT_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBBOT_TZ,
    'yyyy-MM-dd'
  );

}


function WEBBOT_entreFechas_(
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


function WEBBOT_parseFecha_(
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

function WEBBOT_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBBOT_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        WEBBOT_normalizar_(
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


function WEBBOT_normalizar_(
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


function WEBBOT_val_(
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


function WEBBOT_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBBOT_numero_(
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


  const n =
    Number(
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
      )
    );


  return isFinite(
    n
  )
    ? n
    : 0;

}


function WEBBOT_promedio_(
  valores
) {

  if (
    !valores.length
  ) {
    return 0;
  }


  return WEBBOT_redondear_(
    valores.reduce(
      (a, b) =>
        a + b,
      0
    ) /
    valores.length
  );

}


function WEBBOT_redondear_(
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


function WEBBOT_nombreMes_(
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


function WEBBOT_nombreMesCorto_(
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


function WEBBOT_pctTexto_(
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


// ============================================================================
// DIAGNÓSTICO DE FECHAS DEL BOT
// ============================================================================

function WEB_getBotDiagnosticoFechas() {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const sh =
      ss.getSheetByName(
        WEBBOT_HOJA_ENTRADAS
      );


    if (!sh) {

      return {
        ok: false,
        mensaje: 'No existe ENTRADASWEBOOK.'
      };

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
        filas: 0,
        fechasValidas: 0,
        fechasInvalidas: 0
      };

    }


    const c =
      WEBBOT_columnas_(
        data[0]
      );


    let validas = 0;
    let invalidas = 0;

    let minima = null;
    let maxima = null;

    const meses = {};


    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const valorEntrada =
        WEBBOT_val_(
          data[i],
          c.entrada
        );


      const valorFecha =
        WEBBOT_val_(
          data[i],
          c.fecha
        );


      let fecha =
        WEBBOT_parseFecha_(
          valorEntrada
        );


      if (!fecha) {

        fecha =
          WEBBOT_parseFecha_(
            valorFecha
          );

      }


      if (!fecha) {

        invalidas++;
        continue;

      }


      validas++;


      const key =
        WEBBOT_fechaKey_(
          fecha
        );


      if (
        !minima ||
        key < minima
      ) {
        minima = key;
      }


      if (
        !maxima ||
        key > maxima
      ) {
        maxima = key;
      }


      const mes =
        key.substring(
          0,
          7
        );


      meses[mes] =
        (
          meses[mes] || 0
        ) + 1;

    }


    return {

      ok: true,

      filas:
        data.length - 1,

      fechasValidas:
        validas,

      fechasInvalidas:
        invalidas,

      fechaMinima:
        minima,

      fechaMaxima:
        maxima,

      meses:
        Object
          .keys(meses)
          .sort()
          .map(
            mes => ({
              mes: mes,
              filas: meses[mes]
            })
          )

    };


  } catch (error) {

    return {
      ok: false,
      mensaje: error.message
    };

  }

}
