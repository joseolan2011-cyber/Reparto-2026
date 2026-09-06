// ============================================================================
// WEB_Promociones.gs
// ATLAS REPARTO - ANÁLISIS DE PROMOCIONES
// ============================================================================
// SOLO LECTURA.
//
// FUENTES:
//   ENTREGAS_PROMO
//   VENTAS
//   PARAMETROS
//
// Analiza:
//   - entregas promocionales
//   - clientes impactados
//   - piezas entregadas
//   - códigos autorizados / pendientes / inválidos
//   - liquidaciones
//   - recompra dentro de 30 días
//   - desempeño por vendedor
//
// RECOMPRA 30 DÍAS:
//   Solo se evalúan promociones con al menos 30 días de antigüedad.
//   Se considera recompra cuando existe una venta posterior del mismo
//   teléfono dentro de los 30 días siguientes a la promoción.
//
// NO MODIFICA DATOS.
// ============================================================================


const WEBPR_TZ = 'America/Merida';
const WEBPR_HORA_CORTE = 19;

const WEBPR_HOJA_PROMOS = 'ENTREGAS_PROMO';
const WEBPR_HOJA_VENTAS = 'VENTAS';
const WEBPR_HOJA_PARAMETROS = 'PARAMETROS';


// ============================================================================
// FUNCIÓN PRINCIPAL
// ============================================================================

function WEB_getPromocionesDashboard(filtros) {

  try {

    filtros = filtros || {};


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const shPromos =
      ss.getSheetByName(
        WEBPR_HOJA_PROMOS
      );


    if (!shPromos) {

      throw new Error(
        'No existe la hoja ENTREGAS_PROMO.'
      );

    }


    // ------------------------------------------------------------------------
    // 1. Resolver periodo ANTES de tocar VENTAS.
    // ------------------------------------------------------------------------

    const ahora =
      new Date();


    const hoy =
      WEBPR_hoyLocal_(
        ahora
      );


    const periodo =
      WEBPR_getPeriodo_(
        filtros.periodo || 'mes',
        ahora
      );


    const limiteEvaluable =
      WEBPR_addDias_(
        hoy,
        -30
      );


    const mapaVendedores =
      WEBPR_getMapaVendedores_(
        ss
      );


    const data =
      shPromos
        .getDataRange()
        .getValues();


    if (
      !data ||
      data.length < 2
    ) {

      return {

        ok: true,
        sinDatos: true,
        mensaje: 'ENTREGAS_PROMO no contiene registros.'

      };

    }


    const c =
      WEBPR_columnas_(
        data[0]
      );


    // ------------------------------------------------------------------------
    // 2. Leer ENTREGAS_PROMO solamente.
    // ------------------------------------------------------------------------

    const registros = [];


    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const fila =
        data[i];


      let fecha =
        WEBPR_parseFecha_(
          WEBPR_val_(
            fila,
            c.fechaEntrega
          )
        );


      if (!fecha) {

        fecha =
          WEBPR_parseFecha_(
            WEBPR_val_(
              fila,
              c.fechaDia
            )
          );

      }


      if (!fecha) {
        continue;
      }


      const telefono =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.telefono
          )
        );


      const clienteKey =
        WEBPR_clienteKey_(
          telefono,
          i + 1
        );


      const idVenta =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.idVenta
          )
        );


      const vendedorId =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.vendedorId
          )
        );


      const vendedorFila =
        WEBPR_texto_(
          WEBPR_val_(
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


      const codigoEnviado =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.codigoEnviado
          )
        );


      const codigoRecibido =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.codigoRecibido
          )
        );


      const estatusHoja =
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            c.estatus
          )
        );


      const estatus =
        WEBPR_determinarEstatus_(
          estatusHoja,
          codigoEnviado,
          codigoRecibido
        );


      registros.push({

        idEntrega:
          WEBPR_texto_(
            WEBPR_val_(
              fila,
              c.idEntrega
            )
          ) ||
          'PROMO_ROW_' +
          (i + 1),

        idVenta:
          idVenta,

        vendedorId:
          vendedorId,

        vendedor:
          vendedor,

        telefono:
          telefono,

        clienteKey:
          clienteKey,

        cliente:
          WEBPR_texto_(
            WEBPR_val_(
              fila,
              c.nombreCliente
            )
          ) ||
          telefono ||
          'Sin nombre',

        piezas:
          WEBPR_numero_(
            WEBPR_val_(
              fila,
              c.piezas
            )
          ),

        fecha:
          fecha,

        fechaKey:
          WEBPR_fechaKey_(
            fecha
          ),

        codigoEnviado:
          codigoEnviado,

        codigoRecibido:
          codigoRecibido,

        estatus:
          estatus,

        liquidada:
          WEBPR_esVerdadero_(
            WEBPR_val_(
              fila,
              c.liquidada
            )
          ),

        evaluable30:
          fecha <=
          limiteEvaluable,

        recompra30:
          null,

        fechaRecompra:
          ''

      });

    }


    // ------------------------------------------------------------------------
    // 3. Aplicar filtro de vendedor y fechas.
    // ------------------------------------------------------------------------

    const filtrados =
      registros.filter(
        r =>
          WEBPR_cumpleFiltros_(
            r,
            filtros
          )
      );


    const actuales =
      filtrados.filter(
        r =>
          WEBPR_entreFechas_(
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
            WEBPR_entreFechas_(
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
            WEBPR_entreFechas_(
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


    // ------------------------------------------------------------------------
    // 4. RECOMPRA:
    //    Solo consultar VENTAS si alguno de los registros que realmente se
    //    utilizarán en pantalla/comparativo ya tiene 30 días cumplidos.
    // ------------------------------------------------------------------------

    const promosAAnalizar = [];


    const vistos =
      new Set();


    [
      actuales,
      actualesComp,
      anterioresComp
    ].forEach(
      lista => {

        lista.forEach(
          r => {

            if (
              !r.evaluable30
            ) {
              return;
            }


            const llave =
              r.idEntrega ||
              (
                r.clienteKey +
                '_' +
                r.fechaKey
              );


            if (
              vistos.has(
                llave
              )
            ) {
              return;
            }


            vistos.add(
              llave
            );


            promosAAnalizar.push(
              r
            );

          }
        );

      }
    );


    if (
      promosAAnalizar.length
    ) {

      const clavesClientes =
        new Set();


      let fechaMin =
        null;


      let fechaMax =
        null;


      promosAAnalizar.forEach(
        r => {

          if (
            r.telefono
          ) {

            clavesClientes.add(
              r.clienteKey
            );

          }


          if (
            !fechaMin ||
            r.fecha <
            fechaMin
          ) {

            fechaMin =
              r.fecha;

          }


          const limitePromo =
            WEBPR_addDias_(
              r.fecha,
              30
            );


          if (
            !fechaMax ||
            limitePromo >
            fechaMax
          ) {

            fechaMax =
              limitePromo;

          }

        }
      );


      if (
        clavesClientes.size
      ) {

        const ventasCliente =
          WEBPR_cargarVentasPorClienteFiltrado_(
            ss,
            clavesClientes,
            fechaMin,
            fechaMax
          );


        promosAAnalizar.forEach(
          r => {

            const recompra =
              WEBPR_buscarRecompra30_(
                ventasCliente[
                  r.clienteKey
                ] || [],
                r.fecha,
                r.idVenta
              );


            r.recompra30 =
              recompra.encontro;


            r.fechaRecompra =
              recompra.fecha
                ? WEBPR_fechaKey_(
                    recompra.fecha
                  )
                : '';

          }
        );

      }

      else {

        promosAAnalizar.forEach(
          r => {

            r.recompra30 =
              false;

          }
        );

      }

    }


    // ------------------------------------------------------------------------
    // 5. Métricas.
    // ------------------------------------------------------------------------

    const resumen =
      WEBPR_resumen_(
        actuales
      );


    const resumenCompActual =
      WEBPR_resumen_(
        actualesComp
      );


    const resumenCompAnterior =
      WEBPR_resumen_(
        anterioresComp
      );


    const kpis = {

      entregas:
        WEBPR_kpi_(
          resumen.entregas,
          resumenCompActual.entregas,
          resumenCompAnterior.entregas,
          periodo.comparacionHabilitada
        ),

      clientes:
        WEBPR_kpi_(
          resumen.clientes,
          resumenCompActual.clientes,
          resumenCompAnterior.clientes,
          periodo.comparacionHabilitada
        ),

      piezas:
        WEBPR_kpi_(
          resumen.piezas,
          resumenCompActual.piezas,
          resumenCompAnterior.piezas,
          periodo.comparacionHabilitada
        ),

      autorizadas:
        WEBPR_kpi_(
          resumen.autorizadas,
          resumenCompActual.autorizadas,
          resumenCompAnterior.autorizadas,
          periodo.comparacionHabilitada
        ),

      pendientes:
        WEBPR_kpi_(
          resumen.pendientes,
          resumenCompActual.pendientes,
          resumenCompAnterior.pendientes,
          periodo.comparacionHabilitada
        ),

      invalidas:
        WEBPR_kpi_(
          resumen.invalidas,
          resumenCompActual.invalidas,
          resumenCompAnterior.invalidas,
          periodo.comparacionHabilitada
        ),

      liquidacionPct: {

        actual:
          resumen.liquidacionPct,

        delta:
          periodo.comparacionHabilitada
            ? WEBPR_deltaNullable_(
                resumenCompActual.liquidacionPct,
                resumenCompAnterior.liquidacionPct
              )
            : null

      },

      recompraPct: {

        actual:
          resumen.recompraPct,

        evaluables:
          resumen.evaluables30,

        recompras:
          resumen.recompras30,

        delta:
          periodo.comparacionHabilitada
            ? WEBPR_deltaNullable_(
                resumenCompActual.recompraPct,
                resumenCompAnterior.recompraPct
              )
            : null

      }

    };


    const tendencia =
      WEBPR_tendencia_(
        actuales,
        periodo
      );


    const estatus =
      WEBPR_agruparEstatus_(
        actuales
      );


    const vendedores =
      WEBPR_vendedores_(
        actuales
      );


    const detalle =
      actuales
        .slice()
        .sort(
          (a, b) =>
            b.fecha.getTime() -
            a.fecha.getTime()
        )
        .slice(
          0,
          60
        )
        .map(
          r => ({

            idEntrega:
              r.idEntrega,

            cliente:
              r.cliente,

            telefono:
              r.telefono,

            vendedor:
              r.vendedor,

            fecha:
              r.fechaKey,

            piezas:
              r.piezas,

            estatus:
              r.estatus,

            liquidada:
              r.liquidada,

            evaluable30:
              r.evaluable30,

            recompra30:
              r.recompra30,

            fechaRecompra:
              r.fechaRecompra

          })
        );


    const insights =
      WEBPR_insights_(
        resumen,
        vendedores,
        periodo
      );


    const catalogos =
      WEBPR_catalogos_(
        registros
      );


    return {

      ok: true,
      sinDatos: false,

      generado:
        Utilities.formatDate(
          new Date(),
          WEBPR_TZ,
          'dd/MM/yyyy HH:mm'
        ),

      periodo:
        periodo,

      kpis:
        kpis,

      tendencia:
        tendencia,

      estatus:
        estatus,

      vendedores:
        vendedores,

      detalle:
        detalle,

      insights:
        insights,

      catalogos:
        catalogos,

      meta: {

        promocionesTotales:
          registros.length,

        promocionesPeriodo:
          actuales.length,

        promocionesEvaluadasRecompra:
          promosAAnalizar.length

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
// RESUMEN
// ============================================================================

function WEBPR_resumen_(
  registros
) {

  const clientes =
    new Set();


  let piezas = 0;
  let autorizadas = 0;
  let pendientes = 0;
  let invalidas = 0;
  let liquidadas = 0;

  let evaluables30 = 0;
  let recompras30 = 0;


  registros.forEach(
    r => {

      clientes.add(
        r.clienteKey
      );


      piezas +=
        r.piezas;


      if (
        r.estatus ===
        'Autorizado'
      ) {

        autorizadas++;

      }

      else if (
        r.estatus ===
        'Inválido'
      ) {

        invalidas++;

      }

      else {

        pendientes++;

      }


      if (
        r.liquidada
      ) {

        liquidadas++;

      }


      if (
        r.evaluable30
      ) {

        evaluables30++;


        if (
          r.recompra30
        ) {

          recompras30++;

        }

      }

    }
  );


  return {

    entregas:
      registros.length,

    clientes:
      clientes.size,

    piezas:
      WEBPR_redondear_(
        piezas
      ),

    autorizadas:
      autorizadas,

    pendientes:
      pendientes,

    invalidas:
      invalidas,

    liquidadas:
      liquidadas,

    liquidacionPct:
      registros.length
        ? WEBPR_redondear_(
            liquidadas /
            registros.length *
            100
          )
        : null,

    evaluables30:
      evaluables30,

    recompras30:
      recompras30,

    recompraPct:
      evaluables30
        ? WEBPR_redondear_(
            recompras30 /
            evaluables30 *
            100
          )
        : null

  };

}


// ============================================================================
// VENTAS Y RECOMPRA
// ============================================================================

function WEBPR_cargarVentasPorClienteFiltrado_(
  ss,
  clavesClientes,
  fechaMin,
  fechaMax
) {

  const sh =
    ss.getSheetByName(
      WEBPR_HOJA_VENTAS
    );


  if (
    !sh ||
    !clavesClientes ||
    !clavesClientes.size
  ) {
    return {};
  }


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


  const headers =
    data[0];


  const telCol =
    WEBPR_findHeader_(
      headers,
      [
        'telefono_cliente',
        'telefono cliente'
      ]
    );


  const fechaCol =
    WEBPR_findHeader_(
      headers,
      [
        'fecha'
      ]
    );


  const fechaHoraCol =
    WEBPR_findHeader_(
      headers,
      [
        'fecha_hora',
        'fecha hora'
      ]
    );


  const idVentaCol =
    WEBPR_findHeader_(
      headers,
      [
        'id_venta',
        'id venta'
      ]
    );


  const mapa = {};


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const fila =
      data[i];


    const telefono =
      WEBPR_texto_(
        WEBPR_val_(
          fila,
          telCol
        )
      );


    if (!telefono) {
      continue;
    }


    const key =
      WEBPR_clienteKey_(
        telefono,
        i + 1
      );


    if (
      !clavesClientes.has(
        key
      )
    ) {
      continue;
    }


    let fecha =
      WEBPR_parseFecha_(
        WEBPR_val_(
          fila,
          fechaCol
        )
      );


    if (!fecha) {

      fecha =
        WEBPR_parseFecha_(
          WEBPR_val_(
            fila,
            fechaHoraCol
          )
        );

    }


    if (!fecha) {
      continue;
    }


    if (
      fechaMin &&
      fecha <
      fechaMin
    ) {
      continue;
    }


    if (
      fechaMax &&
      fecha >
      fechaMax
    ) {
      continue;
    }


    if (
      !mapa[key]
    ) {
      mapa[key] = [];
    }


    mapa[key].push({

      fecha:
        fecha,

      idVenta:
        WEBPR_texto_(
          WEBPR_val_(
            fila,
            idVentaCol
          )
        )

    });

  }


  // Las listas son muy pequeñas al estar filtradas por teléfono y ventana.
  Object
    .keys(mapa)
    .forEach(
      key => {

        mapa[key].sort(
          (a, b) =>
            a.fecha.getTime() -
            b.fecha.getTime()
        );

      }
    );


  return mapa;

}


function WEBPR_buscarRecompra30_(
  ventas,
  fechaPromo,
  idVentaPromo
) {

  if (
    !ventas ||
    !ventas.length
  ) {

    return {
      encontro: false,
      fecha: null
    };

  }


  const limite =
    WEBPR_addDias_(
      fechaPromo,
      30
    );


  for (
    let i = 0;
    i < ventas.length;
    i++
  ) {

    const venta =
      ventas[i];


    if (
      idVentaPromo &&
      venta.idVenta &&
      String(
        venta.idVenta
      ) ===
      String(
        idVentaPromo
      )
    ) {
      continue;
    }


    if (
      venta.fecha >
      fechaPromo &&
      venta.fecha <=
      limite
    ) {

      return {
        encontro: true,
        fecha: venta.fecha
      };

    }

  }


  return {
    encontro: false,
    fecha: null
  };

}


// ============================================================================
// ESTATUS
// ============================================================================

function WEBPR_determinarEstatus_(
  estatusHoja,
  codigoEnviado,
  codigoRecibido
) {

  const estatus =
    WEBPR_normalizar_(
      estatusHoja
    );


  if (
    estatus ===
    'autorizado'
  ) {

    return 'Autorizado';

  }


  if (
    [
      'invalido',
      'invalidado'
    ].includes(
      estatus
    )
  ) {

    return 'Inválido';

  }


  if (
    estatus ===
    'enviado'
  ) {

    return 'Pendiente';

  }


  if (
    codigoEnviado
  ) {

    if (
      codigoRecibido
    ) {

      return String(
        codigoEnviado
      ).trim() ===
      String(
        codigoRecibido
      ).trim()
        ? 'Autorizado'
        : 'Inválido';

    }


    return 'Pendiente';

  }


  return 'Pendiente';

}


// ============================================================================
// TENDENCIA
// ============================================================================

function WEBPR_tendencia_(
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
        'mes'
      ) {

        key =
          r.fechaKey.substring(
            0,
            7
          );


        label =
          WEBPR_nombreMesCorto_(
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

          entregas:
            0,

          piezas:
            0,

          autorizadas:
            0

        };

      }


      mapa[key].entregas++;


      mapa[key].piezas +=
        r.piezas;


      if (
        r.estatus ===
        'Autorizado'
      ) {

        mapa[key].autorizadas++;

      }

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

        entregas:
          mapa[key].entregas,

        piezas:
          WEBPR_redondear_(
            mapa[key].piezas
          ),

        autorizadas:
          mapa[key].autorizadas

      })
    );

}


// ============================================================================
// ESTATUS AGRUPADOS
// ============================================================================

function WEBPR_agruparEstatus_(
  registros
) {

  const mapa = {
    Autorizado: 0,
    Pendiente: 0,
    Inválido: 0
  };


  registros.forEach(
    r => {

      if (
        mapa[
          r.estatus
        ] ===
        undefined
      ) {

        mapa[
          r.estatus
        ] = 0;

      }


      mapa[
        r.estatus
      ]++;

    }
  );


  const total =
    registros.length;


  return Object
    .keys(mapa)
    .map(
      estatus => ({

        estatus:
          estatus,

        cantidad:
          mapa[estatus],

        porcentaje:
          total
            ? WEBPR_redondear_(
                mapa[estatus] /
                total *
                100
              )
            : 0

      })
    );

}


// ============================================================================
// VENDEDORES
// ============================================================================

function WEBPR_vendedores_(
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

          entregas:
            0,

          clientes:
            new Set(),

          piezas:
            0,

          autorizadas:
            0,

          invalidas:
            0,

          liquidadas:
            0,

          evaluables30:
            0,

          recompras30:
            0

        };

      }


      const item =
        mapa[key];


      item.entregas++;


      item.clientes.add(
        r.clienteKey
      );


      item.piezas +=
        r.piezas;


      if (
        r.estatus ===
        'Autorizado'
      ) {
        item.autorizadas++;
      }


      if (
        r.estatus ===
        'Inválido'
      ) {
        item.invalidas++;
      }


      if (
        r.liquidada
      ) {
        item.liquidadas++;
      }


      if (
        r.evaluable30
      ) {

        item.evaluables30++;


        if (
          r.recompra30
        ) {
          item.recompras30++;
        }

      }

    }
  );


  return Object
    .values(mapa)
    .map(
      x => ({

        vendedor:
          x.vendedor,

        entregas:
          x.entregas,

        clientes:
          x.clientes.size,

        piezas:
          WEBPR_redondear_(
            x.piezas
          ),

        autorizadas:
          x.autorizadas,

        invalidas:
          x.invalidas,

        liquidacionPct:
          x.entregas
            ? WEBPR_redondear_(
                x.liquidadas /
                x.entregas *
                100
              )
            : 0,

        recompraPct:
          x.evaluables30
            ? WEBPR_redondear_(
                x.recompras30 /
                x.evaluables30 *
                100
              )
            : null,

        evaluables30:
          x.evaluables30

      })
    )
    .sort(
      (a, b) =>
        b.entregas -
        a.entregas
    );

}


// ============================================================================
// INSIGHTS
// ============================================================================

function WEBPR_insights_(
  resumen,
  vendedores,
  periodo
) {

  const resultado = [];


  if (
    resumen.entregas
  ) {

    const autorizacion =
      resumen.autorizadas /
      resumen.entregas *
      100;


    resultado.push({

      tipo:
        autorizacion >= 85
          ? 'positivo'
          : autorizacion >= 65
            ? 'info'
            : 'advertencia',

      titulo:
        'Promociones autorizadas',

      texto:
        WEBPR_pctTexto_(
          autorizacion
        ) +
        ' de las entregas promocionales aparecen autorizadas.'

    });

  }


  if (
    resumen.invalidas
  ) {

    resultado.push({

      tipo:
        'advertencia',

      titulo:
        'Códigos inválidos',

      texto:
        resumen.invalidas +
        ' promociones del periodo tienen un código recibido diferente al enviado.'

    });

  }


  if (
    resumen.liquidacionPct !==
    null
  ) {

    resultado.push({

      tipo:
        resumen.liquidacionPct >= 90
          ? 'positivo'
          : resumen.liquidacionPct >= 70
            ? 'info'
            : 'advertencia',

      titulo:
        'Liquidación de promociones',

      texto:
        WEBPR_pctTexto_(
          resumen.liquidacionPct
        ) +
        ' de las entregas del periodo aparecen liquidadas.'

    });

  }


  if (
    resumen.recompraPct !==
    null
  ) {

    resultado.push({

      tipo:
        resumen.recompraPct >= 50
          ? 'positivo'
          : resumen.recompraPct >= 30
            ? 'info'
            : 'advertencia',

      titulo:
        'Recompra a 30 días',

      texto:
        WEBPR_pctTexto_(
          resumen.recompraPct
        ) +
        ' de las promociones con ventana completa de 30 días generaron una compra posterior.'

    });

  }

  else {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Recompra todavía no evaluable',

      texto:
        'Las promociones de este periodo aún no tienen 30 días completos para medir su recompra sin sesgo.'

    });

  }


  if (
    vendedores.length
  ) {

    const top =
      vendedores[0];


    resultado.push({

      tipo:
        'info',

      titulo:
        'Mayor actividad promocional',

      texto:
        top.vendedor +
        ' registra ' +
        top.entregas +
        ' entregas promocionales en el periodo.'

    });

  }


  if (
    !periodo.comparacionHabilitada
  ) {

    resultado.push({

      tipo:
        'info',

      titulo:
        'Periodo todavía abierto',

      texto:
        'La comparación del día actual se habilita después de las 7:00 p. m.'

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

function WEBPR_catalogos_(
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

function WEBPR_cumpleFiltros_(
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

function WEBPR_columnas_(
  headers
) {

  return {

    idEntrega:
      WEBPR_findHeader_(
        headers,
        [
          'id_entrega'
        ]
      ),

    vendedorId:
      WEBPR_findHeader_(
        headers,
        [
          'id_vendedor'
        ]
      ),

    idVenta:
      WEBPR_findHeader_(
        headers,
        [
          'id_venta'
        ]
      ),

    telefono:
      WEBPR_findHeader_(
        headers,
        [
          'telefono_cliente',
          'telefono cliente'
        ]
      ),

    nombreCliente:
      WEBPR_findHeader_(
        headers,
        [
          'nombre_cliente',
          'nombre cliente'
        ]
      ),

    piezas:
      WEBPR_findHeader_(
        headers,
        [
          'piezas_entregadas',
          'piezas entregadas'
        ]
      ),

    fechaEntrega:
      WEBPR_findHeader_(
        headers,
        [
          'fecha_entrega',
          'fecha entrega'
        ]
      ),

    fechaDia:
      WEBPR_findHeader_(
        headers,
        [
          'fecha_dia',
          'fecha dia'
        ]
      ),

    vendedorNombre:
      WEBPR_findHeader_(
        headers,
        [
          'vendedor'
        ]
      ),

    codigoEnviado:
      WEBPR_findHeader_(
        headers,
        [
          'codigo_enviado',
          'codigo enviado'
        ]
      ),

    codigoRecibido:
      WEBPR_findHeader_(
        headers,
        [
          'codigo_recibido',
          'codigo recibido'
        ]
      ),

    estatus:
      WEBPR_findHeader_(
        headers,
        [
          'estatus',
          'estado'
        ]
      ),

    liquidada:
      WEBPR_findHeader_(
        headers,
        [
          'liquidada',
          'liquidado'
        ]
      )

  };

}


// ============================================================================
// PARAMETROS
// ============================================================================

function WEBPR_getMapaVendedores_(
  ss
) {

  const resultado = {};


  const sh =
    ss.getSheetByName(
      WEBPR_HOJA_PARAMETROS
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
    WEBPR_findHeader_(
      data[0],
      [
        'id_vendedor'
      ]
    );


  const nombreCol =
    WEBPR_findHeader_(
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
      WEBPR_texto_(
        data[i][idCol]
      );


    const nombre =
      WEBPR_texto_(
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
// KPI
// ============================================================================

function WEBPR_kpi_(
  actualVisible,
  actualComparable,
  anteriorComparable,
  habilitada
) {

  return {

    actual:
      WEBPR_redondear_(
        actualVisible
      ),

    variacion:
      habilitada
        ? WEBPR_variacion_(
            actualComparable,
            anteriorComparable
          )
        : null

  };

}


function WEBPR_variacion_(
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


  return WEBPR_redondear_(
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


function WEBPR_deltaNullable_(
  actual,
  anterior
) {

  if (
    actual === null ||
    anterior === null
  ) {
    return null;
  }


  return WEBPR_redondear_(
    actual -
    anterior
  );

}


// ============================================================================
// PERIODOS
// ============================================================================

function WEBPR_getPeriodo_(
  tipo,
  ahora
) {

  tipo =
    String(
      tipo || 'mes'
    ).toLowerCase();


  const hoy =
    WEBPR_hoyLocal_(
      ahora
    );


  const corte =
    WEBPR_corteAlcanzado_(
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


    if (
      corte
    ) {

      compActualInicio =
        hoy;

      compActualFin =
        hoy;


      const anterior =
        WEBPR_diaOperativoAnterior_(
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
      WEBPR_inicioSemanaLunes_(
        hoy
      );


    actualInicio =
      WEBPR_addDias_(
        lunesActual,
        -7
      );


    actualFin =
      WEBPR_addDias_(
        actualInicio,
        5
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBPR_addDias_(
        actualInicio,
        -7
      );


    compAnteriorFin =
      WEBPR_addDias_(
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
      WEBPR_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualInicio =
      WEBPR_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 1,
        1
      );


    actualFin =
      WEBPR_addDias_(
        primerMesActual,
        -1
      );


    compActualInicio =
      actualInicio;

    compActualFin =
      actualFin;


    compAnteriorInicio =
      WEBPR_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth() - 2,
        1
      );


    compAnteriorFin =
      WEBPR_addDias_(
        actualInicio,
        -1
      );


    etiqueta =
      WEBPR_nombreMes_(
        actualInicio.getMonth()
      ) +
      ' ' +
      actualInicio.getFullYear();


    etiquetaAnterior =
      WEBPR_nombreMes_(
        compAnteriorInicio.getMonth()
      ) +
      ' ' +
      compAnteriorInicio.getFullYear();

  }

  else if (
    tipo === 'anio'
  ) {

    actualInicio =
      WEBPR_fechaLocal_(
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
        : WEBPR_addDias_(
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
        WEBPR_fechaLocal_(
          hoy.getFullYear() - 1,
          0,
          1
        );


      compAnteriorFin =
        WEBPR_fechaClampeada_(
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
      WEBPR_fechaLocal_(
        hoy.getFullYear(),
        hoy.getMonth(),
        1
      );


    actualFin =
      hoy;


    etiqueta =
      WEBPR_nombreMes_(
        hoy.getMonth()
      ) +
      ' ' +
      hoy.getFullYear();


    const ultimoComparable =
      corte
        ? hoy
        : WEBPR_addDias_(
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
        WEBPR_fechaLocal_(
          hoy.getFullYear(),
          hoy.getMonth() - 1,
          1
        );


      compAnteriorInicio =
        mesAnterior;


      compAnteriorFin =
        WEBPR_fechaClampeada_(
          mesAnterior.getFullYear(),
          mesAnterior.getMonth(),
          ultimoComparable.getDate()
        );


      etiquetaAnterior =
        WEBPR_nombreMes_(
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
      WEBPR_fechaKey_(
        actualInicio
      ),

    actualFinKey:
      WEBPR_fechaKey_(
        actualFin
      ),

    comparacionActualInicioKey:
      comparacionHabilitada
        ? WEBPR_fechaKey_(
            compActualInicio
          )
        : '',

    comparacionActualFinKey:
      comparacionHabilitada
        ? WEBPR_fechaKey_(
            compActualFin
          )
        : '',

    comparacionAnteriorInicioKey:
      comparacionHabilitada
        ? WEBPR_fechaKey_(
            compAnteriorInicio
          )
        : '',

    comparacionAnteriorFinKey:
      comparacionHabilitada
        ? WEBPR_fechaKey_(
            compAnteriorFin
          )
        : ''

  };

}


// ============================================================================
// FECHAS
// ============================================================================

function WEBPR_hoyLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      WEBPR_TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


  return WEBPR_fechaLocal_(
    Number(p[0]),
    Number(p[1]) - 1,
    Number(p[2])
  );

}


function WEBPR_corteAlcanzado_(
  fecha
) {

  return Number(
    Utilities.formatDate(
      fecha,
      WEBPR_TZ,
      'H'
    )
  ) >= WEBPR_HORA_CORTE;

}


function WEBPR_inicioSemanaLunes_(
  fecha
) {

  const dia =
    fecha.getDay();


  return WEBPR_addDias_(
    fecha,
    dia === 0
      ? -6
      : 1 - dia
  );

}


function WEBPR_diaOperativoAnterior_(
  fecha
) {

  let r =
    WEBPR_addDias_(
      fecha,
      -1
    );


  while (
    r.getDay() === 0
  ) {

    r =
      WEBPR_addDias_(
        r,
        -1
      );

  }


  return r;

}


function WEBPR_fechaLocal_(
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


function WEBPR_addDias_(
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


  return WEBPR_fechaLocal_(
    r.getFullYear(),
    r.getMonth(),
    r.getDate()
  );

}


function WEBPR_fechaClampeada_(
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


  return WEBPR_fechaLocal_(
    year,
    month,
    Math.min(
      day,
      ultimo
    )
  );

}


function WEBPR_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    WEBPR_TZ,
    'yyyy-MM-dd'
  );

}


function WEBPR_entreFechas_(
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


function WEBPR_parseFecha_(
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

function WEBPR_clienteKey_(
  telefono,
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


  return tel
    ? 'T:' + tel
    : 'R:' + fallback;

}


function WEBPR_esVerdadero_(
  valor
) {

  if (
    valor === true
  ) {
    return true;
  }


  const n =
    WEBPR_normalizar_(
      valor
    );


  return [
    'true',
    'si',
    '1',
    'yes',
    'liquidada',
    'liquidado'
  ].includes(
    n
  );

}


function WEBPR_findHeader_(
  headers,
  nombres
) {

  const normalizados =
    headers.map(
      WEBPR_normalizar_
    );


  for (
    let i = 0;
    i < nombres.length;
    i++
  ) {

    const idx =
      normalizados.indexOf(
        WEBPR_normalizar_(
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


function WEBPR_normalizar_(
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


function WEBPR_val_(
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


function WEBPR_texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


function WEBPR_numero_(
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


function WEBPR_redondear_(
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


function WEBPR_nombreMes_(
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


function WEBPR_nombreMesCorto_(
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


function WEBPR_pctTexto_(
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
// DIAGNÓSTICO RÁPIDO DE PROMOCIONES
// ============================================================================

function WEB_getPromocionesDiagnostico() {

  try {

    const inicio =
      new Date().getTime();


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const shPromo =
      ss.getSheetByName(
        WEBPR_HOJA_PROMOS
      );


    const shVentas =
      ss.getSheetByName(
        WEBPR_HOJA_VENTAS
      );


    return {

      ok: true,

      promociones:
        shPromo
          ? Math.max(
              0,
              shPromo.getLastRow() - 1
            )
          : 0,

      ventas:
        shVentas
          ? Math.max(
              0,
              shVentas.getLastRow() - 1
            )
          : 0,

      ms:
        new Date().getTime() -
        inicio

    };


  } catch (error) {

    return {
      ok: false,
      mensaje: error.message
    };

  }

}
