// ============================================================================
// SYNC_CARGAS_OPSU.gs
// REPARTO <- OPSU
// VERSION 1.9 - MONITOR CENTRAL + REINTENTO RAPIDO DE FILAS INCOMPLETAS
// ============================================================================
//
// OPSU: SOLO LECTURA.
// REPARTO:
//   - Lee PARAMETROS.
//   - Lee VENTAS.
//   - Escribe únicamente en CARGAS_OPSU.
//
// FOLIO = llave entre OPSU y Reparto.
// VENTAS[IDCarga_ref] debe contener el Folio de OPSU.
//
// Además de los datos de carga, esta versión escribe físicamente:
//   Agua_Venta_Reparto
//   SuperIce_Venta_Reparto
//   Comision_Reparto
//   Agua_Diferencia
//   SuperIce_Diferencia
//   Cuadre
//
// Diferencia = Venta_OPSU - Venta_Reparto
//
// Cuadre:
//   EXTERNO     -> no participa en conciliación Reparto
//   CUADRA      -> Agua y SuperIce tienen diferencia 0
//   DIFERENCIA  -> al menos uno no coincide
//
// FOLIO = llave entre OPSU y Reparto.
//
// Agua 20 L:
//   58, 52, 13, 55, 5, 53
//
// SuperIce:
//   61, 63, 43
//
// VACIOS: NO SE SINCRONIZAN.
//
// Tipo_Operacion:
//   Reparto  -> COD existe en PARAMETROS
//   Externo  -> COD no existe en PARAMETROS
//
// ============================================================================


const SYNC_OPSU = {

  PROP_ID:
    'CARGAS_OPSU_SPREADSHEET_ID',

  PROP_ID_REPARTO:
    'CARGAS_OPSU_REPARTO_SPREADSHEET_ID',

  TZ:
    'America/Merida',

  HOJA_TRANS:
    'Transacciones',

  HOJA_DETALLE:
    'Detalle_Transac',

  HOJA_DESTINO:
    'CARGAS_OPSU',

  HOJA_PARAMETROS:
    'PARAMETROS',

  HOJA_VENTAS:
    'VENTAS',

  COD_AGUA:
    [58, 52, 13, 55, 5, 53],

  COD_SUPERICE:
    [61, 63, 43],

  CACHE_PARAMETROS:
    'CARGAS_OPSU_MAPA_PARAMETROS_V1',

  CACHE_SEGUNDOS:
    50,

  // Si AppSheet de OPSU acaba de crear la transacción pero todavía no terminó
  // de llenar Fecha / FechaMedidor, no esperamos obligatoriamente otro minuto.
  // Reintentamos dentro de la MISMA ejecución.
  REINTENTOS_FUENTE:
    3,

  PAUSA_REINTENTO_MS:
    1500

};


// ============================================================================
// CONFIGURAR CONEXION CON OPSU
// ============================================================================

function CARGAS_OPSU_configurarConexion() {

  const ui =
    SpreadsheetApp.getUi();


  const r =
    ui.prompt(
      'Conectar Reparto con OPSU',
      'Pega la URL completa del Google Sheet de OPSU.',
      ui.ButtonSet.OK_CANCEL
    );


  if (
    r.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }


  const id =
    CARGAS_OPSU_extraerId_(
      r.getResponseText()
    );


  if (!id) {

    ui.alert(
      'No pude obtener el ID del archivo.'
    );

    return;

  }


  const ss =
    SpreadsheetApp.openById(
      id
    );


  if (
    !ss.getSheetByName(
      SYNC_OPSU.HOJA_TRANS
    )
  ) {

    throw new Error(
      'No existe la hoja Transacciones en OPSU.'
    );

  }


  if (
    !ss.getSheetByName(
      SYNC_OPSU.HOJA_DETALLE
    )
  ) {

    throw new Error(
      'No existe la hoja Detalle_Transac en OPSU.'
    );

  }


  const ssReparto =
    SpreadsheetApp.getActiveSpreadsheet();


  PropertiesService
    .getScriptProperties()
    .setProperties(
      {
        [SYNC_OPSU.PROP_ID]:
          id,

        [SYNC_OPSU.PROP_ID_REPARTO]:
          ssReparto.getId()
      },
      false
    );


  ui.alert(
    'Conexión correcta.\n\nOPSU: ' +
    ss.getName()
  );

}


// ============================================================================
// BORRAR CACHE DE PARAMETROS
// Ejecutar solo si cambias COD / vendedores y quieres refrescar al instante.
// ============================================================================

function CARGAS_OPSU_refrescarVendedores() {

  CacheService
    .getScriptCache()
    .remove(
      SYNC_OPSU.CACHE_PARAMETROS
    );


  const mapa =
    CARGAS_OPSU_getMapaVendedores_(
      true
    );


  SpreadsheetApp
    .getUi()
    .alert(
      'Mapa de vendedores actualizado y cacheado por 50 segundos.\n\n' +
      'Vínculos encontrados: ' +
      Object.keys(mapa).length
    );

}


// ============================================================================
// PRUEBA - NO ESCRIBE NADA
// ============================================================================

function CARGAS_OPSU_probarSincronizacion() {

  const inicio =
    Date.now();


  const datos =
    CARGAS_OPSU_leerCargas_(
      true,
      new Set()
    );


  const agua =
    datos.reduce(
      (s, x) =>
        s +
        x.Agua_Cargada,
      0
    );


  const hielo =
    datos.reduce(
      (s, x) =>
        s +
        x.SuperIce_Cargado,
      0
    );


  let texto =
    'PRUEBA OPSU → REPARTO\n\n' +
    'Cargas: ' +
    datos.length +
    '\n' +
    'Agua: ' +
    agua +
    '\n' +
    'SuperIce: ' +
    hielo +
    '\n' +
    'Tiempo total: ' +
    (
      (
        Date.now() -
        inicio
      ) /
      1000
    ).toFixed(1) +
    ' s\n\n';


  datos.forEach(
    x => {

      texto +=
        'Moto ' +
        x.Moto +
        ' | Carga ' +
        x.No_Carga +
        ' | Agua ' +
        x.Agua_Cargada +
        ' | Ice ' +
        x.SuperIce_Cargado +
        ' | COD ' +
        x.COD +
        ' | ' +
        x.Tipo_Operacion +
        '\n';

    }
  );


  texto +=
    '\nNO SE ESCRIBIÓ NADA.';


  SpreadsheetApp
    .getUi()
    .alert(
      texto
    );


  console.log(
    texto
  );

}


// ============================================================================
// SINCRONIZACION REAL
// ============================================================================

function CARGAS_OPSU_sincronizarCargasCore_(mostrarUI) {

  const inicio =
    Date.now();


  const lock =
    LockService.getUserLock();


  if (
    !lock.tryLock(
      1000
    )
  ) {

    console.log(
      'CARGAS OPSU: otra sincronización CARGAS_OPSU está activa. Este ciclo se omite.'
    );


    return {
      omitida:
        true,

      motivo:
        'LOCK_OCUPADO',

      segundos:
        (
          Date.now() -
          inicio
        ) /
        1000
    };

  }


  try {

    const ssDestino =
      CARGAS_OPSU_getSpreadsheetReparto_();


    const sh =
      ssDestino.getSheetByName(
        SYNC_OPSU.HOJA_DESTINO
      );


    if (!sh) {

      throw new Error(
        'No existe CARGAS_OPSU en Reparto.'
      );

    }


    const lastRow =
      Math.max(
        sh.getLastRow(),
        1
      );


    const lastCol =
      sh.getLastColumn();


    if (!lastCol) {

      throw new Error(
        'CARGAS_OPSU no tiene encabezados.'
      );

    }


    const tabla =
      sh
        .getRange(
          1,
          1,
          lastRow,
          lastCol
        )
        .getValues();


    const h =
      CARGAS_OPSU_headers_(
        tabla[0]
      );


    CARGAS_OPSU_validarDestino_(
      h
    );


    const cFolio =
      CARGAS_OPSU_col_(
        h,
        'Folio'
      );


    const cEstado =
      CARGAS_OPSU_col_(
        h,
        'Estatus_OPSU'
      );


    const pendientes =
      new Set();


    for (
      let i = 1;
      i < tabla.length;
      i++
    ) {

      const folio =
        CARGAS_OPSU_txt_(
          tabla[i][
            cFolio
          ]
        );


      const estatus =
        CARGAS_OPSU_norm_(
          tabla[i][
            cEstado
          ]
        );


      if (
        folio &&
        estatus !==
        'cerrado'
      ) {

        pendientes.add(
          folio
        );

      }

    }


    let cargas =
      CARGAS_OPSU_leerCargas_(
        false,
        pendientes
      );


    for (
      let intento = 1;
      intento <=
        SYNC_OPSU.REINTENTOS_FUENTE;
      intento++
    ) {

      const nuevasIncompletas =
        cargas.filter(
          carga =>
            !carga.Fuente_Completa &&
            !pendientes.has(
              CARGAS_OPSU_txt_(
                carga.Folio
              )
            )
        );


      if (
        !nuevasIncompletas.length
      ) {
        break;
      }


      console.log(
        'CARGAS OPSU: ' +
        nuevasIncompletas.length +
        ' carga(s) nueva(s) todavía incompleta(s). ' +
        'Reintento ' +
        intento +
        '/' +
        SYNC_OPSU.REINTENTOS_FUENTE +
        ' en ' +
        SYNC_OPSU.PAUSA_REINTENTO_MS +
        ' ms.'
      );


      Utilities.sleep(
        SYNC_OPSU.PAUSA_REINTENTO_MS
      );


      cargas =
        CARGAS_OPSU_leerCargas_(
          false,
          pendientes
        );

    }


    const foliosCiclo =
      new Set(
        cargas.map(
          x =>
            CARGAS_OPSU_txt_(
              x.Folio
            )
        )
      );


    const ventasPorFolio =
      CARGAS_OPSU_leerVentasReparto_(
        foliosCiclo
      );


    cargas.forEach(
      carga => {

        const venta =
          ventasPorFolio[
            carga.Folio
          ] ||
          {
            agua: 0,
            superice: 0,
            comision: 0
          };


        carga.Agua_Venta_Reparto =
          CARGAS_OPSU_num_(
            venta.agua
          );


        carga.SuperIce_Venta_Reparto =
          CARGAS_OPSU_num_(
            venta.superice
          );


        carga.Comision_Reparto =
          CARGAS_OPSU_num_(
            venta.comision
          );


        carga.Agua_Diferencia =
          carga.Agua_Venta_OPSU -
          carga.Agua_Venta_Reparto;


        carga.SuperIce_Diferencia =
          carga.SuperIce_Venta_OPSU -
          carga.SuperIce_Venta_Reparto;


        if (
          carga.Tipo_Operacion ===
          'Externo'
        ) {

          carga.Cuadre =
            'EXTERNO';

        }

        else if (
          carga.Agua_Diferencia ===
            0 &&
          carga.SuperIce_Diferencia ===
            0
        ) {

          carga.Cuadre =
            'CUADRA';

        }

        else {

          carga.Cuadre =
            'DIFERENCIA';

        }

      }
    );


    const porFolio = {};


    for (
      let i = 1;
      i < tabla.length;
      i++
    ) {

      const folio =
        CARGAS_OPSU_txt_(
          tabla[i][
            cFolio
          ]
        );


      if (folio) {

        porFolio[
          folio
        ] =
          i;

      }

    }


    let creados = 0;
    let actualizados = 0;


    cargas.forEach(
      carga => {

        let pos =
          porFolio[
            carga.Folio
          ];


        let fila;


        if (
          pos ===
          undefined
        ) {

          if (
            !carga.Fuente_Completa
          ) {

            console.warn(
              'CARGAS OPSU: folio ' +
              carga.Folio +
              ' todavía incompleto en OPSU. Se omite hasta el próximo ciclo.'
            );


            return;

          }


          fila =
            new Array(
              lastCol
            ).fill(
              ''
            );


          tabla.push(
            fila
          );


          pos =
            tabla.length -
            1;


          porFolio[
            carga.Folio
          ] =
            pos;


          creados++;

        }

        else {

          fila =
            tabla[
              pos
            ];


          actualizados++;

        }


        CARGAS_OPSU_set_(
          fila,
          h,
          'Folio',
          carga.Folio
        );


        const cFechaCargaDestino =
          CARGAS_OPSU_col_(
            h,
            'Fecha_Carga_OPSU'
          );


        const cFechaOperacionDestino =
          CARGAS_OPSU_col_(
            h,
            'Fecha_Operacion'
          );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Fecha_Carga_OPSU',
          carga.Fecha_Carga_OPSU ||
          fila[
            cFechaCargaDestino
          ] ||
          ''
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Fecha_Operacion',
          carga.Fecha_Operacion ||
          fila[
            cFechaOperacionDestino
          ] ||
          ''
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Moto',
          carga.Moto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'No_Carga',
          carga.No_Carga
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'COD',
          carga.COD
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'id_vendedor',
          carga.id_vendedor
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Tipo_Operacion',
          carga.Tipo_Operacion
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Ruta_Reparto',
          carga.Ruta_Reparto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Estatus_OPSU',
          carga.Estatus_OPSU
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'CargasContinuas',
          carga.CargasContinuas
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Agua_Cargada',
          carga.Agua_Cargada
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Agua_Devuelta',
          carga.Agua_Devuelta
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Agua_Venta_OPSU',
          carga.Agua_Venta_OPSU
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'SuperIce_Cargado',
          carga.SuperIce_Cargado
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'SuperIce_Devuelto',
          carga.SuperIce_Devuelto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'SuperIce_Venta_OPSU',
          carga.SuperIce_Venta_OPSU
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Agua_Venta_Reparto',
          carga.Agua_Venta_Reparto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'SuperIce_Venta_Reparto',
          carga.SuperIce_Venta_Reparto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Comision_Reparto',
          carga.Comision_Reparto
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Agua_Diferencia',
          carga.Agua_Diferencia
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'SuperIce_Diferencia',
          carga.SuperIce_Diferencia
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'Cuadre',
          carga.Cuadre
        );


        CARGAS_OPSU_set_(
          fila,
          h,
          'FechaHora_Sync',
          new Date()
        );

      }
    );


    CARGAS_OPSU_recalcularVentasFisicasTabla_(
      tabla,
      h
    );


    if (
      tabla.length >
      1
    ) {

      sh
        .getRange(
          1,
          1,
          tabla.length,
          lastCol
        )
        .setValues(
          tabla
        );

    }


    const segundos =
      (
        (
          Date.now() -
          inicio
        ) /
        1000
      ).toFixed(
        1
      );


    const resultado = {

      ok:
        true,

      encontrados:
        cargas.length,

      creados:
        creados,

      actualizados:
        actualizados,

      folios_conciliados:
        cargas.filter(
          x =>
            x.Tipo_Operacion ===
            'Reparto'
        ).length,

      folios_con_diferencia:
        cargas.filter(
          x =>
            x.Tipo_Operacion ===
              'Reparto' &&
            x.Cuadre ===
              'DIFERENCIA'
        ).length,

      segundos:
        segundos

    };


    console.log(
      JSON.stringify(
        resultado
      )
    );


    if (
      mostrarUI
    ) {

      SpreadsheetApp
        .getUi()
        .alert(
          'SINCRONIZACIÓN TERMINADA\n\n' +
          'Encontrados: ' +
          cargas.length +
          '\n' +
          'Nuevos: ' +
          creados +
          '\n' +
          'Actualizados: ' +
          actualizados +
          '\n' +
          'Tiempo: ' +
          segundos +
          ' s'
        );

    }


    return resultado;


  } finally {

    lock.releaseLock();

  }

}


function CARGAS_OPSU_sincronizarCargas() {

  return CARGAS_OPSU_sincronizarCargasCore_(
    true
  );

}


function CARGAS_OPSU_sincronizarSilenciosoAhora() {

  return CARGAS_OPSU_sincronizarCargasCore_(
    false
  );

}


function CARGAS_OPSU_monitoreoCadaMinuto() {

  try {

    const resultado =
      CARGAS_OPSU_sincronizarCargasCore_(
        false
      );


    console.log(
      'MONITOREO CARGAS OPSU: ' +
      JSON.stringify(
        resultado
      )
    );


    return resultado;


  } catch (error) {

    console.error(
      'ERROR MONITOREO CARGAS OPSU: ' +
      error.stack
    );


    return {
      error:
        true,

      mensaje:
        error.message
    };

  }

}


function MONITOREO_CENTRAL_cadaMinuto() {

  const inicio =
    Date.now();


  try {

    ejecutarMonitoreosCadaMinuto();

  } catch (error) {

    console.error(
      'MONITOR CENTRAL - error en monitoreos existentes: ' +
      error.stack
    );

  }


  try {

    const resultadoCargas =
      CARGAS_OPSU_monitoreoCadaMinuto();


    console.log(
      'MONITOR CENTRAL - resultado CARGAS OPSU: ' +
      JSON.stringify(
        resultadoCargas
      )
    );

  } catch (error) {

    console.error(
      'MONITOR CENTRAL - error en CARGAS OPSU: ' +
      error.stack
    );

  }


  console.log(
    'MONITOR CENTRAL terminado en ' +
    (
      (
        Date.now() -
        inicio
      ) /
      1000
    ).toFixed(
      1
    ) +
    ' s'
  );

}


function CARGAS_OPSU_instalarMonitoreo1Minuto() {

  const ssReparto =
    SpreadsheetApp.getActiveSpreadsheet();


  if (!ssReparto) {

    throw new Error(
      'Ejecuta esta función manualmente desde el proyecto ligado a Reparto.'
    );

  }


  PropertiesService
    .getScriptProperties()
    .setProperty(
      SYNC_OPSU.PROP_ID_REPARTO,
      ssReparto.getId()
    );


  ScriptApp
    .getProjectTriggers()
    .forEach(
      trigger => {

        const handler =
          trigger.getHandlerFunction();


        if (
          handler ===
            'ejecutarMonitoreosCadaMinuto' ||
          handler ===
            'CARGAS_OPSU_monitoreoCadaMinuto' ||
          handler ===
            'MONITOREO_CENTRAL_cadaMinuto'
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );

        }

      }
    );


  ScriptApp
    .newTrigger(
      'MONITOREO_CENTRAL_cadaMinuto'
    )
    .timeBased()
    .everyMinutes(
      1
    )
    .create();


  ScriptApp
    .getProjectTriggers()
    .forEach(
      trigger => {

        if (
          trigger.getHandlerFunction() ===
          'CARGAS_OPSU_alEditarParametros'
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );

        }

      }
    );


  ScriptApp
    .newTrigger(
      'CARGAS_OPSU_alEditarParametros'
    )
    .forSpreadsheet(
      ssReparto
    )
    .onEdit()
    .create();


  ssReparto.toast(
    'Monitor central único activado cada 1 minuto.',
    'MONITOREO',
    10
  );


  console.log(
    'Listo: quedó un solo trigger cada minuto para MONITOREO_CENTRAL_cadaMinuto.'
  );

}


function CARGAS_OPSU_eliminarMonitoreo1Minuto() {

  ScriptApp
    .getProjectTriggers()
    .forEach(
      trigger => {

        const handler =
          trigger.getHandlerFunction();


        if (
          handler ===
            'MONITOREO_CENTRAL_cadaMinuto' ||
          handler ===
            'CARGAS_OPSU_monitoreoCadaMinuto' ||
          handler ===
            'CARGAS_OPSU_alEditarParametros'
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );

        }

      }
    );


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  if (ss) {

    ss.toast(
      'Monitor central OPSU detenido.',
      'MONITOREO',
      8
    );

  }

}


function CARGAS_OPSU_repararFilasIncompletas() {

  const resultado =
    CARGAS_OPSU_sincronizarCargasCore_(
      false
    );


  console.log(
    'REPARACION CARGAS OPSU: ' +
    JSON.stringify(
      resultado
    )
  );


  return resultado;

}


function CARGAS_OPSU_recalcularVentasAhora() {

  const ss =
    CARGAS_OPSU_getSpreadsheetReparto_();


  const sh =
    ss.getSheetByName(
      SYNC_OPSU.HOJA_DESTINO
    );


  if (!sh) {

    throw new Error(
      'No existe la hoja CARGAS_OPSU.'
    );

  }


  const lastRow =
    sh.getLastRow();


  const lastCol =
    sh.getLastColumn();


  if (
    lastRow <
      2
  ) {

    console.log(
      'CARGAS_OPSU no tiene filas para recalcular.'
    );


    return;

  }


  const tabla =
    sh
      .getRange(
        1,
        1,
        lastRow,
        lastCol
      )
      .getValues();


  const h =
    CARGAS_OPSU_headers_(
      tabla[0]
    );


  CARGAS_OPSU_validarDestino_(
    h
  );


  CARGAS_OPSU_recalcularVentasFisicasTabla_(
    tabla,
    h
  );


  sh
    .getRange(
      1,
      1,
      tabla.length,
      lastCol
    )
    .setValues(
      tabla
    );


  console.log(
    'Recalculo físico de VENTAS -> CARGAS_OPSU completado.'
  );

}


function CARGAS_OPSU_probarConciliacionFisica() {

  const resultado =
    CARGAS_OPSU_sincronizarCargasCore_(
      false
    );


  console.log(
    'PRUEBA CONCILIACION FISICA: ' +
    JSON.stringify(
      resultado
    )
  );


  return resultado;

}


function CARGAS_OPSU_alEditarParametros(e) {

  try {

    if (
      !e ||
      !e.range
    ) {
      return;
    }


    const hoja =
      e.range.getSheet();


    if (
      hoja.getName() !==
      SYNC_OPSU.HOJA_PARAMETROS
    ) {
      return;
    }


    CacheService
      .getScriptCache()
      .remove(
        SYNC_OPSU.CACHE_PARAMETROS
      );


    console.log(
      'PARAMETROS cambió: cache COD -> id_vendedor invalidado.'
    );


  } catch (error) {

    console.error(
      'Error invalidando cache de PARAMETROS: ' +
      error.stack
    );

  }

}


function CARGAS_OPSU_getSpreadsheetReparto_() {

  const props =
    PropertiesService.getScriptProperties();


  let id =
    props.getProperty(
      SYNC_OPSU.PROP_ID_REPARTO
    );


  if (!id) {

    const activo =
      SpreadsheetApp.getActiveSpreadsheet();


    if (!activo) {

      throw new Error(
        'No está configurado el ID del archivo Reparto. ' +
        'Ejecuta CARGAS_OPSU_instalarMonitoreo1Minuto manualmente.'
      );

    }


    id =
      activo.getId();


    props.setProperty(
      SYNC_OPSU.PROP_ID_REPARTO,
      id
    );

  }


  return SpreadsheetApp.openById(
    id
  );

}


function CARGAS_OPSU_leerCargas_(
  soloActualesAbiertas,
  foliosPendientes
) {

  const id =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        SYNC_OPSU.PROP_ID
      );


  if (!id) {

    throw new Error(
      'Primero ejecuta CARGAS_OPSU_configurarConexion.'
    );

  }


  const ssOPSU =
    SpreadsheetApp.openById(
      id
    );


  const shT =
    ssOPSU.getSheetByName(
      SYNC_OPSU.HOJA_TRANS
    );


  const shD =
    ssOPSU.getSheetByName(
      SYNC_OPSU.HOJA_DETALLE
    );


  if (
    !shT ||
    !shD
  ) {

    throw new Error(
      'Falta Transacciones o Detalle_Transac en OPSU.'
    );

  }


  const emailPorCOD =
    CARGAS_OPSU_getMapaVendedores_(
      false
    );


  const rutaPorCOD =
    CARGAS_OPSU_getMapaRutas_();


  const tLastRow =
    shT.getLastRow();


  if (
    tLastRow <
    2
  ) {

    return [];

  }


  const t =
    shT
      .getRange(
        1,
        1,
        tLastRow,
        Math.min(
          shT.getLastColumn(),
          10
        )
      )
      .getValues();


  const ht =
    CARGAS_OPSU_headers_(
      t[0]
    );


  const tEstatus =
    CARGAS_OPSU_col_(
      ht,
      'Estatus'
    );


  const tFolio =
    CARGAS_OPSU_col_(
      ht,
      'Folio'
    );


  const tFecha =
    CARGAS_OPSU_col_(
      ht,
      'Fecha'
    );


  const tMoto =
    CARGAS_OPSU_col_(
      ht,
      'Ruta'
    );


  const tCarga =
    CARGAS_OPSU_colMulti_(
      ht,
      [
        'No. Carga',
        'No_Carga'
      ]
    );


  const tVendedor =
    CARGAS_OPSU_col_(
      ht,
      'Vendedor'
    );


  const tContinuas =
    CARGAS_OPSU_col_(
      ht,
      'CargasContinuas'
    );


  const tFechaMedidor =
    CARGAS_OPSU_col_(
      ht,
      'FechaMedidor'
    );


  const hoy =
    CARGAS_OPSU_fechaLocal_(
      new Date()
    );


  // Siguiente día operativo: el domingo no se trabaja.
  // Sábado -> lunes; domingo -> lunes; resto -> día siguiente.
  const manana =
    CARGAS_OPSU_addDias_(
      hoy,
      hoy.getDay() === 6
        ? 2
        : 1
    );


  const keyHoy =
    CARGAS_OPSU_fechaKey_(
      hoy
    );


  const keyManana =
    CARGAS_OPSU_fechaKey_(
      manana
    );


  const cargas = {};
  const folios =
    new Set();


  for (
    let i = 1;
    i < t.length;
    i++
  ) {

    const fila =
      t[i];


    const folio =
      CARGAS_OPSU_txt_(
        fila[
          tFolio
        ]
      );


    if (!folio) {
      continue;
    }


    const estado =
      CARGAS_OPSU_txt_(
        fila[
          tEstatus
        ]
      );


    const abierto =
      CARGAS_OPSU_norm_(
        estado
      ) ===
      'abierto';


    const fechaOperacion =
      fila[
        tFecha
      ];


    const keyOperacion =
      CARGAS_OPSU_fechaKeySeguro_(
        fechaOperacion
      );


    const esOperacionActual =
      (
        keyOperacion ===
        keyHoy
      ) ||
      (
        keyOperacion ===
        keyManana
      );


    const sinFechaOperacion =
      !keyOperacion;


    let incluir =
      false;


    if (
      soloActualesAbiertas
    ) {

      incluir =
        abierto &&
        (
          esOperacionActual ||
          sinFechaOperacion
        );

    }

    else {

      incluir =
        (
          abierto &&
          (
            esOperacionActual ||
            sinFechaOperacion
          )
        ) ||
        foliosPendientes.has(
          folio
        );

    }


    if (!incluir) {
      continue;
    }


    const codKey =
      CARGAS_OPSU_cod_(
        fila[
          tVendedor
        ]
      );


    const idVendedor =
      emailPorCOD[
        codKey
      ] ||
      '';


    const rutaReparto =
      rutaPorCOD[
        codKey
      ] ||
      '';


    cargas[
      folio
    ] = {

      Folio:
        folio,

      Fecha_Carga_OPSU:
        fila[
          tFechaMedidor
        ] ||
        '',

      Fecha_Operacion:
        fechaOperacion ||
        '',

      Fuente_Completa:
        Boolean(
          fila[
            tFechaMedidor
          ] &&
          fechaOperacion
        ),

      Moto:
        fila[
          tMoto
        ] ||
        '',

      No_Carga:
        CARGAS_OPSU_num_(
          fila[
            tCarga
          ]
        ),

      COD:
        fila[
          tVendedor
        ] ||
        '',

      id_vendedor:
        idVendedor,

      Ruta_Reparto:
        rutaReparto,

      Tipo_Operacion:
        idVendedor
          ? 'Reparto'
          : 'Externo',

      Estatus_OPSU:
        estado,

      CargasContinuas:
        fila[
          tContinuas
        ],

      Agua_Cargada:
        0,

      Agua_Devuelta:
        0,

      Agua_Venta_OPSU:
        0,

      SuperIce_Cargado:
        0,

      SuperIce_Devuelto:
        0,

      SuperIce_Venta_OPSU:
        0

    };


    folios.add(
      folio
    );

  }


  if (
    !folios.size
  ) {

    return [];

  }


  const dLastRow =
    shD.getLastRow();


  if (
    dLastRow >=
    2
  ) {

    const d =
      shD
        .getRange(
          1,
          2,
          dLastRow,
          10
        )
        .getValues();


    const DFOLIO = 0;
    const DCOD = 5;
    const DCARGA = 6;
    const DDEV = 7;
    const DVENTA = 9;


    const agua =
      new Set(
        SYNC_OPSU
          .COD_AGUA
          .map(
            String
          )
      );


    const ice =
      new Set(
        SYNC_OPSU
          .COD_SUPERICE
          .map(
            String
          )
      );


    for (
      let i = 1;
      i < d.length;
      i++
    ) {

      const folio =
        CARGAS_OPSU_txt_(
          d[i][
            DFOLIO
          ]
        );


      if (
        !folios.has(
          folio
        )
      ) {
        continue;
      }


      const cod =
        CARGAS_OPSU_cod_(
          d[i][
            DCOD
          ]
        );


      if (
        agua.has(
          cod
        )
      ) {

        cargas[
          folio
        ].Agua_Cargada +=
          CARGAS_OPSU_num_(
            d[i][
              DCARGA
            ]
          );


        cargas[
          folio
        ].Agua_Devuelta +=
          CARGAS_OPSU_num_(
            d[i][
              DDEV
            ]
          );


        cargas[
          folio
        ].Agua_Venta_OPSU +=
          CARGAS_OPSU_num_(
            d[i][
              DVENTA
            ]
          );

      }

      else if (
        ice.has(
          cod
        )
      ) {

        cargas[
          folio
        ].SuperIce_Cargado +=
          CARGAS_OPSU_num_(
            d[i][
              DCARGA
            ]
          );


        cargas[
          folio
        ].SuperIce_Devuelto +=
          CARGAS_OPSU_num_(
            d[i][
              DDEV
            ]
          );


        cargas[
          folio
        ].SuperIce_Venta_OPSU +=
          CARGAS_OPSU_num_(
            d[i][
              DVENTA
            ]
          );

      }

    }

  }


  return Object
    .values(
      cargas
    )
    .sort(
      (a, b) => {

        const ma =
          Number(
            a.Moto
          ) ||
          0;


        const mb =
          Number(
            b.Moto
          ) ||
          0;


        if (
          ma !== mb
        ) {

          return ma -
            mb;

        }


        return (
          Number(
            a.No_Carga ||
            0
          ) -
          Number(
            b.No_Carga ||
            0
          )
        );

      }
    );

}


function CARGAS_OPSU_recalcularVentasFisicasTabla_(
  tabla,
  h
) {

  if (
    !tabla ||
    tabla.length <
      2
  ) {

    return;

  }


  const cFolio =
    CARGAS_OPSU_col_(
      h,
      'Folio'
    );


  const cTipo =
    CARGAS_OPSU_col_(
      h,
      'Tipo_Operacion'
    );


  const cAguaOpsu =
    CARGAS_OPSU_col_(
      h,
      'Agua_Venta_OPSU'
    );


  const cIceOpsu =
    CARGAS_OPSU_col_(
      h,
      'SuperIce_Venta_OPSU'
    );


  const folios =
    new Set();


  for (
    let i = 1;
    i < tabla.length;
    i++
  ) {

    const folio =
      CARGAS_OPSU_txt_(
        tabla[i][
          cFolio
        ]
      );


    const tipo =
      CARGAS_OPSU_txt_(
        tabla[i][
          cTipo
        ]
      );


    if (
      folio &&
      tipo ===
        'Reparto'
    ) {

      folios.add(
        folio
      );

    }

  }


  const ventas =
    CARGAS_OPSU_leerVentasReparto_(
      folios
    );


  for (
    let i = 1;
    i < tabla.length;
    i++
  ) {

    const fila =
      tabla[i];


    const folio =
      CARGAS_OPSU_txt_(
        fila[
          cFolio
        ]
      );


    const tipo =
      CARGAS_OPSU_txt_(
        fila[
          cTipo
        ]
      );


    if (!folio) {
      continue;
    }


    if (
      tipo ===
      'Externo'
    ) {

      CARGAS_OPSU_set_(
        fila,
        h,
        'Cuadre',
        'EXTERNO'
      );


      continue;

    }


    if (
      tipo !==
      'Reparto'
    ) {

      continue;

    }


    const venta =
      ventas[
        folio
      ] ||
      {
        agua:
          0,

        superice:
          0,

        comision:
          0
      };


    const aguaReparto =
      CARGAS_OPSU_num_(
        venta.agua
      );


    const iceReparto =
      CARGAS_OPSU_num_(
        venta.superice
      );


    const comisionReparto =
      CARGAS_OPSU_num_(
        venta.comision
      );


    const aguaOpsu =
      CARGAS_OPSU_num_(
        fila[
          cAguaOpsu
        ]
      );


    const iceOpsu =
      CARGAS_OPSU_num_(
        fila[
          cIceOpsu
        ]
      );


    const diferenciaAgua =
      aguaOpsu -
      aguaReparto;


    const diferenciaIce =
      iceOpsu -
      iceReparto;


    CARGAS_OPSU_set_(
      fila,
      h,
      'Agua_Venta_Reparto',
      aguaReparto
    );


    CARGAS_OPSU_set_(
      fila,
      h,
      'SuperIce_Venta_Reparto',
      iceReparto
    );


    CARGAS_OPSU_set_(
      fila,
      h,
      'Comision_Reparto',
      comisionReparto
    );


    CARGAS_OPSU_set_(
      fila,
      h,
      'Agua_Diferencia',
      diferenciaAgua
    );


    CARGAS_OPSU_set_(
      fila,
      h,
      'SuperIce_Diferencia',
      diferenciaIce
    );


    CARGAS_OPSU_set_(
      fila,
      h,
      'Cuadre',
      (
        diferenciaAgua ===
          0 &&
        diferenciaIce ===
          0
      )
        ? 'CUADRA'
        : 'DIFERENCIA'
    );

  }

}


function CARGAS_OPSU_leerVentasReparto_(
  folios
) {

  const resultado = {};


  if (
    !folios ||
    !folios.size
  ) {

    return resultado;

  }


  const ss =
    CARGAS_OPSU_getSpreadsheetReparto_();


  const sh =
    ss.getSheetByName(
      SYNC_OPSU.HOJA_VENTAS
    );


  if (!sh) {

    throw new Error(
      'No existe la hoja VENTAS en Reparto.'
    );

  }


  const lastRow =
    sh.getLastRow();


  const lastCol =
    sh.getLastColumn();


  if (
    lastRow <
      2 ||
    lastCol <
      1
  ) {

    return resultado;

  }


  const headers =
    sh
      .getRange(
        1,
        1,
        1,
        lastCol
      )
      .getValues()[0];


  const h =
    CARGAS_OPSU_headers_(
      headers
    );


  const cFolio =
    CARGAS_OPSU_col_(
      h,
      'IDCarga_ref'
    );


  const cAgua =
    CARGAS_OPSU_col_(
      h,
      'Agua 20 Litros'
    );


  const cIce =
    CARGAS_OPSU_col_(
      h,
      'SuperIce'
    );


  const cComision =
    CARGAS_OPSU_col_(
      h,
      'Comision'
    );


  const filas =
    lastRow -
    1;


  const valoresFolio =
    sh
      .getRange(
        2,
        cFolio +
        1,
        filas,
        1
      )
      .getValues();


  const valoresAgua =
    sh
      .getRange(
        2,
        cAgua +
        1,
        filas,
        1
      )
      .getValues();


  const valoresIce =
    sh
      .getRange(
        2,
        cIce +
        1,
        filas,
        1
      )
      .getValues();


  const valoresComision =
    sh
      .getRange(
        2,
        cComision +
        1,
        filas,
        1
      )
      .getValues();


  for (
    let i = 0;
    i < filas;
    i++
  ) {

    const folio =
      CARGAS_OPSU_txt_(
        valoresFolio[i][0]
      );


    if (
      !folio ||
      !folios.has(
        folio
      )
    ) {

      continue;

    }


    if (
      !resultado[
        folio
      ]
    ) {

      resultado[
        folio
      ] = {

        agua:
          0,

        superice:
          0,

        comision:
          0

      };

    }


    resultado[
      folio
    ].agua +=
      CARGAS_OPSU_num_(
        valoresAgua[i][0]
      );


    resultado[
      folio
    ].superice +=
      CARGAS_OPSU_num_(
        valoresIce[i][0]
      );


    resultado[
      folio
    ].comision +=
      CARGAS_OPSU_num_(
        valoresComision[i][0]
      );

  }


  return resultado;

}


function CARGAS_OPSU_getMapaVendedores_(
  forzar
) {

  const cache =
    CacheService.getScriptCache();


  if (!forzar) {

    const guardado =
      cache.get(
        SYNC_OPSU.CACHE_PARAMETROS
      );


    if (guardado) {

      try {

        return JSON.parse(
          guardado
        );

      } catch (e) {

      }

    }

  }


  const ss =
    CARGAS_OPSU_getSpreadsheetReparto_();


  const sh =
    ss.getSheetByName(
      SYNC_OPSU.HOJA_PARAMETROS
    );


  if (!sh) {

    throw new Error(
      'No existe PARAMETROS en Reparto.'
    );

  }


  const lastRow =
    sh.getLastRow();


  const lastCol =
    sh.getLastColumn();


  const mapa = {};


  if (
    lastRow >=
    2
  ) {

    const headers =
      sh
        .getRange(
          1,
          1,
          1,
          lastCol
        )
        .getValues()[0];


    const h =
      CARGAS_OPSU_headers_(
        headers
      );


    const cCod =
      CARGAS_OPSU_col_(
        h,
        'COD'
      );


    const cEmail =
      CARGAS_OPSU_col_(
        h,
        'id_vendedor'
      );


    const filas =
      lastRow -
      1;


    const valoresCod =
      sh
        .getRange(
          2,
          cCod +
          1,
          filas,
          1
        )
        .getValues();


    const valoresEmail =
      sh
        .getRange(
          2,
          cEmail +
          1,
          filas,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < filas;
      i++
    ) {

      const cod =
        CARGAS_OPSU_cod_(
          valoresCod[i][0]
        );


      const email =
        CARGAS_OPSU_txt_(
          valoresEmail[i][0]
        );


      if (
        cod &&
        email
      ) {

        mapa[
          cod
        ] =
          email;

      }

    }

  }


  cache.put(
    SYNC_OPSU.CACHE_PARAMETROS,
    JSON.stringify(
      mapa
    ),
    SYNC_OPSU.CACHE_SEGUNDOS
  );


  return mapa;

}


function CARGAS_OPSU_getMapaRutas_() {

  const ss =
    CARGAS_OPSU_getSpreadsheetReparto_();


  const sh =
    ss.getSheetByName(
      SYNC_OPSU.HOJA_PARAMETROS
    );


  if (!sh) {

    throw new Error(
      'No existe PARAMETROS en Reparto.'
    );

  }


  const lastRow =
    sh.getLastRow();


  const lastCol =
    sh.getLastColumn();


  const mapa = {};


  if (
    lastRow >=
    2
  ) {

    const headers =
      sh
        .getRange(
          1,
          1,
          1,
          lastCol
        )
        .getValues()[0];


    const h =
      CARGAS_OPSU_headers_(
        headers
      );


    const cCod =
      CARGAS_OPSU_col_(
        h,
        'COD'
      );


    const cRuta =
      CARGAS_OPSU_col_(
        h,
        'Ruta'
      );


    const filas =
      lastRow -
      1;


    const valoresCod =
      sh
        .getRange(
          2,
          cCod +
          1,
          filas,
          1
        )
        .getValues();


    const valoresRuta =
      sh
        .getRange(
          2,
          cRuta +
          1,
          filas,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < filas;
      i++
    ) {

      const cod =
        CARGAS_OPSU_cod_(
          valoresCod[i][0]
        );


      const ruta =
        valoresRuta[i][0];


      if (
        cod &&
        CARGAS_OPSU_txt_(
          ruta
        )
      ) {

        mapa[
          cod
        ] =
          ruta;

      }

    }

  }


  return mapa;

}


function CARGAS_OPSU_validarDestino_(
  h
) {

  [

    'Folio',
    'Fecha_Carga_OPSU',
    'Fecha_Operacion',
    'Moto',
    'No_Carga',
    'COD',
    'id_vendedor',
    'Ruta_Reparto',
    'Tipo_Operacion',
    'Estatus_OPSU',
    'CargasContinuas',
    'Agua_Cargada',
    'Agua_Devuelta',
    'Agua_Venta_OPSU',
    'SuperIce_Cargado',
    'SuperIce_Devuelto',
    'SuperIce_Venta_OPSU',
    'Agua_Venta_Reparto',
    'SuperIce_Venta_Reparto',
    'Comision_Reparto',
    'Agua_Diferencia',
    'SuperIce_Diferencia',
    'Cuadre',
    'FechaHora_Sync'

  ].forEach(
    x =>
      CARGAS_OPSU_col_(
        h,
        x
      )
  );

}


function CARGAS_OPSU_headers_(
  fila
) {

  const h = {};


  fila.forEach(
    (x, i) => {

      h[
        CARGAS_OPSU_norm_(
          x
        )
      ] =
        i;

    }
  );


  return h;

}


function CARGAS_OPSU_col_(
  h,
  nombre
) {

  const key =
    CARGAS_OPSU_norm_(
      nombre
    );


  if (
    h[
      key
    ] ===
    undefined
  ) {

    throw new Error(
      'No encontré la columna: ' +
      nombre
    );

  }


  return h[
    key
  ];

}


function CARGAS_OPSU_colMulti_(
  h,
  nombres
) {

  for (
    const nombre
    of nombres
  ) {

    const key =
      CARGAS_OPSU_norm_(
        nombre
      );


    if (
      h[
        key
      ] !==
      undefined
    ) {

      return h[
        key
      ];

    }

  }


  throw new Error(
    'No encontré la columna: ' +
    nombres.join(
      ' / '
    )
  );

}


function CARGAS_OPSU_set_(
  fila,
  h,
  campo,
  valor
) {

  fila[
    CARGAS_OPSU_col_(
      h,
      campo
    )
  ] =
    valor;

}


function CARGAS_OPSU_norm_(
  v
) {

  return String(
    v ??
    ''
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
      ''
    );

}


function CARGAS_OPSU_txt_(
  v
) {

  return String(
    v ??
    ''
  ).trim();

}


function CARGAS_OPSU_cod_(
  v
) {

  const txt =
    CARGAS_OPSU_txt_(
      v
    );


  if (!txt) {
    return '';
  }


  const n =
    Number(
      txt
    );


  return Number.isFinite(
    n
  )
    ? String(
        n
      )
    : txt;

}


function CARGAS_OPSU_num_(
  v
) {

  if (
    typeof v ===
    'number'
  ) {

    return Number.isFinite(
      v
    )
      ? v
      : 0;

  }


  const n =
    Number(
      String(
        v ??
        ''
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


  return Number.isFinite(
    n
  )
    ? n
    : 0;

}


function CARGAS_OPSU_extraerId_(
  texto
) {

  const m =
    String(
      texto ??
      ''
    ).match(
      /[-\w]{25,}/
    );


  return m
    ? m[0]
    : '';

}


function CARGAS_OPSU_fechaLocal_(
  fecha
) {

  const key =
    Utilities.formatDate(
      fecha,
      SYNC_OPSU.TZ,
      'yyyy-MM-dd'
    );


  const p =
    key.split(
      '-'
    );


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


function CARGAS_OPSU_addDias_(
  fecha,
  dias
) {

  const x =
    new Date(
      fecha.getTime()
    );


  x.setDate(
    x.getDate() +
    dias
  );


  return new Date(
    x.getFullYear(),
    x.getMonth(),
    x.getDate(),
    12,
    0,
    0,
    0
  );

}


function CARGAS_OPSU_fechaKey_(
  fecha
) {

  return Utilities.formatDate(
    fecha,
    SYNC_OPSU.TZ,
    'yyyy-MM-dd'
  );

}


function CARGAS_OPSU_fechaKeySeguro_(
  valor
) {

  if (
    valor instanceof Date &&
    !isNaN(
      valor.getTime()
    )
  ) {

    return CARGAS_OPSU_fechaKey_(
      valor
    );

  }


  const txt =
    String(
      valor ??
      ''
    ).trim();


  let m =
    txt.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
    );


  if (m) {

    return (
      m[3] +
      '-' +
      String(
        m[2]
      ).padStart(
        2,
        '0'
      ) +
      '-' +
      String(
        m[1]
      ).padStart(
        2,
        '0'
      )
    );

  }


  m =
    txt.match(
      /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/
    );


  if (m) {

    return (
      m[1] +
      '-' +
      String(
        m[2]
      ).padStart(
        2,
        '0'
      ) +
      '-' +
      String(
        m[3]
      ).padStart(
        2,
        '0'
      )
    );

  }


  return '';

}
