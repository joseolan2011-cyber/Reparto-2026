// ============================================================================
// CIERRES_REPARTO.gs
// VERSION 1.2 - CIERRE INDIVIDUAL POR VENDEDOR + RUTA DESDE PARAMETROS
// ============================================================================
//
// OBJETIVO DE ESTA ETAPA:
//
//   VENTAS + CARGAS_OPSU -> CIERRES_REPARTO
//
// - Puede generar todos los cierres del día para prueba.
// - NUEVO: genera/actualiza SOLO el cierre del vendedor solicitado por AppSheet.
// - La Ruta oficial se obtiene de PARAMETROS por id_vendedor.
// - Total_Cargas = cantidad de Folios OPSU usados por ese vendedor/ruta.
// - La comisión se SUMA directamente desde VENTAS[Comision].
// - NO recalcula comisiones.
// - Cuadre y diferencias son INFORMATIVOS.
// - EN ESTA ETAPA, UNA DIFERENCIA NO BLOQUEA EL CIERRE.
// - Archivo queda vacío por ahora.
// - Estatus = PRUEBA.
//
// IMPORTANTE:
// CIERRES_REPARTO debe tener exactamente estas columnas:
//
// id_cierre
// Fecha
// Ruta
// id_vendedor
// Total_Cargas
// Agua_Cargada
// Agua_Venta_OPSU
// Agua_Venta_Reparto
// Agua_Devuelta
// Agua_Diferencia
// SuperIce_Cargado
// SuperIce_Venta_OPSU
// SuperIce_Venta_Reparto
// SuperIce_Devuelto
// SuperIce_Diferencia
// Venta_Total_Dinero
// Comision_Total
// Cuadre
// FechaHora_Cierre
// Estatus
// Archivo
//
// ============================================================================


const CR_CFG = {

  TZ:
    'America/Merida',

  HOJA_VENTAS:
    'VENTAS',

  HOJA_CARGAS:
    'CARGAS_OPSU',

  HOJA_PARAMETROS:
    'PARAMETROS',

  HOJA_CIERRES:
    'CIERRES_REPARTO',

  ESTATUS_PRUEBA:
    'PRUEBA'

};


// ============================================================================
// PRUEBA SIN ESCRIBIR
// ============================================================================

function CIERRES_REPARTO_probarHoy() {

  const fechaKey =
    CR_fechaKey_(
      new Date()
    );


  const cierres =
    CR_calcularCierres_(
      fechaKey
    );


  let texto =
    'PRUEBA CIERRES REPARTO\n\n' +
    'Fecha: ' +
    fechaKey +
    '\n' +
    'Cierres encontrados: ' +
    cierres.length +
    '\n\n';


  if (
    !cierres.length
  ) {

    texto +=
      'No encontré operación de Reparto para esa fecha.';

  }
  else {

    cierres.forEach(
      c => {

        texto +=
          'Ruta ' +
          (
            c.Ruta ||
            'SIN RUTA'
          ) +
          ' | ' +
          c.id_vendedor +
          '\n' +
          'Cargas: ' +
          c.Total_Cargas +
          ' | Agua cargada: ' +
          c.Agua_Cargada +
          ' | Agua venta OPSU: ' +
          c.Agua_Venta_OPSU +
          ' | Agua venta Reparto: ' +
          c.Agua_Venta_Reparto +
          '\n' +
          'Ice cargado: ' +
          c.SuperIce_Cargado +
          ' | Ice venta OPSU: ' +
          c.SuperIce_Venta_OPSU +
          ' | Ice venta Reparto: ' +
          c.SuperIce_Venta_Reparto +
          '\n' +
          'Venta $: ' +
          CR_moneda_(
            c.Venta_Total_Dinero
          ) +
          ' | Comisión: ' +
          CR_moneda_(
            c.Comision_Total
          ) +
          '\n' +
          'Cuadre informativo: ' +
          c.Cuadre +
          '\n\n';

      }
    );

  }


  texto +=
    'NO SE ESCRIBIÓ NADA.';


  SpreadsheetApp
    .getUi()
    .alert(
      texto
    );


  console.log(
    texto
  );


  return cierres;

}


// ============================================================================
// GENERAR / ACTUALIZAR CIERRE DE PRUEBA DE HOY
// ============================================================================

function CIERRES_REPARTO_generarPruebaHoy() {

  const fechaKey =
    CR_fechaKey_(
      new Date()
    );


  return CR_generarCierres_(
    fechaKey,
    true
  );

}


// ============================================================================
// CIERRE INDIVIDUAL - FUNCIÓN PARA APPSHEET
// ============================================================================
//
// AppSheet puede llamar esta función con:
//
//   fecha       -> [Fecha] o TODAY()
//   ruta        -> [Ruta] (se recibe, pero PARAMETROS es la fuente oficial)
//   id_vendedor -> [id_vendedor] o USEREMAIL()
//
// IMPORTANTE:
// - Solo crea/actualiza UNA fila en CIERRES_REPARTO.
// - NO genera cierres de otros vendedores.
// - NO bloquea por Cuadre = DIFERENCIA.
// - Estatus continúa en PRUEBA durante esta etapa.
// - Si AppSheet manda una Ruta distinta a PARAMETROS, usa PARAMETROS y deja log.
// ============================================================================

function CIERRES_REPARTO_generarCierreVendedor(
  fecha,
  ruta,
  id_vendedor
) {

  const inicio =
    Date.now();


  const vendedor =
    CR_txt_(
      id_vendedor
    );


  if (
    !vendedor
  ) {

    throw new Error(
      'Falta id_vendedor para generar el cierre.'
    );

  }


  const fechaKey =
    CR_resolverFechaEntrada_(
      fecha
    );


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'No pude abrir el archivo de Reparto.'
    );

  }


  const shParametros =
    ss.getSheetByName(
      CR_CFG.HOJA_PARAMETROS
    );


  if (!shParametros) {

    throw new Error(
      'No existe la hoja PARAMETROS.'
    );

  }


  const rutasParametros =
    CR_leerRutasParametros_(
      shParametros
    );


  const rutaOficial =
    CR_txt_(
      rutasParametros[
        vendedor.toLowerCase()
      ]
    );


  if (
    !rutaOficial
  ) {

    throw new Error(
      'El vendedor ' +
      vendedor +
      ' no tiene Ruta configurada en PARAMETROS.'
    );

  }


  const rutaRecibida =
    CR_txt_(
      ruta
    );


  if (
    rutaRecibida &&
    rutaRecibida !==
      rutaOficial
  ) {

    console.warn(
      'CIERRES_REPARTO: AppSheet envió Ruta ' +
      rutaRecibida +
      ' para ' +
      vendedor +
      ', pero PARAMETROS indica Ruta ' +
      rutaOficial +
      '. Se usará PARAMETROS.'
    );

  }


  // Refrescamos CARGAS_OPSU antes de calcular el cierre.
  CR_refrescarCargasAntesDeCierre_();


  const cierres =
    CR_calcularCierres_(
      fechaKey
    );


  const cierre =
    cierres.find(
      c =>
        CR_txt_(
          c.id_vendedor
        ).toLowerCase() ===
          vendedor.toLowerCase() &&
        CR_txt_(
          c.Ruta
        ) ===
          rutaOficial
    );


  if (
    !cierre
  ) {

    throw new Error(
      'No encontré operación para ' +
      vendedor +
      ' | Ruta ' +
      rutaOficial +
      ' | Fecha ' +
      fechaKey +
      '.'
    );

  }


  const resultado =
    CR_guardarCierreIndividual_(
      fechaKey,
      cierre
    );


  resultado.segundos =
    (
      (
        Date.now() -
        inicio
      ) /
      1000
    ).toFixed(
      1
    );


  console.log(
    'CIERRES_REPARTO INDIVIDUAL: ' +
    JSON.stringify(
      resultado
    )
  );


  return resultado;

}


// ============================================================================
// PRUEBA MANUAL FÁCIL DESDE PARAMETROS
//
// 1) Abre la hoja PARAMETROS.
// 2) Selecciona cualquier celda de la fila del vendedor.
// 3) Ejecuta esta función.
//
// NO ESCRIBE NADA.
// ============================================================================

function CIERRES_REPARTO_probarVendedorSeleccionadoHoy() {

  const datos =
    CR_vendedorSeleccionadoEnParametros_();


  const fechaKey =
    CR_fechaKey_(
      new Date()
    );


  CR_refrescarCargasAntesDeCierre_();


  const cierres =
    CR_calcularCierres_(
      fechaKey
    );


  const cierre =
    cierres.find(
      c =>
        CR_txt_(
          c.id_vendedor
        ).toLowerCase() ===
          datos.id_vendedor.toLowerCase() &&
        CR_txt_(
          c.Ruta
        ) ===
          CR_txt_(
            datos.Ruta
          )
    );


  if (
    !cierre
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        'No encontré operación para:\n\n' +
        datos.id_vendedor +
        '\nRuta ' +
        datos.Ruta +
        '\nFecha ' +
        fechaKey +
        '\n\nNO SE ESCRIBIÓ NADA.'
      );


    return null;

  }


  const texto =
    'PRUEBA CIERRE INDIVIDUAL\n\n' +
    'Fecha: ' +
    fechaKey +
    '\n' +
    'Ruta: ' +
    cierre.Ruta +
    '\n' +
    'Vendedor: ' +
    cierre.id_vendedor +
    '\n\n' +
    'Cargas: ' +
    cierre.Total_Cargas +
    '\n' +
    'Agua cargada: ' +
    cierre.Agua_Cargada +
    '\n' +
    'Agua venta OPSU: ' +
    cierre.Agua_Venta_OPSU +
    '\n' +
    'Agua venta Reparto: ' +
    cierre.Agua_Venta_Reparto +
    '\n' +
    'Agua devuelta: ' +
    cierre.Agua_Devuelta +
    '\n\n' +
    'SuperIce cargado: ' +
    cierre.SuperIce_Cargado +
    '\n' +
    'SuperIce venta OPSU: ' +
    cierre.SuperIce_Venta_OPSU +
    '\n' +
    'SuperIce venta Reparto: ' +
    cierre.SuperIce_Venta_Reparto +
    '\n' +
    'SuperIce devuelto: ' +
    cierre.SuperIce_Devuelto +
    '\n\n' +
    'Venta: ' +
    CR_moneda_(
      cierre.Venta_Total_Dinero
    ) +
    '\n' +
    'Comisión: ' +
    CR_moneda_(
      cierre.Comision_Total
    ) +
    '\n' +
    'Cuadre informativo: ' +
    cierre.Cuadre +
    '\n\n' +
    'NO SE ESCRIBIÓ NADA.';


  SpreadsheetApp
    .getUi()
    .alert(
      texto
    );


  return cierre;

}


// ============================================================================
// PRUEBA MANUAL QUE SÍ ESCRIBE SOLO EL VENDEDOR SELECCIONADO
//
// 1) Abre PARAMETROS.
// 2) Selecciona la fila del vendedor.
// 3) Ejecuta.
//
// Útil antes de conectar AppSheet.
// ============================================================================

function CIERRES_REPARTO_generarVendedorSeleccionadoHoy() {

  const datos =
    CR_vendedorSeleccionadoEnParametros_();


  return CIERRES_REPARTO_generarCierreVendedor(
    new Date(),
    datos.Ruta,
    datos.id_vendedor
  );

}


// ============================================================================
// GUARDAR UNA SOLA FILA DE CIERRE
// ============================================================================

function CR_guardarCierreIndividual_(
  fechaKey,
  cierre
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sh =
    ss.getSheetByName(
      CR_CFG.HOJA_CIERRES
    );


  if (!sh) {

    throw new Error(
      'No existe la hoja CIERRES_REPARTO.'
    );

  }


  const lastCol =
    sh.getLastColumn();


  if (
    lastCol <
    1
  ) {

    throw new Error(
      'CIERRES_REPARTO no tiene encabezados.'
    );

  }


  const lastRow =
    Math.max(
      1,
      sh.getLastRow()
    );


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
    CR_headers_(
      tabla[0]
    );


  CR_validarCierres_(
    h
  );


  const cId =
    CR_col_(
      h,
      'id_cierre'
    );


  const cFecha =
    CR_col_(
      h,
      'Fecha'
    );


  const cRuta =
    CR_col_(
      h,
      'Ruta'
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cArchivo =
    CR_col_(
      h,
      'Archivo'
    );


  const claveBuscada =
    CR_clave_(
      fechaKey,
      cierre.Ruta,
      cierre.id_vendedor
    );


  let filaHoja =
    0;


  let fila =
    null;


  for (
    let i = 1;
    i < tabla.length;
    i++
  ) {

    const f =
      CR_fechaKeySeguro_(
        tabla[i][
          cFecha
        ]
      );


    const r =
      CR_txt_(
        tabla[i][
          cRuta
        ]
      );


    const v =
      CR_txt_(
        tabla[i][
          cVendedor
        ]
      );


    if (
      CR_clave_(
        f,
        r,
        v
      ) ===
      claveBuscada
    ) {

      filaHoja =
        i +
        1;


      fila =
        tabla[i]
          .slice();


      break;

    }

  }


  let creado =
    false;


  if (
    !fila
  ) {

    fila =
      new Array(
        lastCol
      ).fill(
        ''
      );


    CR_set_(
      fila,
      h,
      'id_cierre',
      CR_nuevoIdCierre_(
        fechaKey,
        cierre.Ruta
      )
    );


    filaHoja =
      sh.getLastRow() +
      1;


    creado =
      true;

  }


  const idCierreExistente =
    CR_txt_(
      fila[
        cId
      ]
    );


  const archivoExistente =
    fila[
      cArchivo
    ];


  if (
    !idCierreExistente
  ) {

    CR_set_(
      fila,
      h,
      'id_cierre',
      CR_nuevoIdCierre_(
        fechaKey,
        cierre.Ruta
      )
    );

  }


  CR_set_(
    fila,
    h,
    'Fecha',
    CR_fechaDesdeKey_(
      fechaKey
    )
  );


  CR_set_(
    fila,
    h,
    'Ruta',
    cierre.Ruta
  );


  CR_set_(
    fila,
    h,
    'id_vendedor',
    cierre.id_vendedor
  );


  CR_set_(
    fila,
    h,
    'Total_Cargas',
    cierre.Total_Cargas
  );


  CR_set_(
    fila,
    h,
    'Agua_Cargada',
    cierre.Agua_Cargada
  );


  CR_set_(
    fila,
    h,
    'Agua_Venta_OPSU',
    cierre.Agua_Venta_OPSU
  );


  CR_set_(
    fila,
    h,
    'Agua_Venta_Reparto',
    cierre.Agua_Venta_Reparto
  );


  CR_set_(
    fila,
    h,
    'Agua_Devuelta',
    cierre.Agua_Devuelta
  );


  CR_set_(
    fila,
    h,
    'Agua_Diferencia',
    cierre.Agua_Diferencia
  );


  CR_set_(
    fila,
    h,
    'SuperIce_Cargado',
    cierre.SuperIce_Cargado
  );


  CR_set_(
    fila,
    h,
    'SuperIce_Venta_OPSU',
    cierre.SuperIce_Venta_OPSU
  );


  CR_set_(
    fila,
    h,
    'SuperIce_Venta_Reparto',
    cierre.SuperIce_Venta_Reparto
  );


  CR_set_(
    fila,
    h,
    'SuperIce_Devuelto',
    cierre.SuperIce_Devuelto
  );


  CR_set_(
    fila,
    h,
    'SuperIce_Diferencia',
    cierre.SuperIce_Diferencia
  );


  CR_set_(
    fila,
    h,
    'Venta_Total_Dinero',
    cierre.Venta_Total_Dinero
  );


  CR_set_(
    fila,
    h,
    'Comision_Total',
    cierre.Comision_Total
  );


  // POR AHORA SOLO INFORMATIVO.
  // NO BLOQUEA EL CIERRE.
  CR_set_(
    fila,
    h,
    'Cuadre',
    cierre.Cuadre
  );


  CR_set_(
    fila,
    h,
    'FechaHora_Cierre',
    new Date()
  );


  CR_set_(
    fila,
    h,
    'Estatus',
    CR_CFG.ESTATUS_PRUEBA
  );


  CR_set_(
    fila,
    h,
    'Archivo',
    archivoExistente ||
    ''
  );


  sh
    .getRange(
      filaHoja,
      1,
      1,
      lastCol
    )
    .setValues(
      [
        fila
      ]
    );


  SpreadsheetApp.flush();


  return {

    ok:
      true,

    id_cierre:
      CR_txt_(
        fila[
          cId
        ]
      ),

    fecha:
      fechaKey,

    Ruta:
      CR_txt_(
        cierre.Ruta
      ),

    id_vendedor:
      CR_txt_(
        cierre.id_vendedor
      ),

    Total_Cargas:
      cierre.Total_Cargas,

    Venta_Total_Dinero:
      cierre.Venta_Total_Dinero,

    Comision_Total:
      cierre.Comision_Total,

    Cuadre:
      cierre.Cuadre,

    accion:
      creado
        ? 'CREADO'
        : 'ACTUALIZADO'

  };

}


// ============================================================================
// REFRESCAR CARGAS ANTES DE CERRAR
// ============================================================================

function CR_refrescarCargasAntesDeCierre_() {

  try {

    if (
      typeof CARGAS_OPSU_sincronizarCargasCore_ ===
      'function'
    ) {

      const sync =
        CARGAS_OPSU_sincronizarCargasCore_(
          false
        );


      console.log(
        'CIERRES_REPARTO - sync previo CARGAS_OPSU: ' +
        JSON.stringify(
          sync
        )
      );


      return sync;

    }

  } catch (error) {

    console.warn(
      'CIERRES_REPARTO - no se pudo refrescar CARGAS_OPSU antes del cierre: ' +
      error.message
    );

  }


  return null;

}


// ============================================================================
// RESOLVER FECHA RECIBIDA DESDE APPSHEET
// ============================================================================

function CR_resolverFechaEntrada_(
  fecha
) {

  if (
    fecha instanceof Date &&
    !isNaN(
      fecha.getTime()
    )
  ) {

    return CR_fechaKey_(
      fecha
    );

  }


  const key =
    CR_fechaKeySeguro_(
      fecha
    );


  if (
    key
  ) {

    return key;

  }


  if (
    fecha ===
      undefined ||
    fecha ===
      null ||
    CR_txt_(
      fecha
    ) ===
      ''
  ) {

    return CR_fechaKey_(
      new Date()
    );

  }


  throw new Error(
    'No pude interpretar la fecha recibida: ' +
    fecha
  );

}


// ============================================================================
// OBTENER VENDEDOR DE LA FILA SELECCIONADA EN PARAMETROS
// ============================================================================

function CR_vendedorSeleccionadoEnParametros_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sh =
    ss.getActiveSheet();


  if (
    !sh ||
    sh.getName() !==
      CR_CFG.HOJA_PARAMETROS
  ) {

    throw new Error(
      'Abre la hoja PARAMETROS y selecciona una fila de vendedor.'
    );

  }


  const row =
    sh
      .getActiveRange()
      .getRow();


  if (
    row <
      2
  ) {

    throw new Error(
      'Selecciona una fila de datos en PARAMETROS, no el encabezado.'
    );

  }


  const lastCol =
    sh.getLastColumn();


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
    CR_headers_(
      headers
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cRuta =
    CR_col_(
      h,
      'Ruta'
    );


  const fila =
    sh
      .getRange(
        row,
        1,
        1,
        lastCol
      )
      .getValues()[0];


  const vendedor =
    CR_txt_(
      fila[
        cVendedor
      ]
    );


  const ruta =
    CR_txt_(
      fila[
        cRuta
      ]
    );


  if (
    !vendedor
  ) {

    throw new Error(
      'La fila seleccionada no tiene id_vendedor.'
    );

  }


  if (
    !ruta
  ) {

    throw new Error(
      'El vendedor seleccionado no tiene Ruta en PARAMETROS.'
    );

  }


  return {

    id_vendedor:
      vendedor,

    Ruta:
      ruta

  };

}


// ============================================================================
// NUCLEO DE ESCRITURA
// ============================================================================

function CR_generarCierres_(
  fechaKey,
  mostrarUI
) {

  const inicio =
    Date.now();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'No pude abrir el archivo de Reparto.'
    );

  }


  const shCierres =
    ss.getSheetByName(
      CR_CFG.HOJA_CIERRES
    );


  if (!shCierres) {

    throw new Error(
      'No existe la hoja CIERRES_REPARTO.'
    );

  }


  // Refresca CARGAS_OPSU antes del cálculo.
  CR_refrescarCargasAntesDeCierre_();


  const cierres =
    CR_calcularCierres_(
      fechaKey
    );


  if (
    !cierres.length
  ) {

    const mensaje =
      'No encontré operación de Reparto para ' +
      fechaKey +
      '.';


    if (
      mostrarUI
    ) {

      SpreadsheetApp
        .getUi()
        .alert(
          mensaje
        );

    }


    return {
      ok:
        true,

      fecha:
        fechaKey,

      creados:
        0,

      actualizados:
        0,

      mensaje:
        mensaje
    };

  }


  const lastCol =
    shCierres.getLastColumn();


  if (
    lastCol <
    1
  ) {

    throw new Error(
      'CIERRES_REPARTO no tiene encabezados.'
    );

  }


  const lastRow =
    Math.max(
      1,
      shCierres.getLastRow()
    );


  const tabla =
    shCierres
      .getRange(
        1,
        1,
        lastRow,
        lastCol
      )
      .getValues();


  const headers =
    tabla[0];


  const h =
    CR_headers_(
      headers
    );


  CR_validarCierres_(
    h
  );


  const cId =
    CR_col_(
      h,
      'id_cierre'
    );


  const cFecha =
    CR_col_(
      h,
      'Fecha'
    );


  const cRuta =
    CR_col_(
      h,
      'Ruta'
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cArchivo =
    CR_col_(
      h,
      'Archivo'
    );


  // --------------------------------------------------------------------------
  // Índice de cierres existentes.
  //
  // Re-ejecutar la prueba del mismo día NO crea duplicados:
  // actualiza el registro de Fecha + Ruta + id_vendedor.
  // --------------------------------------------------------------------------

  const existentes = {};


  for (
    let i = 1;
    i < tabla.length;
    i++
  ) {

    const f =
      CR_fechaKeySeguro_(
        tabla[i][
          cFecha
        ]
      );


    const ruta =
      CR_txt_(
        tabla[i][
          cRuta
        ]
      );


    const vendedor =
      CR_txt_(
        tabla[i][
          cVendedor
        ]
      );


    if (
      !f ||
      !vendedor
    ) {

      continue;

    }


    existentes[
      CR_clave_(
        f,
        ruta,
        vendedor
      )
    ] =
      i;

  }


  let creados = 0;
  let actualizados = 0;


  cierres.forEach(
    cierre => {

      const clave =
        CR_clave_(
          fechaKey,
          cierre.Ruta,
          cierre.id_vendedor
        );


      let pos =
        existentes[
          clave
        ];


      let fila;


      if (
        pos ===
        undefined
      ) {

        fila =
          new Array(
            lastCol
          ).fill(
            ''
          );


        CR_set_(
          fila,
          h,
          'id_cierre',
          CR_nuevoIdCierre_(
            fechaKey,
            cierre.Ruta
          )
        );


        tabla.push(
          fila
        );


        pos =
          tabla.length -
          1;


        existentes[
          clave
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


      // Preservar Archivo si posteriormente ya existe.
      const archivoExistente =
        fila[
          cArchivo
        ];


      CR_set_(
        fila,
        h,
        'Fecha',
        CR_fechaDesdeKey_(
          fechaKey
        )
      );


      CR_set_(
        fila,
        h,
        'Ruta',
        cierre.Ruta
      );


      CR_set_(
        fila,
        h,
        'id_vendedor',
        cierre.id_vendedor
      );


      CR_set_(
        fila,
        h,
        'Total_Cargas',
        cierre.Total_Cargas
      );


      CR_set_(
        fila,
        h,
        'Agua_Cargada',
        cierre.Agua_Cargada
      );


      CR_set_(
        fila,
        h,
        'Agua_Venta_OPSU',
        cierre.Agua_Venta_OPSU
      );


      CR_set_(
        fila,
        h,
        'Agua_Venta_Reparto',
        cierre.Agua_Venta_Reparto
      );


      CR_set_(
        fila,
        h,
        'Agua_Devuelta',
        cierre.Agua_Devuelta
      );


      CR_set_(
        fila,
        h,
        'Agua_Diferencia',
        cierre.Agua_Diferencia
      );


      CR_set_(
        fila,
        h,
        'SuperIce_Cargado',
        cierre.SuperIce_Cargado
      );


      CR_set_(
        fila,
        h,
        'SuperIce_Venta_OPSU',
        cierre.SuperIce_Venta_OPSU
      );


      CR_set_(
        fila,
        h,
        'SuperIce_Venta_Reparto',
        cierre.SuperIce_Venta_Reparto
      );


      CR_set_(
        fila,
        h,
        'SuperIce_Devuelto',
        cierre.SuperIce_Devuelto
      );


      CR_set_(
        fila,
        h,
        'SuperIce_Diferencia',
        cierre.SuperIce_Diferencia
      );


      CR_set_(
        fila,
        h,
        'Venta_Total_Dinero',
        cierre.Venta_Total_Dinero
      );


      CR_set_(
        fila,
        h,
        'Comision_Total',
        cierre.Comision_Total
      );


      // SOLO INFORMATIVO EN PRUEBAS.
      CR_set_(
        fila,
        h,
        'Cuadre',
        cierre.Cuadre
      );


      CR_set_(
        fila,
        h,
        'FechaHora_Cierre',
        new Date()
      );


      CR_set_(
        fila,
        h,
        'Estatus',
        CR_CFG.ESTATUS_PRUEBA
      );


      CR_set_(
        fila,
        h,
        'Archivo',
        archivoExistente ||
        ''
      );

    }
  );


  shCierres
    .getRange(
      1,
      1,
      tabla.length,
      lastCol
    )
    .setValues(
      tabla
    );


  SpreadsheetApp.flush();


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

    fecha:
      fechaKey,

    cierres:
      cierres.length,

    creados:
      creados,

    actualizados:
      actualizados,

    segundos:
      segundos

  };


  console.log(
    'CIERRES_REPARTO: ' +
    JSON.stringify(
      resultado
    )
  );


  if (
    mostrarUI
  ) {

    let texto =
      'CIERRES REPARTO - PRUEBA\n\n' +
      'Fecha: ' +
      fechaKey +
      '\n' +
      'Cierres procesados: ' +
      cierres.length +
      '\n' +
      'Nuevos: ' +
      creados +
      '\n' +
      'Actualizados: ' +
      actualizados +
      '\n' +
      'Tiempo: ' +
      segundos +
      ' s\n\n';


    cierres.forEach(
      c => {

        texto +=
          'Ruta ' +
          (
            c.Ruta ||
            'SIN RUTA'
          ) +
          ' | Cargas ' +
          c.Total_Cargas +
          ' | Venta ' +
          CR_moneda_(
            c.Venta_Total_Dinero
          ) +
          ' | Comisión ' +
          CR_moneda_(
            c.Comision_Total
          ) +
          ' | ' +
          c.Cuadre +
          '\n';

      }
    );


    texto +=
      '\nIMPORTANTE: Cuadre es informativo. No bloqueó ningún cierre.';


    SpreadsheetApp
      .getUi()
      .alert(
        texto
      );

  }


  return resultado;

}


// ============================================================================
// CALCULAR CIERRES
// ============================================================================

function CR_calcularCierres_(
  fechaKey
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const shVentas =
    ss.getSheetByName(
      CR_CFG.HOJA_VENTAS
    );


  const shCargas =
    ss.getSheetByName(
      CR_CFG.HOJA_CARGAS
    );


  const shParametros =
    ss.getSheetByName(
      CR_CFG.HOJA_PARAMETROS
    );


  if (!shVentas) {

    throw new Error(
      'No existe la hoja VENTAS.'
    );

  }


  if (!shCargas) {

    throw new Error(
      'No existe la hoja CARGAS_OPSU.'
    );

  }


  if (!shParametros) {

    throw new Error(
      'No existe la hoja PARAMETROS.'
    );

  }


  // Ruta oficial del vendedor.
  // PARAMETROS manda; ya no dependemos de SWITCH por correo
  // ni de que el vendedor haya hecho una venta para conocer su ruta.
  const rutasParametros =
    CR_leerRutasParametros_(
      shParametros
    );


  const ventas =
    CR_leerVentas_(
      shVentas,
      fechaKey,
      rutasParametros
    );


  const cargas =
    CR_leerCargas_(
      shCargas,
      fechaKey,
      ventas.rutasPorVendedor,
      rutasParametros
    );


  const claves =
    new Set(
      [
        ...Object.keys(
          ventas.porClave
        ),
        ...Object.keys(
          cargas.porClave
        )
      ]
    );


  const cierres = [];


  claves.forEach(
    clave => {

      const v =
        ventas.porClave[
          clave
        ] ||
        CR_ventaVacia_();


      const c =
        cargas.porClave[
          clave
        ] ||
        CR_cargaVacia_();


      const vendedor =
        v.id_vendedor ||
        c.id_vendedor;


      const ruta =
        v.Ruta ||
        c.Ruta ||
        '';


      if (
        !vendedor
      ) {

        return;

      }


      // ----------------------------------------------------------------------
      // La venta de Reparto y la comisión vienen DIRECTAMENTE DE VENTAS.
      // No recalculamos comisión aquí.
      // ----------------------------------------------------------------------

      const aguaReparto =
        CR_num_(
          v.Agua_Venta_Reparto
        );


      const iceReparto =
        CR_num_(
          v.SuperIce_Venta_Reparto
        );


      const aguaOpsu =
        CR_num_(
          c.Agua_Venta_OPSU
        );


      const iceOpsu =
        CR_num_(
          c.SuperIce_Venta_OPSU
        );


      const aguaDiferencia =
        aguaOpsu -
        aguaReparto;


      const iceDiferencia =
        iceOpsu -
        iceReparto;


      const cuadre =
        (
          aguaDiferencia ===
            0 &&
          iceDiferencia ===
            0
        )
          ? 'CUADRA'
          : 'DIFERENCIA';


      cierres.push({

        Fecha:
          fechaKey,

        Ruta:
          ruta,

        id_vendedor:
          vendedor,

        Total_Cargas:
          c.folios.size,

        Agua_Cargada:
          c.Agua_Cargada,

        Agua_Venta_OPSU:
          aguaOpsu,

        Agua_Venta_Reparto:
          aguaReparto,

        Agua_Devuelta:
          c.Agua_Devuelta,

        Agua_Diferencia:
          aguaDiferencia,

        SuperIce_Cargado:
          c.SuperIce_Cargado,

        SuperIce_Venta_OPSU:
          iceOpsu,

        SuperIce_Venta_Reparto:
          iceReparto,

        SuperIce_Devuelto:
          c.SuperIce_Devuelto,

        SuperIce_Diferencia:
          iceDiferencia,

        Venta_Total_Dinero:
          CR_num_(
            v.Venta_Total_Dinero
          ),

        Comision_Total:
          CR_num_(
            v.Comision_Total
          ),

        Cuadre:
          cuadre,

        Folios:
          Array.from(
            c.folios
          )

      });

    }
  );


  cierres.sort(
    (a, b) => {

      const ra =
        CR_txt_(
          a.Ruta
        );


      const rb =
        CR_txt_(
          b.Ruta
        );


      return ra.localeCompare(
        rb,
        'es',
        {
          numeric:
            true
        }
      );

    }
  );


  return cierres;

}


// ============================================================================
// LEER VENTAS DEL DÍA
// ============================================================================

function CR_leerVentas_(
  sh,
  fechaKey,
  rutasParametros
) {

  const lastRow =
    sh.getLastRow();


  const lastCol =
    sh.getLastColumn();


  const resultado = {

    porClave:
      {},

    rutasPorVendedor:
      {}

  };


  if (
    lastRow <
      2 ||
    lastCol <
      1
  ) {

    return resultado;

  }


  const data =
    sh
      .getRange(
        1,
        1,
        lastRow,
        lastCol
      )
      .getValues();


  const h =
    CR_headers_(
      data[0]
    );


  const cFecha =
    CR_colMulti_(
      h,
      [
        'Fecha',
        'fecha_hora'
      ]
    );


  const cRuta =
    CR_colMulti_(
      h,
      [
        'ruta',
        'Ruta'
      ]
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cAgua =
    CR_colMulti_(
      h,
      [
        'Agua 20 Litros',
        'Agua 20 Litros '
      ]
    );


  const cIce =
    CR_col_(
      h,
      'SuperIce'
    );


  const cImporte =
    CR_col_(
      h,
      'Importe'
    );


  const cComision =
    CR_col_(
      h,
      'Comision'
    );


  const cFolio =
    CR_col_(
      h,
      'IDCarga_ref'
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const fila =
      data[i];


    const f =
      CR_fechaKeySeguro_(
        fila[
          cFecha
        ]
      );


    if (
      f !==
      fechaKey
    ) {

      continue;

    }


    const vendedor =
      CR_txt_(
        fila[
          cVendedor
        ]
      );


    if (
      !vendedor
    ) {

      continue;

    }


    // PARAMETROS es la fuente principal de Ruta.
    // Si el vendedor aún no está configurado ahí, conservamos
    // temporalmente la ruta física de VENTAS como respaldo.
    const ruta =
      CR_txt_(
        rutasParametros[
          vendedor.toLowerCase()
        ] ||
        fila[
          cRuta
        ]
      );


    const clave =
      CR_clave_(
        fechaKey,
        ruta,
        vendedor
      );


    if (
      !resultado.porClave[
        clave
      ]
    ) {

      resultado.porClave[
        clave
      ] = {

        Ruta:
          ruta,

        id_vendedor:
          vendedor,

        Agua_Venta_Reparto:
          0,

        SuperIce_Venta_Reparto:
          0,

        Venta_Total_Dinero:
          0,

        Comision_Total:
          0,

        folios:
          new Set()

      };

    }


    const x =
      resultado.porClave[
        clave
      ];


    x.Agua_Venta_Reparto +=
      CR_num_(
        fila[
          cAgua
        ]
      );


    x.SuperIce_Venta_Reparto +=
      CR_num_(
        fila[
          cIce
        ]
      );


    x.Venta_Total_Dinero +=
      CR_num_(
        fila[
          cImporte
        ]
      );


    x.Comision_Total +=
      CR_num_(
        fila[
          cComision
        ]
      );


    const folio =
      CR_txt_(
        fila[
          cFolio
        ]
      );


    if (
      folio
    ) {

      x.folios.add(
        folio
      );

    }


    if (
      !resultado.rutasPorVendedor[
        vendedor
      ]
    ) {

      resultado.rutasPorVendedor[
        vendedor
      ] =
        new Set();

    }


    if (
      ruta
    ) {

      resultado
        .rutasPorVendedor[
          vendedor
        ]
        .add(
          ruta
        );

    }

  }


  return resultado;

}


// ============================================================================
// LEER CARGAS OPSU DEL DÍA
// ============================================================================

function CR_leerCargas_(
  sh,
  fechaKey,
  rutasPorVendedor,
  rutasParametros
) {

  const resultado = {

    porClave:
      {}

  };


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


  const data =
    sh
      .getRange(
        1,
        1,
        lastRow,
        lastCol
      )
      .getValues();


  const h =
    CR_headers_(
      data[0]
    );


  const cFolio =
    CR_col_(
      h,
      'Folio'
    );


  const cFecha =
    CR_col_(
      h,
      'Fecha_Operacion'
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cRuta =
    CR_col_(
      h,
      'Ruta_Reparto'
    );


  const cTipo =
    CR_col_(
      h,
      'Tipo_Operacion'
    );


  const cAguaCarga =
    CR_col_(
      h,
      'Agua_Cargada'
    );


  const cAguaDev =
    CR_col_(
      h,
      'Agua_Devuelta'
    );


  const cAguaOpsu =
    CR_col_(
      h,
      'Agua_Venta_OPSU'
    );


  const cIceCarga =
    CR_col_(
      h,
      'SuperIce_Cargado'
    );


  const cIceDev =
    CR_col_(
      h,
      'SuperIce_Devuelto'
    );


  const cIceOpsu =
    CR_col_(
      h,
      'SuperIce_Venta_OPSU'
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const fila =
      data[i];


    const tipo =
      CR_norm_(
        fila[
          cTipo
        ]
      );


    if (
      tipo !==
      'reparto'
    ) {

      continue;

    }


    const f =
      CR_fechaKeySeguro_(
        fila[
          cFecha
        ]
      );


    if (
      f !==
      fechaKey
    ) {

      continue;

    }


    const vendedor =
      CR_txt_(
        fila[
          cVendedor
        ]
      );


    if (
      !vendedor
    ) {

      continue;

    }


    // PARAMETROS es la fuente principal.
    // Orden de respaldo:
    //   1) PARAMETROS[Ruta]
    //   2) CARGAS_OPSU[Ruta_Reparto]
    //   3) única ruta encontrada en VENTAS
    let ruta =
      CR_txt_(
        rutasParametros[
          vendedor.toLowerCase()
        ] ||
        fila[
          cRuta
        ]
      );


    if (
      !ruta &&
      rutasPorVendedor[
        vendedor
      ] &&
      rutasPorVendedor[
        vendedor
      ].size ===
        1
    ) {

      ruta =
        Array.from(
          rutasPorVendedor[
            vendedor
          ]
        )[0];

    }


    const clave =
      CR_clave_(
        fechaKey,
        ruta,
        vendedor
      );


    if (
      !resultado.porClave[
        clave
      ]
    ) {

      resultado.porClave[
        clave
      ] = {

        Ruta:
          ruta,

        id_vendedor:
          vendedor,

        folios:
          new Set(),

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

    }


    const x =
      resultado.porClave[
        clave
      ];


    const folio =
      CR_txt_(
        fila[
          cFolio
        ]
      );


    if (
      folio
    ) {

      x.folios.add(
        folio
      );

    }


    x.Agua_Cargada +=
      CR_num_(
        fila[
          cAguaCarga
        ]
      );


    x.Agua_Devuelta +=
      CR_num_(
        fila[
          cAguaDev
        ]
      );


    x.Agua_Venta_OPSU +=
      CR_num_(
        fila[
          cAguaOpsu
        ]
      );


    x.SuperIce_Cargado +=
      CR_num_(
        fila[
          cIceCarga
        ]
      );


    x.SuperIce_Devuelto +=
      CR_num_(
        fila[
          cIceDev
        ]
      );


    x.SuperIce_Venta_OPSU +=
      CR_num_(
        fila[
          cIceOpsu
        ]
      );

  }


  return resultado;

}


// ============================================================================
// RUTAS OFICIALES DESDE PARAMETROS
//
// Lee:
//   PARAMETROS[id_vendedor]
//   PARAMETROS[Ruta]
//
// Devuelve:
// {
//   "ventas1.superagua@gmail.com": "1",
//   "ventas2.superagua@gmail.com": "2",
//   ...
// }
//
// La llave se normaliza a minúsculas para evitar problemas de mayúsculas.
// ============================================================================

function CR_leerRutasParametros_(
  sh
) {

  const resultado = {};


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


  const data =
    sh
      .getRange(
        1,
        1,
        lastRow,
        lastCol
      )
      .getValues();


  const h =
    CR_headers_(
      data[0]
    );


  const cVendedor =
    CR_col_(
      h,
      'id_vendedor'
    );


  const cRuta =
    CR_col_(
      h,
      'Ruta'
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const vendedor =
      CR_txt_(
        data[i][
          cVendedor
        ]
      );


    const ruta =
      CR_txt_(
        data[i][
          cRuta
        ]
      );


    if (
      !vendedor ||
      !ruta
    ) {

      continue;

    }


    resultado[
      vendedor.toLowerCase()
    ] =
      ruta;

  }


  return resultado;

}


// ============================================================================
// ESTRUCTURAS VACÍAS
// ============================================================================

function CR_ventaVacia_() {

  return {

    Ruta:
      '',

    id_vendedor:
      '',

    Agua_Venta_Reparto:
      0,

    SuperIce_Venta_Reparto:
      0,

    Venta_Total_Dinero:
      0,

    Comision_Total:
      0,

    folios:
      new Set()

  };

}


function CR_cargaVacia_() {

  return {

    Ruta:
      '',

    id_vendedor:
      '',

    folios:
      new Set(),

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

}


// ============================================================================
// VALIDAR CIERRES_REPARTO
// ============================================================================

function CR_validarCierres_(
  h
) {

  [

    'id_cierre',
    'Fecha',
    'Ruta',
    'id_vendedor',
    'Total_Cargas',
    'Agua_Cargada',
    'Agua_Venta_OPSU',
    'Agua_Venta_Reparto',
    'Agua_Devuelta',
    'Agua_Diferencia',
    'SuperIce_Cargado',
    'SuperIce_Venta_OPSU',
    'SuperIce_Venta_Reparto',
    'SuperIce_Devuelto',
    'SuperIce_Diferencia',
    'Venta_Total_Dinero',
    'Comision_Total',
    'Cuadre',
    'FechaHora_Cierre',
    'Estatus',
    'Archivo'

  ].forEach(
    campo =>
      CR_col_(
        h,
        campo
      )
  );

}


// ============================================================================
// UTILIDADES
// ============================================================================

function CR_headers_(
  fila
) {

  const h = {};


  fila.forEach(
    (valor, i) => {

      const key =
        CR_norm_(
          valor
        );


      if (
        key
      ) {

        h[
          key
        ] =
          i;

      }

    }
  );


  return h;

}


function CR_col_(
  h,
  nombre
) {

  const key =
    CR_norm_(
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


function CR_colMulti_(
  h,
  nombres
) {

  for (
    const nombre
    of nombres
  ) {

    const key =
      CR_norm_(
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
    'No encontré ninguna de estas columnas: ' +
    nombres.join(
      ' / '
    )
  );

}


function CR_set_(
  fila,
  h,
  campo,
  valor
) {

  fila[
    CR_col_(
      h,
      campo
    )
  ] =
    valor;

}


function CR_norm_(
  valor
) {

  return String(
    valor ??
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


function CR_txt_(
  valor
) {

  return String(
    valor ??
    ''
  ).trim();

}


function CR_num_(
  valor
) {

  if (
    typeof valor ===
    'number'
  ) {

    return Number.isFinite(
      valor
    )
      ? valor
      : 0;

  }


  const n =
    Number(
      String(
        valor ??
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


function CR_fechaKey_(
  fecha
) {

  return Utilities
    .formatDate(
      fecha,
      CR_CFG.TZ,
      'yyyy-MM-dd'
    );

}


function CR_fechaKeySeguro_(
  valor
) {

  if (
    valor instanceof Date &&
    !isNaN(
      valor.getTime()
    )
  ) {

    return CR_fechaKey_(
      valor
    );

  }


  const txt =
    CR_txt_(
      valor
    );


  if (
    !txt
  ) {

    return '';

  }


  let m =
    txt.match(
      /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/
    );


  if (
    m
  ) {

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


  m =
    txt.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
    );


  if (
    m
  ) {

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


  return '';

}


function CR_fechaDesdeKey_(
  key
) {

  const p =
    String(
      key
    ).split(
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


function CR_clave_(
  fechaKey,
  ruta,
  vendedor
) {

  return [
    CR_txt_(
      fechaKey
    ),
    CR_txt_(
      ruta
    ),
    CR_txt_(
      vendedor
    ).toLowerCase()
  ].join(
    '||'
  );

}


function CR_nuevoIdCierre_(
  fechaKey,
  ruta
) {

  const fecha =
    String(
      fechaKey
    ).replace(
      /-/g,
      ''
    );


  const r =
    CR_txt_(
      ruta
    )
      .replace(
        /[^a-zA-Z0-9]/g,
        ''
      ) ||
    'SR';


  return (
    'CR-' +
    fecha +
    '-' +
    r +
    '-' +
    Utilities
      .getUuid()
      .substring(
        0,
        8
      )
  );

}


function CR_moneda_(
  valor
) {

  return '$' +
    Number(
      valor ||
      0
    ).toLocaleString(
      'es-MX',
      {
        minimumFractionDigits:
          2,

        maximumFractionDigits:
          2
      }
    );

}
