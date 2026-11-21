/**
 * REPARTO 2026 - LIMPIEZA NOCTURNA DE FILAS COMPLETAMENTE VACÍAS
 * Archivo independiente; no cambia la lógica de ventas, pedidos o recorridos.
 *
 * Instalación automática: MONITOREO_CENTRAL_cadaMinuto llama a
 * LIMPIEZA_REPARTO_verificarInstalacion(), que sólo comprueba los triggers
 * una vez al día. La limpieza corre mediante un trigger nocturno separado.
 */
const LIMPIEZA_REPARTO_CONFIG = Object.freeze({
  spreadsheetId: '1sPVwm2CV2b4CwydTLunN_DKyoHnPLeUtfkgvPgxtPNw',
  zona: 'America/Merida',
  hora: 2,
  revision: 'LIMPIEZA_REPARTO_REVISION_TRIGGER',
  progreso: 'LIMPIEZA_REPARTO_PROGRESO',
  maxMilisegundos: 225000
});

/**
 * Instala los activadores faltantes sin afectar los existentes:
 * - Limpieza diaria entre 02:00 y 03:00, hora de Mérida.
 * - Menú "Limpieza" al abrir la hoja.
 * No borra datos durante la verificación.
 */
function LIMPIEZA_REPARTO_verificarInstalacion() {
  const config = LIMPIEZA_REPARTO_CONFIG;
  const hoy = Utilities.formatDate(new Date(), config.zona, 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(config.revision) === hoy) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    if (props.getProperty(config.revision) === hoy) return;

    const existentes = ScriptApp.getProjectTriggers()
      .map(trigger => trigger.getHandlerFunction());

    if (!existentes.includes('LIMPIEZA_REPARTO_ejecutar')) {
      ScriptApp.newTrigger('LIMPIEZA_REPARTO_ejecutar')
        .timeBased()
        .atHour(config.hora)
        .everyDays(1)
        .inTimezone(config.zona)
        .create();
      console.log('Limpieza REPARTO: activador nocturno instalado.');
    }

    if (!existentes.includes('LIMPIEZA_REPARTO_mostrarMenu')) {
      const libro = SpreadsheetApp.openById(config.spreadsheetId);
      ScriptApp.newTrigger('LIMPIEZA_REPARTO_mostrarMenu')
        .forSpreadsheet(libro)
        .onOpen()
        .create();
      console.log('Limpieza REPARTO: activador de menú instalado.');
    }

    props.setProperty(config.revision, hoy);

  } finally {
    lock.releaseLock();
  }
}

function LIMPIEZA_REPARTO_mostrarMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Limpieza')
    .addItem('Eliminar filas vacías ahora', 'LIMPIEZA_REPARTO_ejecutar')
    .addToUi();
}

/**
 * Borra filas vacías internas de todas las hojas, incluso las respaldadas.
 * No borra encabezados, filas congeladas, fórmulas ni filas de reserva
 * posteriores al último dato. 0 y FALSE cuentan como datos.
 *
 * Se procesan bloques de abajo hacia arriba. Si la ejecución se acerca
 * al límite, guarda progreso y continúa con un activador temporal.
 */
function LIMPIEZA_REPARTO_ejecutar() {
  const config = LIMPIEZA_REPARTO_CONFIG;
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    console.warn('Limpieza REPARTO omitida: otra operación tiene el bloqueo.');
    return;
  }

  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  const hoy = Utilities.formatDate(new Date(), config.zona, 'yyyy-MM-dd');
  let eliminadas = 0;
  let errores = 0;

  try {
    const libro = SpreadsheetApp.openById(config.spreadsheetId);
    const hojas = libro.getSheets();
    let progreso = null;

    try {
      progreso = JSON.parse(props.getProperty(config.progreso) || 'null');
    } catch (_) {
      progreso = null;
    }

    const desde = progreso && progreso.fecha === hoy &&
      Number.isInteger(progreso.indice) &&
      progreso.indice >= 0 && progreso.indice < hojas.length
      ? progreso.indice : 0;

    for (let i = desde; i < hojas.length; i++) {
      if (Date.now() - inicio > config.maxMilisegundos) {
        props.setProperty(config.progreso, JSON.stringify({
          fecha: hoy, indice: i
        }));
        LIMPIEZA_REPARTO_programarContinuacion();
        console.log('Limpieza REPARTO: continuará desde hoja ' + (i + 1));
        return;
      }

      const hoja = hojas[i];
      try {
        const borradas = LIMPIEZA_REPARTO_limpiarHoja(hoja);
        eliminadas += borradas;
        console.log('Limpieza REPARTO [' + hoja.getName() +
          ']: ' + borradas + ' filas eliminadas.');
      } catch (error) {
        errores++;
        console.error('Limpieza REPARTO [' + hoja.getName() +
          ']: ' + (error.stack || error.message));
      }

      props.setProperty(config.progreso, JSON.stringify({
        fecha: hoy, indice: i + 1
      }));
    }

    props.deleteProperty(config.progreso);

    // Un activador temporal puede borrarse después de completar el lote.
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'LIMPIEZA_REPARTO_continuar')
      .forEach(t => ScriptApp.deleteTrigger(t));

    console.log('Limpieza REPARTO completada. Eliminadas en esta ejecución: ' +
      eliminadas + '; hojas con error: ' + errores + '.');

  } finally {
    lock.releaseLock();
  }
}

function LIMPIEZA_REPARTO_limpiarHoja(hoja) {
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  const primeraFila = Math.max(2, hoja.getFrozenRows() + 1);

  if (ultimaFila < primeraFila || ultimaColumna === 0) return 0;

  const rango = hoja.getRange(
    primeraFila, 1, ultimaFila - primeraFila + 1, ultimaColumna
  );
  const valores = rango.getValues();
  const formulas = rango.getFormulas();

  let inicioBloque = 0;
  let cantidadBloque = 0;
  let eliminadas = 0;

  for (let i = valores.length - 1; i >= 0; i--) {
    const vacia = valores[i].every((valor, columna) => {
      const sinValor = valor === '' || valor === null ||
        valor === undefined ||
        (typeof valor === 'string' && valor.trim() === '');
      return sinValor && formulas[i][columna] === '';
    });

    if (vacia) {
      inicioBloque = primeraFila + i;
      cantidadBloque++;
    } else if (cantidadBloque > 0) {
      hoja.deleteRows(inicioBloque, cantidadBloque);
      eliminadas += cantidadBloque;
      cantidadBloque = 0;
    }
  }

  if (cantidadBloque > 0) {
    hoja.deleteRows(inicioBloque, cantidadBloque);
    eliminadas += cantidadBloque;
  }

  return eliminadas;
}

function LIMPIEZA_REPARTO_programarContinuacion() {
  if (ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'LIMPIEZA_REPARTO_continuar')) {
    return;
  }

  ScriptApp.newTrigger('LIMPIEZA_REPARTO_continuar')
    .timeBased()
    .after(120000)
    .create();
}

function LIMPIEZA_REPARTO_continuar() {
  LIMPIEZA_REPARTO_ejecutar();
}
