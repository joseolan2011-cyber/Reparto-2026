// ============================================================
// WEB_App.gs
// APPWEB ANALÍTICA - REPARTO 2026
// ============================================================
// Este módulo pertenece exclusivamente a la AppWeb.
// NO modifica información de las hojas.
// NO interfiere con doPost() ni con Telegram / Make.
// ============================================================

const WEB_VERSION = '0.1.0';

/**
 * Entrada principal de la AppWeb.
 * Telegram continúa utilizando doPost(e) en el script existente.
 */
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Reparto 2026 | Análisis')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Permite incluir archivos HTML posteriormente si dividimos
 * CSS y JavaScript en archivos separados.
 */
function include(nombreArchivo) {
  return HtmlService
    .createHtmlOutputFromFile(nombreArchivo)
    .getContent();
}


/**
 * Información básica de la AppWeb.
 */
function WEB_getInfoApp() {
  return {
    ok: true,
    version: WEB_VERSION,
    nombre: 'Reparto 2026',
    modulo: 'Análisis',
    fechaServidor: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm:ss'
    )
  };
}
