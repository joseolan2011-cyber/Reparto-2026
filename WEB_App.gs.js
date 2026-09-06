// ============================================================
// WEB_App.gs
// APPWEB ANALÍTICA - REPARTO 2026
// ============================================================
// Este módulo pertenece exclusivamente a la AppWeb.
// NO modifica información de las hojas.
// NO interfiere con doPost() ni con Telegram / Make.
// ============================================================

const WEB_VERSION = '0.2.0';


/**
 * Entrada principal de la AppWeb.
 *
 * Index sigue siendo el archivo principal.
 * AtlasLocal se inyecta al final para añadir el motor local sin tener que
 * reescribir el Index monolítico.
 */
function doGet(e) {

  const index =
    HtmlService
      .createTemplateFromFile(
        'Index'
      )
      .evaluate()
      .getContent();


  const atlasLocal =
    HtmlService
      .createHtmlOutputFromFile(
        'AtlasLocal'
      )
      .getContent();


  const html =
    index.includes(
      '</body>'
    )
      ? index.replace(
          '</body>',
          atlasLocal +
          '\n</body>'
        )
      : index +
        atlasLocal;


  return HtmlService
    .createHtmlOutput(
      html
    )
    .setTitle(
      'Reparto 2026 | Análisis'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/**
 * Permite incluir archivos HTML posteriormente si dividimos
 * CSS y JavaScript en archivos separados.
 */
function include(
  nombreArchivo
) {

  return HtmlService
    .createHtmlOutputFromFile(
      nombreArchivo
    )
    .getContent();

}


/**
 * Información básica de la AppWeb.
 */
function WEB_getInfoApp() {

  return {

    ok:
      true,

    version:
      WEB_VERSION,

    nombre:
      'Reparto 2026',

    modulo:
      'Análisis',

    fechaServidor:
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'dd/MM/yyyy HH:mm:ss'
      )

  };

}
