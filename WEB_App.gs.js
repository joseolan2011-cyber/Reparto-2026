// ============================================================
// WEB_App.gs
// APPWEB ANALÍTICA - REPARTO 2026
// ============================================================
// Este módulo pertenece exclusivamente a la AppWeb.
// NO modifica información de las hojas.
// NO interfiere con doPost() ni con Telegram / Make.
// ============================================================

const WEB_VERSION = '0.2.1';


/**
 * Entrada principal de la AppWeb.
 *
 * Index sigue siendo el archivo principal.
 * AtlasLocal se inyecta al final para añadir el motor local sin tener que
 * reescribir el Index monolítico.
 *
 * También se inyecta una etiqueta discreta con la versión actual en la
 * esquina superior derecha del reporte.
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


  const versionBadge =
    '<style>' +
      '#atlas-version-badge{' +
        'position:fixed;' +
        'top:10px;' +
        'right:14px;' +
        'z-index:9999;' +
        'font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;' +
        'font-size:10px;' +
        'font-weight:700;' +
        'letter-spacing:.3px;' +
        'color:#758096;' +
        'background:rgba(255,255,255,.88);' +
        'border:1px solid #e5e9f0;' +
        'border-radius:999px;' +
        'padding:4px 8px;' +
        'box-shadow:0 2px 8px rgba(15,23,42,.06);' +
        'pointer-events:none;' +
      '}' +
      '@media(max-width:700px){' +
        '#atlas-version-badge{' +
          'top:8px;' +
          'right:8px;' +
          'font-size:9px;' +
          'padding:3px 7px;' +
        '}' +
      '}' +
    '</style>' +
    '<div id="atlas-version-badge">v' +
      WEB_VERSION +
    '</div>';


  let html =
    index;


  if (
    html.includes(
      '</body>'
    )
  ) {

    html =
      html.replace(
        '</body>',
        atlasLocal +
        '\n' +
        versionBadge +
        '\n</body>'
      );

  }

  else {

    html +=
      atlasLocal +
      '\n' +
      versionBadge;

  }


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
