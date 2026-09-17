// ============================================================================
// ZZ_SYNC_PROMOS_OPSU.gs
// EXTENSION CARGAS_OPSU - PROMOCIONES SUPERICE
// VERSION 1.0
// ============================================================================
//
// Objetivo:
//   Mantener separadas las ventas de SuperIce y las piezas entregadas como
//   promoción, pero considerar ambas como salida física para el cuadre OPSU.
//
// Flujo:
//   ENTREGAS_PROMO[id_venta]
//          -> VENTAS[id_venta]
//          -> VENTAS[IDCarga_ref]
//          -> CARGAS_OPSU[Folio]
//
// Solo se contabilizan promociones con Estatus = Autorizado.
// La comisión NO se modifica: continúa calculándose únicamente con VENTAS.
//
// Fórmula de conciliación:
//   SuperIce_Diferencia =
//     SuperIce_Venta_OPSU - SuperIce_Venta_Reparto - SuperIce_Promo_Reparto
//
// Este archivo extiende el recálculo físico existente sin modificar el archivo
// principal SYNC_CARGAS_OPSU.gs.js.
// ============================================================================

const CARGAS_PROMO_OPSU_VERSION = '1.0';

// Guardamos la implementación original y la extendemos.
// En Apps Script las funciones del proyecto comparten el mismo ámbito global.
if (typeof CARGAS_OPSU_recalcularVentasFisicasTabla_ === 'function') {
  var CARGAS_OPSU_recalcularVentasFisicasTabla_BASE_PROMO_ =
    CARGAS_OPSU_recalcularVentasFisicasTabla_;

  CARGAS_OPSU_recalcularVentasFisicasTabla_ = function(tabla, h) {
    // Primero conserva exactamente el cálculo actual de ventas/comisiones.
    CARGAS_OPSU_recalcularVentasFisicasTabla_BASE_PROMO_(tabla, h);

    // Después agrega las promociones de SuperIce al cuadre físico.
    CARGAS_PROMO_OPSU_aplicarPromosTabla_(tabla, h);
  };
}


function CARGAS_PROMO_OPSU_aplicarPromosTabla_(tabla, h) {
  if (!tabla || tabla.length < 2 || !h) return;

  const cFolio = CARGAS_PROMO_OPSU_col_(h, 'Folio');
  const cTipo = CARGAS_PROMO_OPSU_col_(h, 'Tipo_Operacion');
  const cAguaDif = CARGAS_PROMO_OPSU_col_(h, 'Agua_Diferencia');
  const cIceOpsu = CARGAS_PROMO_OPSU_col_(h, 'SuperIce_Venta_OPSU');
  const cIceVenta = CARGAS_PROMO_OPSU_col_(h, 'SuperIce_Venta_Reparto');
  const cIcePromo = CARGAS_PROMO_OPSU_col_(h, 'SuperIce_Promo_Reparto');
  const cIceDif = CARGAS_PROMO_OPSU_col_(h, 'SuperIce_Diferencia');
  const cCuadre = CARGAS_PROMO_OPSU_col_(h, 'Cuadre');

  const folios = new Set();

  for (let i = 1; i < tabla.length; i++) {
    const folio = CARGAS_PROMO_OPSU_txt_(tabla[i][cFolio]);
    const tipo = CARGAS_PROMO_OPSU_txt_(tabla[i][cTipo]);

    if (folio && tipo === 'Reparto') folios.add(folio);
  }

  const promosPorFolio = CARGAS_PROMO_OPSU_leerPromosPorFolio_(folios);

  for (let i = 1; i < tabla.length; i++) {
    const fila = tabla[i];
    const folio = CARGAS_PROMO_OPSU_txt_(fila[cFolio]);
    const tipo = CARGAS_PROMO_OPSU_txt_(fila[cTipo]);

    if (!folio) continue;

    if (tipo === 'Externo') {
      fila[cIcePromo] = 0;
      fila[cCuadre] = 'EXTERNO';
      continue;
    }

    if (tipo !== 'Reparto') continue;

    const promo = CARGAS_PROMO_OPSU_num_(promosPorFolio[folio] || 0);
    const iceOpsu = CARGAS_PROMO_OPSU_num_(fila[cIceOpsu]);
    const iceVenta = CARGAS_PROMO_OPSU_num_(fila[cIceVenta]);
    const aguaDif = CARGAS_PROMO_OPSU_num_(fila[cAguaDif]);

    const iceDif = iceOpsu - iceVenta - promo;

    fila[cIcePromo] = promo;
    fila[cIceDif] = iceDif;
    fila[cCuadre] = (aguaDif === 0 && iceDif === 0)
      ? 'CUADRA'
      : 'DIFERENCIA';
  }
}


function CARGAS_PROMO_OPSU_leerPromosPorFolio_(folios) {
  const resultado = {};
  if (!folios || !folios.size) return resultado;

  const ss = CARGAS_OPSU_getSpreadsheetReparto_();
  const shVentas = ss.getSheetByName('VENTAS');
  const shPromos = ss.getSheetByName('ENTREGAS_PROMO');

  if (!shVentas) throw new Error('No existe la hoja VENTAS en Reparto.');
  if (!shPromos) throw new Error('No existe la hoja ENTREGAS_PROMO en Reparto.');

  // --------------------------------------------------------------------------
  // 1) Mapa id_venta -> IDCarga_ref
  // --------------------------------------------------------------------------
  const mapaVentaFolio = {};
  const lrVentas = shVentas.getLastRow();
  const lcVentas = shVentas.getLastColumn();

  if (lrVentas >= 2 && lcVentas >= 1) {
    const hv = shVentas.getRange(1, 1, 1, lcVentas).getValues()[0];
    const ivId = CARGAS_PROMO_OPSU_indiceHeader_(hv, 'id_venta');
    const ivFolio = CARGAS_PROMO_OPSU_indiceHeader_(hv, 'IDCarga_ref');

    const filasVentas = shVentas
      .getRange(2, 1, lrVentas - 1, lcVentas)
      .getValues();

    for (let i = 0; i < filasVentas.length; i++) {
      const idVenta = CARGAS_PROMO_OPSU_txt_(filasVentas[i][ivId]);
      const folio = CARGAS_PROMO_OPSU_txt_(filasVentas[i][ivFolio]);

      if (idVenta && folio && folios.has(folio)) {
        mapaVentaFolio[idVenta] = folio;
      }
    }
  }

  // --------------------------------------------------------------------------
  // 2) Sumar piezas autorizadas por folio de carga
  // --------------------------------------------------------------------------
  const lrPromos = shPromos.getLastRow();
  const lcPromos = shPromos.getLastColumn();

  if (lrPromos < 2 || lcPromos < 1) return resultado;

  const hp = shPromos.getRange(1, 1, 1, lcPromos).getValues()[0];
  const ipVenta = CARGAS_PROMO_OPSU_indiceHeader_(hp, 'id_venta');
  const ipPiezas = CARGAS_PROMO_OPSU_indiceHeader_(hp, 'piezas_entregadas');
  const ipEstatus = CARGAS_PROMO_OPSU_indiceHeader_(hp, 'Estatus');

  const filasPromos = shPromos
    .getRange(2, 1, lrPromos - 1, lcPromos)
    .getValues();

  for (let i = 0; i < filasPromos.length; i++) {
    const estatus = CARGAS_PROMO_OPSU_norm_(filasPromos[i][ipEstatus]);
    if (estatus !== 'autorizado') continue;

    const idVenta = CARGAS_PROMO_OPSU_txt_(filasPromos[i][ipVenta]);
    const folio = mapaVentaFolio[idVenta];

    if (!folio || !folios.has(folio)) continue;

    if (!resultado[folio]) resultado[folio] = 0;

    resultado[folio] += CARGAS_PROMO_OPSU_num_(
      filasPromos[i][ipPiezas]
    );
  }

  return resultado;
}


function CARGAS_PROMO_OPSU_indiceHeader_(headers, nombre) {
  const buscado = CARGAS_PROMO_OPSU_norm_(nombre);

  for (let i = 0; i < headers.length; i++) {
    if (CARGAS_PROMO_OPSU_norm_(headers[i]) === buscado) return i;
  }

  throw new Error('Falta columna requerida: ' + nombre);
}


function CARGAS_PROMO_OPSU_col_(h, nombre) {
  const key = CARGAS_PROMO_OPSU_norm_(nombre);

  if (Object.prototype.hasOwnProperty.call(h, key)) {
    return h[key];
  }

  throw new Error('Falta columna requerida en CARGAS_OPSU: ' + nombre);
}


function CARGAS_PROMO_OPSU_txt_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}


function CARGAS_PROMO_OPSU_num_(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;

  const n = Number(
    String(v === null || v === undefined ? '' : v)
      .trim()
      .replace(/,/g, '')
  );

  return Number.isFinite(n) ? n : 0;
}


function CARGAS_PROMO_OPSU_norm_(v) {
  return CARGAS_PROMO_OPSU_txt_(v)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}


// Prueba manual opcional. Recalcula todas las cargas existentes usando la
// función ya extendida. No altera VENTAS ni ENTREGAS_PROMO.
function CARGAS_PROMO_OPSU_recalcularAhora() {
  CARGAS_OPSU_recalcularVentasAhora();

  console.log(
    'Promociones SuperIce aplicadas a CARGAS_OPSU. Versión ' +
    CARGAS_PROMO_OPSU_VERSION
  );
}
