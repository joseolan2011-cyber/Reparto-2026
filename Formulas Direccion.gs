function OBTENER_DIRECCION(latlong) {
  var geo = Maps.newGeocoder().reverseGeocode(latlong.split(',')[0], latlong.split(',')[1]);
  if (geo.status == "OK") {
    return geo.results[0].formatted_address;
  }
  return "No encontrada";
}



/**
 * Separa el texto de la dirección SIN llamar al servicio de Google Maps.
 * @param {string} direccion El texto de la dirección que ya tienes.
 * @param {string} parte Qué dato extraer ("calle", "numero", "colonia").
 * @customfunction
 */
function EXTRAER_DIRECCION_TEXTO(direccion, parte) {
  if (!direccion || direccion === "") return "";
  
  // 1. Limpieza básica
  var limpia = direccion.replace(", Camp., Mexico", "").replace(", Mexico", "");
  var segmentos = limpia.split(",").map(s => s.trim());
  
  let calleNumero = "";
  let colonia = "";
  let numero = "";
  let calle = "";

  // 2. Lógica para detectar el formato (si empieza con CP o con Calle)
  if (segmentos[0].match(/^\d{5}/)) { 
    // Caso: "24155 Calle San Antonio Número 106..."
    calleNumero = segmentos[1] || segmentos[0];
    colonia = segmentos[2] || "";
  } else {
    // Caso estándar: "C. Laguna de Términos 12, 23 de Julio..."
    calleNumero = segmentos[0];
    colonia = segmentos[1] || "";
  }

  // 3. Separar Calle de Número de forma inteligente
  // Busca el último número en el segmento de la calle
  var matchNumero = calleNumero.match(/(.*?)(\d+[\s\w]*)$/);
  if (matchNumero) {
    calle = matchNumero[1].trim().replace(/,$/, "");
    numero = matchNumero[2].trim();
  } else {
    calle = calleNumero;
    numero = "S/N";
  }

  parte = parte.toLowerCase();
  if (parte === "calle") return calle;
  if (parte === "numero") return numero;
  if (parte === "colonia") return colonia;
  
  return "Error: parámetro incorrecto";
}
