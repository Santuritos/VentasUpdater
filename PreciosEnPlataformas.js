/**
 * Script: Precios en Plataformas - Amazon
 * Versión: 1.0
 * Fecha: 21 mar 2025
 *
 * Descripción: Este script consulta la API de Productos de Amazon para obtener el precio de un producto a partir de su ISBN y actualiza la columna "Precio_amazon" (columna AJ) en la hoja "Datos_ISBN".
 *
 * Credenciales de Amazon:
 *   - AMAZON_ACCESS_KEY: "AKIAIQFPS2B6YBNLO2YQ"
 *   - AMAZON_SECRET_KEY: "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"
 *   - AMAZON_ASSOCIATE_TAG: "a1500-21"
 */

var AMAZON_ACCESS_KEY   = "AKIAIQFPS2B6YBNLO2YQ";
var AMAZON_SECRET_KEY   = "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";
var AMAZON_ASSOCIATE_TAG = "a1500-21";

/**
 * getAmazonPrice(isbn)
 * Construye la consulta a la API de Amazon, firma la solicitud y extrae el precio del XML de respuesta.
 */
function getAmazonPrice(isbn) {
  // Parámetros y endpoint (versión antigua de la API)
  var endpoint = "webservices.amazon.com";
  var uri = "/onca/xml";
  
  var params = {
    "Service": "AWSECommerceService",
    "Operation": "ItemLookup",
    "AWSAccessKeyId": AMAZON_ACCESS_KEY,
    "AssociateTag": AMAZON_ASSOCIATE_TAG,
    "ItemId": isbn,
    "IdType": "EAN",  // Puedes cambiar a "ISBN" si es necesario
    "ResponseGroup": "Offers"
  };
  
  // Construir la cadena de consulta canónica ordenando los parámetros alfabéticamente
  var sortedKeys = Object.keys(params).sort();
  var canonicalQuery = sortedKeys.map(function(key) {
    return encodeURIComponent(key) + "=" + encodeURIComponent(params[key]);
  }).join("&");
  
  // Construir la cadena a firmar
  var stringToSign = "GET\n" + endpoint + "\n" + uri + "\n" + canonicalQuery;
  
  // Forzar que la clave secreta sea una cadena
  var secretKey = "" + AMAZON_SECRET_KEY;
  
  // Calcular la firma HMAC-SHA256
  var signatureBytes = Utilities.computeHmacSha256Signature(stringToSign, secretKey, Utilities.Charset.US_ASCII);
  var signature = encodeURIComponent(Utilities.base64Encode(signatureBytes));
  
  // Construir la URL final de la solicitud
  var requestUrl = "http://" + endpoint + uri + "?" + canonicalQuery + "&Signature=" + signature;
  Logger.log("Amazon Request URL: " + requestUrl);
  
  // Realizar la consulta
  var response = UrlFetchApp.fetch(requestUrl);
  var responseCode = response.getResponseCode();
  Logger.log("Amazon - Código de respuesta: " + responseCode);
  
  if (responseCode !== 200) {
    throw new Error("Amazon API error: " + responseCode);
  }
  
  // Parsear la respuesta XML
  var xmlText = response.getContentText();
  var document = XmlService.parse(xmlText);
  var root = document.getRootElement();
  var ns = root.getNamespace();
  
  // Navegar por el XML (la estructura puede variar según la respuesta)
  var items = root.getChild("Items", ns);
  if (!items) throw new Error("No se encontraron Items en la respuesta de Amazon.");
  
  var item = items.getChild("Item", ns);
  if (!item) throw new Error("No se encontró Item en la respuesta de Amazon.");
  
  var offers = item.getChild("Offers", ns);
  if (!offers) throw new Error("No se encontraron Offers en la respuesta de Amazon.");
  
  var offer = offers.getChild("Offer", ns);
  if (!offer) throw new Error("No se encontró Offer en la respuesta de Amazon.");
  
  var offerListing = offer.getChild("OfferListing", ns);
  if (!offerListing) throw new Error("No se encontró OfferListing en la respuesta de Amazon.");
  
  var priceElement = offerListing.getChild("Price", ns);
  if (!priceElement) throw new Error("No se encontró Price en la respuesta de Amazon.");
  
  var amountElement = priceElement.getChild("Amount", ns);
  if (!amountElement) throw new Error("No se encontró Amount en la respuesta de Amazon.");
  
  var priceText = amountElement.getText();
  // Suponiendo que el precio se devuelve en céntimos, convertir a euros/dólares según corresponda
  var priceValue = parseFloat(priceText) / 100;
  return priceValue;
}

/**
 * updateAmazonPrices()
 * Recorre la hoja "Datos_ISBN" y actualiza la columna AJ ("Precio_amazon") con el precio obtenido para cada ISBN.
 */
function updateAmazonPrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Datos_ISBN");
  if (!sheet) throw new Error("No se encontró la hoja 'Datos_ISBN'.");
  
  var lastRow = sheet.getLastRow();
  
  // Suponiendo que:
  //   - El ISBN está en la columna C (índice 3)
  //   - La columna AJ ("Precio_amazon") es la columna 36 (A=1, B=2, ..., AJ=36)
  for (var i = 2; i <= lastRow; i++) {
    var isbn = sheet.getRange(i, 3).getValue().toString().trim();
    if (!isbn) {
      Logger.log("Fila " + i + " sin ISBN.");
      continue;
    }
    try {
      Logger.log("Fila " + i + " - ISBN: " + isbn);
      var price = getAmazonPrice(isbn);
      Logger.log("Fila " + i + " - Amazon: " + price);
      sheet.getRange(i, 36).setValue(price);
    } catch (e) {
      Logger.log("Error en getAmazonPrice para ISBN " + isbn + ": " + e);
      sheet.getRange(i, 36).setValue("Error");
    }
  }
  Logger.log("Actualización de precios en Amazon completada.");
}
