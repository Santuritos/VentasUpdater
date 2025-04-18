/**
 * Script Name: Actualizar Tabla Ebay (Con Posicion y Preservando Precio/Portes)
 * Version: 1.4
 * Autor: Tu Nombre o Nombre de la Empresa
 * Fecha: 2025-04-15
 * Descripción:
 *    Este script actualiza la hoja "Ebay" en función de los datos de la hoja "Datos_ISBN".
 *    Extrae el ID, ISBN, EAN, Titulo y Posicion, y genera dos enlaces para búsquedas en eBay España:
 *      - Enlace para buscar mediante el ISBN (o EAN si no existe ISBN).
 *      - Enlace para buscar mediante el Titulo.
 *    Los campos "Precio" y "Portes" se mantienen si ya tienen valores manuales,
 *    de lo contrario se dejan vacíos.
 */

function actualizarEbay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDatos = ss.getSheetByName("Datos_ISBN");
  var sheetEbay = ss.getSheetByName("Ebay");
  
  if (!sheetDatos || !sheetEbay) {
    Logger.log("Error: No se encontró alguna de las hojas: Datos_ISBN o Ebay.");
    return;
  }
  
  Logger.log("Hojas encontradas. Procesando datos...");
  
  // Leer datos de la hoja Datos_ISBN
  var datos = sheetDatos.getDataRange().getValues();
  if (datos.length < 2) {
    Logger.log("La hoja Datos_ISBN no contiene registros de datos.");
    return;
  }
  
  var datosHeaders = datos[0];
  
  // Buscar índices de columnas en Datos_ISBN:
  // Se requieren: "ID", "ISBN", "EAN", "Titulo" y "Posicion" (estos nombres deben coincidir exactamente)
  var idxID = datosHeaders.indexOf("ID");
  var idxISBN = datosHeaders.indexOf("ISBN");
  var idxEAN = datosHeaders.indexOf("EAN");
  var idxTitulo = datosHeaders.indexOf("Titulo");
  var idxPosicion = datosHeaders.indexOf("Posicion");
  
  if (idxID === -1 || idxISBN === -1 || idxEAN === -1 || idxTitulo === -1 || idxPosicion === -1) {
    Logger.log("Error: No se han encontrado uno o más campos requeridos en Datos_ISBN. Revisa los encabezados.");
    return;
  }
  
  // Leer la hoja Ebay existente para preservar valores manuales en "Precio" y "Portes"
  var ebayExistingData = sheetEbay.getDataRange().getValues();
  var manualEntries = {}; // Diccionario: id -> { precio, portes }
  if (ebayExistingData.length > 1) {
    var ebayHeaders = ebayExistingData[0];
    var ebayIdxID = ebayHeaders.indexOf("ID");
    var ebayIdxPrecio = ebayHeaders.indexOf("Precio");
    var ebayIdxPortes = ebayHeaders.indexOf("Portes");
    if (ebayIdxID != -1 && ebayIdxPrecio != -1 && ebayIdxPortes != -1) {
      for (var i = 1; i < ebayExistingData.length; i++) {
        var row = ebayExistingData[i];
        var idValue = row[ebayIdxID];
        manualEntries[idValue] = {
          precio: row[ebayIdxPrecio],
          portes: row[ebayIdxPortes]
        };
      }
    } else {
      Logger.log("Advertencia: No se pudieron encontrar las columnas Precio y Portes en la hoja Ebay existente.");
    }
  }
  
  // Nuevo encabezado para la hoja Ebay (con la columna Posicion incluida)
  var newEbayData = [];
  newEbayData.push(["ID", "ISBN", "EAN", "Titulo", "Posicion", "Precio", "Portes", "Enlace", "Enlace Titulo"]);
  
  // Definir la estructura de la URL de búsqueda en eBay España
  var baseUrl = "https://www.ebay.es/sch/i.html?_nkw=";
  var extraParams = "&_sacat=0&_from=R40&_trksid=m570.l1313";
  
  Logger.log("Procesando " + (datos.length - 1) + " registros.");
  
  // Procesar cada registro de Datos_ISBN
  for (var j = 1; j < datos.length; j++) {
    var row = datos[j];
    var id_val = row[idxID];
    var isbn_val = row[idxISBN];
    var ean_val = row[idxEAN];
    var titulo_val = row[idxTitulo];
    var posicion_val = row[idxPosicion];
    
    // Generar enlace usando ISBN (o, de lo contrario, EAN)
    var codigoBusqueda = isbn_val ? isbn_val : ean_val;
    var enlace = codigoBusqueda ? baseUrl + encodeURIComponent(codigoBusqueda) + extraParams : "";
    
    // Generar el enlace basado en Titulo
    var enlaceTitulo = titulo_val ? baseUrl + encodeURIComponent(titulo_val) + extraParams : "";
    
    // Recuperar los valores manuales de Precio y Portes, si existen para este ID
    var manualPrecio = "";
    var manualPortes = "";
    if (manualEntries.hasOwnProperty(id_val)) {
      manualPrecio = manualEntries[id_val].precio;
      manualPortes = manualEntries[id_val].portes;
    }
    
    newEbayData.push([id_val, isbn_val, ean_val, titulo_val, posicion_val, manualPrecio, manualPortes, enlace, enlaceTitulo]);
  }
  
  Logger.log("Número total de registros que se escribirán en Ebay: " + newEbayData.length);
  
  // Limpiar la hoja Ebay y escribir la nueva matriz actualizada
  sheetEbay.clearContents();
  sheetEbay.getRange(1, 1, newEbayData.length, newEbayData[0].length).setValues(newEbayData);
  
  Logger.log("La tabla Ebay se ha actualizado correctamente, preservando Precio y Portes.");
}
