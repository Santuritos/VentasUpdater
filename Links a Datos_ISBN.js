/**
 * =============================================================
 * Sincronización de Ventas, Proveedores y Enlaces Internos
 * -------------------------------------------------------------
 * Versión:           1.4
 * Fecha:             15 abr 2025 (modificado)
 * Autor:             TuNombre o ChatGPT
 * Descripción:       Este script sincroniza los datos de la hoja "Ventas"
 *                    con las hojas destino: "Datos_ISBN", "Hamelyn", 
 *                    "Momox", "Ammareal", "Ebay" y "Todocoleccion". Se actualiza
 *                    la columna de Posicion (o Posición) y, en las hojas de proveedores,
 *                    también la columna Precio. Además, se actualiza la columna "ID"
 *                    de cada hoja para que contenga un enlace interno que lleve al
 *                    registro correspondiente en Datos_ISBN. La comparación de claves
 *                    se hace de forma insensible a espacios y mayúsculas/minúsculas.
 * =============================================================
 */

/**  
 * Normaliza una clave: la convierte a cadena, le quita espacios y la pone en mayúsculas.
 * @param {*} key  
 * @return {string}
 */
function normalizeKey(key) {
  return String(key).trim().toUpperCase();
}

/**
 * Devuelve el índice (0-based) de una columna cuyo nombre sea colName,
 * comprobando también la variante con acento ("Posición") cuando se busque "Posicion".
 * @param {Array} headers - Array con los encabezados.
 * @param {string} colName - Nombre buscado.
 * @return {number} Índice o -1 si no se encuentra.
 */
function getHeaderIndex(headers, colName) {
  var idx = headers.indexOf(colName);
  if (idx < 0 && colName.toLowerCase() === "posicion") {
    idx = headers.indexOf("Posición");
  }
  return idx;
}

/**
 * onEdit: Se dispara automáticamente cuando se edita una celda en alguna de las hojas
 * permitidas: Ventas, Datos_ISBN, Hamelyn, Momox, Ammareal, Ebay y Todocoleccion.
 * Actualiza la celda correspondiente en las hojas destino según la lógica definida.
 */
function onEdit(e) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("IS_SYNCING") === "true") return;
  props.setProperty("IS_SYNCING", "true");
  try {
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var allowedSheets = ["Ventas", "Datos_ISBN", "Hamelyn", "Momox", "Ammareal", "Ebay", "Todocoleccion"];
    if (allowedSheets.indexOf(sheetName) === -1) return;
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Selecciona la clave: En proveedores (Hamelyn, Momox, Ammareal) se usa "ISBN13"; en el resto, "ISBN".
    var keyField = (["Hamelyn", "Momox", "Ammareal"].indexOf(sheetName) !== -1) ? "ISBN13" : "ISBN";
    var keyIndex = headers.indexOf(keyField);
    var posIndex = getHeaderIndex(headers, "Posicion");
    if (keyIndex < 0 || posIndex < 0) {
      Logger.log("No se encontró " + keyField + " o Posicion en " + sheetName);
      return;
    }
    
    var editedRow = e.range.getRow();
    var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    var keyVal = rowData[keyIndex];
    var posicionVal = rowData[posIndex];
    if (!keyVal) return;
    keyVal = normalizeKey(keyVal);
    
    if (sheetName === "Ventas") {
      var precioHamelynIndex = headers.indexOf("Precio Hamelyn");
      var precioMomoxIndex = headers.indexOf("Precio Momox");
      var precioAmmarealIndex = headers.indexOf("Precio Ammareal");
      var precioHamelynVal = (precioHamelynIndex >= 0) ? rowData[precioHamelynIndex] : null;
      var precioMomoxVal = (precioMomoxIndex >= 0) ? rowData[precioMomoxIndex] : null;
      var precioAmmarealVal = (precioAmmarealIndex >= 0) ? rowData[precioAmmarealIndex] : null;
      
      updateCell("Datos_ISBN", "Posicion", keyVal, posicionVal);
      updateCell("Hamelyn", "Posición", keyVal, posicionVal);
      if (precioHamelynVal !== "" && precioHamelynVal !== null) {
        updateCell("Hamelyn", "Precio", keyVal, precioHamelynVal);
      }
      updateCell("Momox", "Posición", keyVal, posicionVal);
      if (precioMomoxVal !== "" && precioMomoxVal !== null) {
        updateCell("Momox", "Precio", keyVal, precioMomoxVal);
      }
      updateCell("Ammareal", "Posición", keyVal, posicionVal);
      if (precioAmmarealVal !== "" && precioAmmarealVal !== null) {
        updateCell("Ammareal", "Precio", keyVal, precioAmmarealVal);
      }
    } else if (sheetName === "Datos_ISBN") {
      updateCell("Ventas", "Posicion", keyVal, posicionVal);
      updateCell("Hamelyn", "Posición", keyVal, posicionVal);
      updateCell("Momox", "Posición", keyVal, posicionVal);
      updateCell("Ammareal", "Posición", keyVal, posicionVal);
      updateCell("Ebay", "Posicion", keyVal, posicionVal);
      updateCell("Todocoleccion", "Posicion", keyVal, posicionVal);
    } else if (["Hamelyn", "Momox", "Ammareal"].indexOf(sheetName) !== -1) {
      var precioIndex = headers.indexOf("Precio");
      if (precioIndex < 0) return;
      var precioVal = rowData[precioIndex];
      
      updateCell("Ventas", "Posicion", keyVal, posicionVal);
      if (sheetName === "Hamelyn") {
        updateCell("Ventas", "Precio Hamelyn", keyVal, precioVal);
      } else if (sheetName === "Momox") {
        updateCell("Ventas", "Precio Momox", keyVal, precioVal);
      } else if (sheetName === "Ammareal") {
        updateCell("Ventas", "Precio Ammareal", keyVal, precioVal);
      }
      updateCell("Datos_ISBN", "Posicion", keyVal, posicionVal);
      var otherProviders = ["Hamelyn", "Momox", "Ammareal"].filter(function(name) { return name !== sheetName; });
      otherProviders.forEach(function(prov) {
        updateCell(prov, "Posición", keyVal, posicionVal);
      });
    } else if (sheetName === "Ebay" || sheetName === "Todocoleccion") {
      updateCell("Datos_ISBN", "Posicion", keyVal, posicionVal);
      updateCell("Ventas", "Posicion", keyVal, posicionVal);
    }
    
  } catch (err) {
    Logger.log("Error en onEdit: " + err);
  } finally {
    PropertiesService.getScriptProperties().deleteProperty("IS_SYNCING");
  }
}

/**
 * updateCell: Busca en la hoja destino las filas con clave igual (normalizada)
 * y actualiza la columna indicada, siempre que el nuevo valor no sea vacío.
 * Actualiza todas las coincidencias.
 *
 * @param {string} sheetName - Nombre de la hoja destino.
 * @param {string} colName - Nombre de la columna a actualizar ("Posicion", "Precio", etc.).
 * @param {string} normKey - Clave normalizada (ISBN o ISBN13).
 * @param {*} newValue - Nuevo valor a asignar.
 */
function updateCell(sheetName, colName, normKey, newValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Hoja no encontrada: " + sheetName);
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var keyField = (["Hamelyn", "Momox", "Ammareal"].indexOf(sheetName) !== -1) ? "ISBN13" : "ISBN";
  var keyIndex = headers.indexOf(keyField);
  var targetIndex = headers.indexOf(colName);
  if (targetIndex < 0 && colName.toLowerCase() === "posicion") {
    targetIndex = headers.indexOf("Posición");
  }
  if (keyIndex < 0 || targetIndex < 0) {
    Logger.log("No se encontró " + keyField + " o " + colName + " en " + sheetName);
    return;
  }
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn());
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var cellKey = data[i][keyIndex];
    if (cellKey && normalizeKey(cellKey) === normKey) {
      if (newValue !== "" && newValue !== null) {
        sheet.getRange(i + 2, targetIndex + 1).setValue(newValue);
      }
      // No se hace break para actualizar todas las coincidencias.
    }
  }
}

/**
 * syncAllVentasOptimized: Realiza una sincronización global.
 * Recorre todos los registros de "Ventas" y actualiza las hojas destino:
 * Datos_ISBN, Hamelyn, Momox, Ammareal, Ebay y Todocoleccion.
 */
function syncAllVentasOptimized() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = ss.getSheetByName("Ventas");
  var ventasData = ventasSheet.getDataRange().getValues();
  if (ventasData.length < 2) return;
  
  var ventasHeaders = ventasData[0];
  var ventasKeyIndex = ventasHeaders.indexOf("ISBN");
  var ventasPosIndex = getHeaderIndex(ventasHeaders, "Posicion");
  var precioHamelynIndex = ventasHeaders.indexOf("Precio Hamelyn");
  var precioMomoxIndex = ventasHeaders.indexOf("Precio Momox");
  var precioAmmarealIndex = ventasHeaders.indexOf("Precio Ammareal");
  
  if (ventasKeyIndex < 0 || ventasPosIndex < 0) {
    Logger.log("No se encontró 'ISBN' o 'Posicion' en Ventas.");
    return;
  }
  
  function safeValue(val) {
    return (val === "" || val === null) ? undefined : val;
  }
  
  var ventasDict = {};
  for (var i = 1; i < ventasData.length; i++) {
    var row = ventasData[i];
    var key = row[ventasKeyIndex];
    if (key) {
      var normKey = normalizeKey(key);
      ventasDict[normKey] = {
        posicion: safeValue(row[ventasPosIndex]),
        precioHamelyn: safeValue(row[precioHamelynIndex]),
        precioMomox: safeValue(row[precioMomoxIndex]),
        precioAmmareal: safeValue(row[precioAmmarealIndex])
      };
    }
  }
  
  var targetSheets = [
    { name: "Datos_ISBN", keyField: "ISBN", update: { posicion: "Posicion" } },
    { name: "Hamelyn", keyField: "ISBN13", update: { posicion: "Posición", precio: "Precio" } },
    { name: "Momox", keyField: "ISBN13", update: { posicion: "Posición", precio: "Precio" } },
    { name: "Ammareal", keyField: "ISBN13", update: { posicion: "Posición", precio: "Precio" } },
    { name: "Ebay", keyField: "ISBN", update: { posicion: "Posicion" } },
    { name: "Todocoleccion", keyField: "ISBN", update: { posicion: "Posicion" } }
  ];
  
  targetSheets.forEach(function(target) {
    var sheet = ss.getSheetByName(target.name);
    if (!sheet) {
      Logger.log("Hoja no encontrada: " + target.name);
      return;
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    var headers = data[0];
    var keyIdx = headers.indexOf(target.keyField);
    var posIdx = getHeaderIndex(headers, target.update.posicion);
    var precioIdx = -1;
    if (target.update.precio) {
      precioIdx = headers.indexOf(target.update.precio);
    }
    if (keyIdx === -1 || posIdx === -1) {
      Logger.log("Faltan encabezados en " + target.name);
      return;
    }
    
    for (var j = 1; j < data.length; j++) {
      var row = data[j];
      var key = row[keyIdx];
      if (!key) continue;
      var normKey = normalizeKey(key);
      if (ventasDict.hasOwnProperty(normKey)) {
        var newPos = ventasDict[normKey].posicion;
        if (newPos !== undefined) {
          row[posIdx] = newPos;
        }
        if (precioIdx !== -1) {
          if (target.name === "Hamelyn" && ventasDict[normKey].precioHamelyn !== undefined) {
            row[precioIdx] = ventasDict[normKey].precioHamelyn;
          } else if (target.name === "Momox" && ventasDict[normKey].precioMomox !== undefined) {
            row[precioIdx] = ventasDict[normKey].precioMomox;
          } else if (target.name === "Ammareal" && ventasDict[normKey].precioAmmareal !== undefined) {
            row[precioIdx] = ventasDict[normKey].precioAmmareal;
          }
        }
      } else {
        Logger.log("Clave no encontrada en Ventas para " + target.name + ": " + key);
      }
    }
    
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  });
  
  Logger.log("syncAllVentasOptimized completado.");
}

/**
 * getDatosISBNRow: Busca en Datos_ISBN la fila (1-based) cuyo "ISBN" normalizado coincide con normKey.
 * @param {string} normKey - Clave normalizada.
 * @return {number|null} Número de fila o null si no se encuentra.
 */
function getDatosISBNRow(normKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datosSheet = ss.getSheetByName("Datos_ISBN");
  if (!datosSheet) return null;
  var data = datosSheet.getDataRange().getValues();
  if (data.length < 2) return null;
  
  var headers = data[0];
  var isbnIndex = headers.indexOf("ISBN");
  if (isbnIndex == -1) return null;
  
  for (var i = 1; i < data.length; i++) {
    var cellVal = data[i][isbnIndex];
    if (cellVal && normalizeKey(cellVal) === normKey) {
      return i + 1; // Filas 1-based.
    }
  }
  return null;
}

/**
 * updateIDLinksForSheet: Actualiza la columna "ID" de la hoja sheetName para agregar
 * un enlace interno que lleve al registro correspondiente en Datos_ISBN.
 * @param {string} sheetName - Nombre de la hoja a procesar.
 * @param {string} ssId - ID del Spreadsheet.
 * @param {number} datosGid - GID de la hoja Datos_ISBN.
 */
function updateIDLinksForSheet(sheetName, ssId, datosGid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  var keyField = (["Hamelyn", "Momox", "Ammareal"].indexOf(sheetName) !== -1) ? "ISBN13" : "ISBN";
  var headers = data[0];
  var keyIndex = headers.indexOf(keyField);
  var idIndex = headers.indexOf("ID");
  if (idIndex === -1) {
    Logger.log("La hoja " + sheetName + " no tiene la columna 'ID'.");
    return;
  }
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var keyVal = row[keyIndex];
    if (!keyVal) continue;
    var normKey = normalizeKey(keyVal);
    var datosRow = getDatosISBNRow(normKey);
    if (!datosRow) continue;
    
    // Construye la URL interna a la hoja Datos_ISBN (en la celda A{datosRow})
    var url = "https://docs.google.com/spreadsheets/d/" + ssId + "/edit#gid=" + datosGid + "&range=A" + datosRow;
    
    var cell = sheet.getRange(i + 1, idIndex + 1);
    var displayText = cell.getValue();
    if (!displayText) displayText = normKey;
    
    // Usa RichTextValue para agregar el enlace sin necesidad de fórmula.
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(displayText.toString())
      .setLinkUrl(url)
      .build();
    cell.setRichTextValue(richText);
  }
}

/**
 * updateIDLinksAcrossSheets: Recorre las hojas especificadas y actualiza la columna "ID"
 * para que contenga enlaces internos a Datos_ISBN.
 */
function updateIDLinksAcrossSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var datosSheet = ss.getSheetByName("Datos_ISBN");
  if (!datosSheet) {
    Logger.log("No se encontró la hoja Datos_ISBN.");
    return;
  }
  var datosGid = datosSheet.getSheetId();
  
  var sheetsToProcess = ["Ventas", "Datos_ISBN", "Hamelyn", "Momox", "Ammareal", "Ebay", "Todocoleccion"];
  sheetsToProcess.forEach(function(sheetName) {
    updateIDLinksForSheet(sheetName, ssId, datosGid);
  });
  Logger.log("Actualización de links en ID completada.");
}
