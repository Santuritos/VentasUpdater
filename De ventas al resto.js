/**
 * =============================================================
 * Sincronización de Ventas y Proveedores (incluye Ebay y Todocoleccion)
 * -------------------------------------------------------------
 * Versión:           1.3
 * Fecha:             15 abr 2025 (modificado)
 * Autor:             TuNombre o ChatGPT
 * Descripción:       Este script sincroniza los datos de la hoja "Ventas"
 *                    con las hojas destino: "Datos_ISBN", "Hamelyn", 
 *                    "Momox", "Ammareal", y además actualiza la columna 
 *                    Posicion en las hojas "Ebay" y "Todocoleccion" (sólo esa
 *                    columna, dejando intacto el resto) tomando la información
 *                    de Datos_ISBN.
 * =============================================================
 */

/**
 * Función de normalización para claves:
 * Convierte la clave a cadena, elimina espacios y la pasa a mayúsculas.
 */
function normalizeKey(key) {
  return String(key).trim().toUpperCase();
}

/**
 * Función auxiliar para obtener el índice de la columna "Posicion"
 * comprobando tanto "Posicion" como "Posición".
 *
 * @param {Array} headers Array con los encabezados.
 * @param {string} colName Nombre buscado (usualmente "Posicion").
 * @return {number} Índice (0-based) o -1 si no se encuentra.
 */
function getHeaderIndex(headers, colName) {
  var idx = headers.indexOf(colName);
  if (idx < 0 && colName.toLowerCase() === "posicion") {
    idx = headers.indexOf("Posición");
  }
  return idx;
}

/**
 * Función onEdit que se dispara al editar una celda en las hojas involucradas.
 * Se permiten: "Ventas", "Datos_ISBN", "Hamelyn", "Momox", "Ammareal", "Ebay" y "Todocoleccion".
 * Actualiza la celda correspondiente según la información de la fila editada.
 */
function onEdit(e) {
  // Evita recursividad
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("IS_SYNCING") === "true") return;
  props.setProperty("IS_SYNCING", "true");
  
  try {
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var allowedSheets = ["Ventas", "Datos_ISBN", "Hamelyn", "Momox", "Ammareal", "Ebay", "Todocoleccion"];
    if (allowedSheets.indexOf(sheetName) === -1) return;
    
    // Lee la fila de encabezados de la hoja editada
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Define la clave según la hoja:
    // Para Ventas y Datos_ISBN y para Ebay y Todocoleccion se usa "ISBN"
    // Para las hojas de proveedores se usa "ISBN13"
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
    
    // Según la hoja editada, actualiza otras hojas
    if (sheetName === "Ventas") {
      // En Ventas se esperan: "ISBN", "Posicion", "Precio Hamelyn", "Precio Momox", "Precio Ammareal"
      var precioHamelynIndex = headers.indexOf("Precio Hamelyn");
      var precioMomoxIndex = headers.indexOf("Precio Momox");
      var precioAmmarealIndex = headers.indexOf("Precio Ammareal");
      var precioHamelynVal = (precioHamelynIndex >= 0) ? rowData[precioHamelynIndex] : null;
      var precioMomoxVal = (precioMomoxIndex >= 0) ? rowData[precioMomoxIndex] : null;
      var precioAmmarealVal = (precioAmmarealIndex >= 0) ? rowData[precioAmmarealIndex] : null;
      
      // Actualiza desde Ventas a Datos_ISBN y proveedores
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
      
      // Si se edita Ventas, es posible que Datos_ISBN se haya actualizado previamente;
      // en todo caso, no se sincroniza directamente Ebay y Todocoleccion desde Ventas
      
    } else if (sheetName === "Datos_ISBN") {
      // Datos_ISBN usa "ISBN" y "Posicion". Se actualiza a Ventas, proveedores, Ebay y Todocoleccion.
      updateCell("Ventas", "Posicion", keyVal, posicionVal);
      updateCell("Hamelyn", "Posición", keyVal, posicionVal);
      updateCell("Momox", "Posición", keyVal, posicionVal);
      updateCell("Ammareal", "Posición", keyVal, posicionVal);
      updateCell("Ebay", "Posicion", keyVal, posicionVal);
      updateCell("Todocoleccion", "Posicion", keyVal, posicionVal);
      
    } else if (["Hamelyn", "Momox", "Ammareal"].indexOf(sheetName) !== -1) {
      // Hojas de proveedores usan "ISBN13", "Posición" y "Precio"
      var precioIndex = headers.indexOf("Precio");
      if (precioIndex < 0) return;
      var precioVal = rowData[precioIndex];
      
      // Actualiza a Ventas y Datos_ISBN
      updateCell("Ventas", "Posicion", keyVal, posicionVal);
      if (sheetName === "Hamelyn") {
        updateCell("Ventas", "Precio Hamelyn", keyVal, precioVal);
      } else if (sheetName === "Momox") {
        updateCell("Ventas", "Precio Momox", keyVal, precioVal);
      } else if (sheetName === "Ammareal") {
        updateCell("Ventas", "Precio Ammareal", keyVal, precioVal);
      }
      updateCell("Datos_ISBN", "Posicion", keyVal, posicionVal);
      
      // Actualiza la Posición en las otras hojas de proveedor
      var otherProviders = ["Hamelyn", "Momox", "Ammareal"].filter(function(name) {
        return name !== sheetName;
      });
      otherProviders.forEach(function(prov) {
        updateCell(prov, "Posición", keyVal, posicionVal);
      });
      
    } else if (sheetName === "Ebay" || sheetName === "Todocoleccion") {
      // Estas hojas usan "ISBN" y "Posicion".
      // Normalmente se actualizan desde Datos_ISBN, pero si se edita directamente se propaga a Datos_ISBN y Ventas.
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
 * Función updateCell:
 * Busca en la hoja destino (usando la clave) la fila correspondiente y actualiza
 * la columna indicada solo si el nuevo valor no está vacío.
 *
 * @param {string} sheetName - Nombre de la hoja destino.
 * @param {string} colName - Columna a actualizar (ej.: "Posición" o "Posicion", "Precio", etc.).
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
  // Define la clave según la hoja destino: "ISBN13" en proveedores, "ISBN" en el resto.
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
      // Se elimina el "break" para actualizar todas las coincidencias.
    }
  }
}


/**
 * Función global optimizada para sincronizar TODOS los registros de la hoja "Ventas"
 * con las hojas destino: Datos_ISBN, Hamelyn, Momox, Ammareal, Ebay y Todocoleccion.
 * Se actualiza la columna Posicion (y en proveedores, también el Precio) solamente
 * para las filas que tienen coincidencia en Ventas.
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
  
  // Construye el diccionario de Ventas normalizado
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
  
  // Configuración de hojas destino. Ahora se incluyen Ebay y Todocoleccion.
  // Para Ventas, Datos_ISBN, Ebay y Todocoleccion se usa la clave "ISBN"; para proveedores se usa "ISBN13".
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
    
    // Escribe de vuelta la matriz completa en la hoja destino
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  });
  
  Logger.log("syncAllVentasOptimized completado.");
}
