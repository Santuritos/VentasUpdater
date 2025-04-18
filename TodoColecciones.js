/**
 * Script Name: Actualizar Tabla Todocoleccion (Con Posicion y Preservando Precio/Portes)
 * Version: 1.1
 * Autor: Tu Nombre o Nombre de la Empresa
 * Fecha: 2025-04-15
 * Descripción:
 *    Este script actualiza la hoja "Todocoleccion" en función de los datos de la hoja "Datos_ISBN".
 *    Extrae el ID, ISBN, EAN, Titulo y Posicion, y genera dos enlaces para búsquedas en Todocoleccion:
 *      - Enlace para buscar mediante el ISBN (o EAN si no existe ISBN).
 *      - Enlace para buscar mediante el Titulo.
 *    Los campos "Precio" y "Portes" se mantienen si ya tienen valores manuales;
 *    de lo contrario, se dejan vacíos.
 * 
 *    Se utiliza la URL base:
 *      "https://www.todocoleccion.net/buscador?from=top&bu="
 *    para construir los enlaces.
 */

function actualizarTodocoleccion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDatos = ss.getSheetByName("Datos_ISBN");
  var sheetTC = ss.getSheetByName("Todocoleccion");
  
  if (!sheetDatos || !sheetTC) {
    Logger.log("Error: No se encontró alguna de las hojas: Datos_ISBN o Todocoleccion.");
    return;
  }
  
  Logger.log("Hojas encontradas. Procesando datos...");
  
  // Leer todos los datos de la hoja Datos_ISBN
  var datos = sheetDatos.getDataRange().getValues();
  if (datos.length < 2) {
    Logger.log("La hoja Datos_ISBN no contiene registros de datos.");
    return;
  }
  
  var datosHeaders = datos[0];
  
  // Buscar índices de columnas en Datos_ISBN:
  // Se requieren: "ID", "ISBN", "EAN", "Titulo" y "Posicion"
  var idxID = datosHeaders.indexOf("ID");
  var idxISBN = datosHeaders.indexOf("ISBN");
  var idxEAN = datosHeaders.indexOf("EAN");
  var idxTitulo = datosHeaders.indexOf("Titulo");
  var idxPosicion = datosHeaders.indexOf("Posicion");
  
  if (idxID === -1 || idxISBN === -1 || idxEAN === -1 || idxTitulo === -1 || idxPosicion === -1) {
    Logger.log("Error: No se han encontrado uno o más campos requeridos en Datos_ISBN. Revisa los encabezados.");
    return;
  }
  
  // Leer la hoja Todocoleccion existente para preservar los valores manuales de Precio y Portes
  var tcExistingData = sheetTC.getDataRange().getValues();
  var manualEntries = {}; // Diccionario: id -> { precio, portes }
  if (tcExistingData.length > 1) {
    var tcHeaders = tcExistingData[0];
    var tcIdxID = tcHeaders.indexOf("ID");
    var tcIdxPrecio = tcHeaders.indexOf("Precio");
    var tcIdxPortes = tcHeaders.indexOf("Portes");
    if (tcIdxID != -1 && tcIdxPrecio != -1 && tcIdxPortes != -1) {
      for (var i = 1; i < tcExistingData.length; i++) {
        var row = tcExistingData[i];
        var idValue = row[tcIdxID];
        manualEntries[idValue] = {
          precio: row[tcIdxPrecio],
          portes: row[tcIdxPortes]
        };
      }
    } else {
      Logger.log("Advertencia: No se pudieron encontrar las columnas Precio y Portes en la hoja Todocoleccion existente.");
    }
  }
  
  // Nuevo encabezado para la hoja Todocoleccion (incluyendo la columna Posicion)
  var newTCData = [];
  newTCData.push(["ID", "ISBN", "EAN", "Titulo", "Posicion", "Precio", "Portes", "Enlace", "Enlace Titulo"]);
  
  // Definir la estructura de la URL de búsqueda en Todocoleccion (según la URL que has indicado)
  var baseUrl = "https://www.todocoleccion.net/buscador?from=top&bu=";
  var extraParams = ""; // Se pueden añadir parámetros adicionales si es necesario
  
  Logger.log("Procesando " + (datos.length - 1) + " registros.");
  
  // Procesar cada registro de Datos_ISBN (omitiendo la fila de encabezados)
  for (var j = 1; j < datos.length; j++) {
    var row = datos[j];
    var id_val = row[idxID];
    var isbn_val = row[idxISBN];
    var ean_val = row[idxEAN];
    var titulo_val = row[idxTitulo];
    var posicion_val = row[idxPosicion];
    
    // Generar enlace usando ISBN (si existe); de lo contrario, se usa el EAN.
    var codigoBusqueda = isbn_val ? isbn_val : ean_val;
    var enlace = codigoBusqueda ? baseUrl + encodeURIComponent(codigoBusqueda) + extraParams : "";
    
    // Generar enlace basado en el Titulo
    var enlaceTitulo = titulo_val ? baseUrl + encodeURIComponent(titulo_val) + extraParams : "";
    
    // Recuperar los valores manuales de Precio y Portes, si existen para este ID
    var manualPrecio = "";
    var manualPortes = "";
    if (manualEntries.hasOwnProperty(id_val)) {
      manualPrecio = manualEntries[id_val].precio;
      manualPortes = manualEntries[id_val].portes;
    }
    
    newTCData.push([id_val, isbn_val, ean_val, titulo_val, posicion_val, manualPrecio, manualPortes, enlace, enlaceTitulo]);
  }
  
  Logger.log("Número total de registros que se escribirán en Todocoleccion: " + newTCData.length);
  
  // Limpiar la hoja Todocoleccion y escribir la nueva matriz actualizada
  sheetTC.clearContents();
  sheetTC.getRange(1, 1, newTCData.length, newTCData[0].length).setValues(newTCData);
  
  Logger.log("La tabla Todocoleccion se ha actualizado correctamente, preservando Precio y Portes.");
}
