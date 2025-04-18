/**
 * Script: Actualizador de ammareal (sólo libros)
 * Fecha: 12 abr 2025
 * Descripción:
 *  - Usa los datos de "Datos_ISBN" para actualizar la hoja "ammareal".
 *  - Se actualizan únicamente las columnas: ISBN13, Título y Posición (para el registro cuyo ID coincide),
 *    dejando intactos Precio, Vendido, Fecha y EnVenta.
 *  - Se ignoran registros de Datos_ISBN que no tengan ISBN (utilizado para ISBN13)
 *    o Título, de forma que solo se procesen los "libros".
 *  - Agrega al final aquellos registros nuevos (libros) que no existan.
 */
function actualizarAmmarealSoloLibros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaAmmareal = ss.getSheetByName("ammareal");
  var hojaISBN = ss.getSheetByName("Datos_ISBN");
  
  if (!hojaAmmareal || !hojaISBN) {
    throw new Error("Hojas 'ammareal' o 'Datos_ISBN' no encontradas.");
  }
  
  // Leer datos de ambas hojas (incluyendo cabecera en la primera fila)
  var datosAmmareal = hojaAmmareal.getDataRange().getValues();
  var datosISBN = hojaISBN.getDataRange().getValues();
  
  if (datosAmmareal.length < 1) {
    throw new Error("La hoja ammareal está vacía.");
  }
  if (datosISBN.length < 1) {
    throw new Error("La hoja Datos_ISBN está vacía.");
  }
  
  // Cabeceras
  var headerAmmareal = datosAmmareal[0];
  var headerISBN = datosISBN[0];
  
  // Validar que existan las columnas necesarias en ammareal
  var requiredAmmareal = ["ID", "ISBN13", "Título", "Posición", "Precio", "Vendido", "Fecha", "EnVenta"];
  requiredAmmareal.forEach(function(col) {
    if (headerAmmareal.indexOf(col) === -1) {
      throw new Error("Columna '" + col + "' no encontrada en ammareal.");
    }
  });
  
  // En Datos_ISBN se requieren: ID, BC, ISBN, EAN, Titulo y Posicion
  var requiredISBN = ["ID", "BC", "ISBN", "EAN", "Titulo", "Posicion"];
  requiredISBN.forEach(function(col) {
    if (headerISBN.indexOf(col) === -1) {
      throw new Error("Columna '" + col + "' no encontrada en Datos_ISBN.");
    }
  });
  
  // Mapear los índices de cada hoja
  var indicesAmmareal = {};
  headerAmmareal.forEach(function(col, i) {
    indicesAmmareal[col] = i;
  });
  var indicesISBN = {};
  headerISBN.forEach(function(col, i) {
    indicesISBN[col] = i;
  });
  
  // Construir un diccionario (map) de los registros existentes en ammareal por ID
  var destMap = {};
  for (var i = 1; i < datosAmmareal.length; i++) {
    var id = String(datosAmmareal[i][indicesAmmareal["ID"]]).trim();
    if (id) {
      destMap[id] = i; // Guarda el índice en el array
    }
  }
  
  // Contadores para logueo
  var updatedRows = 0;
  var newRows = 0;
  
  // Recorrer cada registro en Datos_ISBN (omitiendo la cabecera)
  for (var i = 1; i < datosISBN.length; i++) {
    var rowISBN = datosISBN[i];
    var idSource = String(rowISBN[indicesISBN["ID"]]).trim();
    if (!idSource) continue; // Si el ID está vacío, se omite
    
    // Extraer información
    // Se actualizará la columna ISBN13 de ammareal usando el valor de "ISBN" de Datos_ISBN
    var nuevoISBN13 = String(rowISBN[indicesISBN["ISBN"]] || "").trim();
    var nuevoTitulo = String(rowISBN[indicesISBN["Titulo"]] || "").trim();
    var nuevaPosicion = String(rowISBN[indicesISBN["Posicion"]] || "").trim();
    
    // Sólo se procesan registros (libros) que tengan ISBN y Título
    if (!nuevoISBN13 || !nuevoTitulo) {
      continue;
    }
    
    // Si el ID ya existe en ammareal, actualizar los campos
    if (destMap.hasOwnProperty(idSource)) {
      var rowIndex = destMap[idSource];
      var currentRow = datosAmmareal[rowIndex];
      currentRow[indicesAmmareal["ISBN13"]] = nuevoISBN13;
      currentRow[indicesAmmareal["Título"]] = nuevoTitulo;
      currentRow[indicesAmmareal["Posición"]] = nuevaPosicion;
      datosAmmareal[rowIndex] = currentRow;
      updatedRows++;
    } else {
      // Registro nuevo: se crea una fila con la cantidad de columnas de headerAmmareal
      var newRow = new Array(headerAmmareal.length).fill("");
      newRow[indicesAmmareal["ID"]] = idSource;
      newRow[indicesAmmareal["ISBN13"]] = nuevoISBN13;
      newRow[indicesAmmareal["Título"]] = nuevoTitulo;
      newRow[indicesAmmareal["Posición"]] = nuevaPosicion;
      // Los campos Precio, Vendido, Fecha y EnVenta quedan vacíos en el nuevo registro
      datosAmmareal.push(newRow);
      newRows++;
    }
  }
  
  // Reescribir la hoja "ammareal" con el arreglo actualizado
  hojaAmmareal.clearContents();
  hojaAmmareal.getRange(1, 1, datosAmmareal.length, datosAmmareal[0].length).setValues(datosAmmareal);
  
  Logger.log("Actualización de ammareal completada. Libros actualizados: " + updatedRows + ", libros nuevos añadidos: " + newRows);
}
