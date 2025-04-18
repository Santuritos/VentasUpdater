/**
 * Script: Actualizador de Momox (sin borrar)
 * Fecha: 12 abr 2025
 * Descripción:
 *  - Actualiza la hoja "Momox" usando los datos de "Datos_ISBN".
 *  - Basándose en el ID, actualiza únicamente: ISBN10, ISBN13, EAN, Título, Posición y Enlace.
 *    La columna Precio (y las demás columnas que no se indiquen) se mantiene sin cambios.
 *  - Agrega las filas que no existan en Momox.
 */
function actualizarMomoxSinBorrar() {
  // Obtener el archivo y las hojas involucradas.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMomox = ss.getSheetByName("Momox");
  var hojaISBN = ss.getSheetByName("Datos_ISBN");
  
  if (!hojaMomox || !hojaISBN) {
    throw new Error("Hojas 'Momox' o 'Datos_ISBN' no encontradas.");
  }
  
  // Leer datos de ambas hojas (incluyendo cabeceras en la primera fila)
  var datosMomox = hojaMomox.getDataRange().getValues();
  var datosISBN = hojaISBN.getDataRange().getValues();
  
  if (datosMomox.length < 1) {
    throw new Error("La hoja Momox está vacía.");
  }
  if (datosISBN.length < 1) {
    throw new Error("La hoja Datos_ISBN está vacía.");
  }
  
  // Obtener las cabeceras de cada hoja
  var headerMomox = datosMomox[0];
  var headerISBN = datosISBN[0];
  
  // Validar que existan las columnas requeridas en cada hoja.
  // En Momox se requieren: ID, ISBN10, ISBN13, EAN, Título, Posición, Enlace y Precio.
  var requiredMomox = ["ID", "ISBN10", "ISBN13", "EAN", "Título", "Posición", "Enlace", "Precio"];
  requiredMomox.forEach(function(col) {
    if (headerMomox.indexOf(col) === -1) {
      throw new Error("Columna '" + col + "' no encontrada en Momox.");
    }
  });
  // En Datos_ISBN se requieren: ID, BC (para ISBN10), ISBN (para ISBN13), EAN, Titulo y Posicion.
  var requiredISBN = ["ID", "BC", "ISBN", "EAN", "Titulo", "Posicion"];
  requiredISBN.forEach(function(col) {
    if (headerISBN.indexOf(col) === -1) {
      throw new Error("Columna '" + col + "' no encontrada en Datos_ISBN.");
    }
  });
  
  // Crear mappings de índices para cada hoja.
  var indicesMomox = {};
  headerMomox.forEach(function(col, i) {
    indicesMomox[col] = i;
  });
  var indicesISBN = {};
  headerISBN.forEach(function(col, i) {
    indicesISBN[col] = i;
  });
  
  // Construir un diccionario de los registros existentes en Momox, usando el ID como clave.
  // Este diccionario indica en qué fila (del array) se encuentra cada registro.
  var destMap = {};
  for (var i = 1; i < datosMomox.length; i++) {
    var id = String(datosMomox[i][indicesMomox["ID"]]).trim();
    if (id) {
      destMap[id] = i;  // se guarda el índice de la fila
    }
  }
  
  // Contadores para loguear resultados
  var updatedRows = 0;
  var newRows = 0;
  
  // Recorrer cada registro de la hoja Datos_ISBN (omitimos la cabecera)
  for (var i = 1; i < datosISBN.length; i++) {
    var rowISBN = datosISBN[i];
    var idSource = String(rowISBN[indicesISBN["ID"]]).trim();
    if (!idSource) continue; // se omite si el ID está vacío
    
    // Extraer datos del registro fuente
    var nuevoISBN10 = String(rowISBN[indicesISBN["BC"]] || "").trim();
    var nuevoISBN13 = String(rowISBN[indicesISBN["ISBN"]] || "").trim();
    var nuevoEAN = String(rowISBN[indicesISBN["EAN"]] || "").trim();
    var nuevoTitulo = String(rowISBN[indicesISBN["Titulo"]] || "").trim();
    var nuevaPosicion = String(rowISBN[indicesISBN["Posicion"]] || "").trim();
    
    // Generar el enlace usando el primer valor no vacío entre ISBN13, ISBN10 y EAN.
    // En este ejemplo se asume que el enlace de Momox sigue la URL "https://momox.com/escaner/producto/"
    var valorParaEnlace = nuevoISBN13 || nuevoISBN10 || nuevoEAN;
    var nuevoEnlace = valorParaEnlace ? "https://www.momox.es/offer/" + valorParaEnlace : "";
    
    // Si el ID existe en Momox se actualizan los campos; de lo contrario se agrega una nueva fila.
    if (destMap.hasOwnProperty(idSource)) {
      // Actualización de registro existente: se conserva el Precio y demás columnas que no se indiquen.
      var rowIndex = destMap[idSource];
      var currentRow = datosMomox[rowIndex];
      
      currentRow[indicesMomox["ISBN10"]] = nuevoISBN10;
      currentRow[indicesMomox["ISBN13"]] = nuevoISBN13;
      currentRow[indicesMomox["EAN"]]    = nuevoEAN;
      currentRow[indicesMomox["Título"]] = nuevoTitulo;
      currentRow[indicesMomox["Posición"]] = nuevaPosicion;
      currentRow[indicesMomox["Enlace"]] = nuevoEnlace;
      
      datosMomox[rowIndex] = currentRow;
      updatedRows++;
    } else {
      // Nuevo registro: se crea una fila nueva con la misma cantidad de columnas que la cabecera.
      var newRow = new Array(headerMomox.length).fill("");
      newRow[indicesMomox["ID"]] = idSource;
      newRow[indicesMomox["ISBN10"]] = nuevoISBN10;
      newRow[indicesMomox["ISBN13"]] = nuevoISBN13;
      newRow[indicesMomox["EAN"]]    = nuevoEAN;
      newRow[indicesMomox["Título"]] = nuevoTitulo;
      newRow[indicesMomox["Posición"]] = nuevaPosicion;
      newRow[indicesMomox["Enlace"]] = nuevoEnlace;
      // Las demás columnas (como Precio, Vendido, Fecha, EnVenta, etc.) se dejan en blanco.
      
      datosMomox.push(newRow);
      newRows++;
    }
  }
  
  // Actualizar la hoja "Momox" reescribiendo todo el rango.
  hojaMomox.clearContents();
  hojaMomox.getRange(1, 1, datosMomox.length, datosMomox[0].length).setValues(datosMomox);
  
  Logger.log("Actualización completada. Filas actualizadas: " + updatedRows + ", filas nuevas añadidas: " + newRows);
}
