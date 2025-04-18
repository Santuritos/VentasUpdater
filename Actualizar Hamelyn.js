/**
 * Script: Actualizador de Hamelyn (sin borrar)
 * Fecha: 12 abr 2025
 * Descripción:
 *  - Actualiza la hoja "Hamelyn" usando los datos de "Datos_ISBN".
 *  - Basándose en el ID, actualiza únicamente: ISBN10, ISBN13, EAN, Título, Posición y Enlace.
 *    La columna Precio (y las demás no incluidas) se mantiene sin cambios.
 *  - Agrega las filas que no existan en Hamelyn.
 */
function actualizarHamelynSinBorrar() {
  // Obtener el archivo y las hojas involucradas.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaHamelyn = ss.getSheetByName("Hamelyn");
  var hojaISBN = ss.getSheetByName("Datos_ISBN");
  
  if (!hojaHamelyn || !hojaISBN) {
    throw new Error("Hojas 'Hamelyn' o 'Datos_ISBN' no encontradas.");
  }
  
  // Leer datos de ambas hojas (incluyendo cabeceras en la primera fila)
  var datosHamelyn = hojaHamelyn.getDataRange().getValues();
  var datosISBN = hojaISBN.getDataRange().getValues();
  
  if (datosHamelyn.length < 1) {
    throw new Error("La hoja Hamelyn está vacía.");
  }
  if (datosISBN.length < 1) {
    throw new Error("La hoja Datos_ISBN está vacía.");
  }
  
  // Obtener las cabeceras de cada hoja
  var headerHamelyn = datosHamelyn[0];
  var headerISBN = datosISBN[0];
  
  // Validar que existan las columnas requeridas en cada hoja.
  // En Hamelyn se requieren al menos: ID, ISBN10, ISBN13, EAN, Título, Posición, Enlace y Precio.
  var requiredHamelyn = ["ID", "ISBN10", "ISBN13", "EAN", "Título", "Posición", "Enlace", "Precio"];
  requiredHamelyn.forEach(function(col) {
    if (headerHamelyn.indexOf(col) === -1) {
      throw new Error("Columna '" + col + "' no encontrada en Hamelyn.");
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
  var indicesHamelyn = {};
  headerHamelyn.forEach(function(col, i) {
    indicesHamelyn[col] = i;
  });
  var indicesISBN = {};
  headerISBN.forEach(function(col, i) {
    indicesISBN[col] = i;
  });
  
  // Construir un diccionario de los registros existentes en Hamelyn, usando el ID como clave.
  // Este diccionario indicará en qué fila se encuentra cada registro.
  var destMap = {};
  for (var i = 1; i < datosHamelyn.length; i++) {
    var id = String(datosHamelyn[i][indicesHamelyn["ID"]]).trim();
    if (id) {
      destMap[id] = i;  // guarda el índice de la fila en el arreglo
    }
  }
  
  // Contadores para loguear resultados
  var updatedRows = 0;
  var newRows = 0;
  
  // Recorrer cada registro de la hoja Datos_ISBN (omitimos la cabecera)
  for (var i = 1; i < datosISBN.length; i++) {
    var rowISBN = datosISBN[i];
    var idSource = String(rowISBN[indicesISBN["ID"]]).trim();
    if (!idSource) continue; // Si el ID está vacío, se omite
    
    // Extraer datos del registro fuente
    var nuevoISBN10 = String(rowISBN[indicesISBN["BC"]] || "").trim();
    var nuevoISBN13 = String(rowISBN[indicesISBN["ISBN"]] || "").trim();
    var nuevoEAN = String(rowISBN[indicesISBN["EAN"]] || "").trim();
    var nuevoTitulo = String(rowISBN[indicesISBN["Titulo"]] || "").trim();
    var nuevaPosicion = String(rowISBN[indicesISBN["Posicion"]] || "").trim();
    
    // Generar el enlace usando el primer valor no vacío entre ISBN13, ISBN10 y EAN
    var valorParaEnlace = nuevoISBN13 || nuevoISBN10 || nuevoEAN;
    var nuevoEnlace = valorParaEnlace ? "https://hamelyn.com/escaner/producto/" + valorParaEnlace : "";
    
    // Si el ID existe en Hamelyn, se actualizan los campos; de lo contrario se agrega una nueva fila.
    if (destMap.hasOwnProperty(idSource)) {
      // Actualización de registro existente: conservar Precio y demás columnas no indicadas.
      var rowIndex = destMap[idSource];
      var currentRow = datosHamelyn[rowIndex];
      
      // Actualizamos solo las siguientes columnas:
      currentRow[indicesHamelyn["ISBN10"]] = nuevoISBN10;
      currentRow[indicesHamelyn["ISBN13"]] = nuevoISBN13;
      currentRow[indicesHamelyn["EAN"]]    = nuevoEAN;
      currentRow[indicesHamelyn["Título"]] = nuevoTitulo;
      currentRow[indicesHamelyn["Posición"]] = nuevaPosicion;
      currentRow[indicesHamelyn["Enlace"]] = nuevoEnlace;
      // Se deja intacto el valor de "Precio" y las demás columnas.
      
      datosHamelyn[rowIndex] = currentRow;
      updatedRows++;
    } else {
      // Nuevo registro: se crea una nueva fila con la misma cantidad de columnas que la cabecera.
      var newRow = new Array(headerHamelyn.length).fill("");
      // Se asigna la información de los campos actualizables.
      newRow[indicesHamelyn["ID"]] = idSource;
      newRow[indicesHamelyn["ISBN10"]] = nuevoISBN10;
      newRow[indicesHamelyn["ISBN13"]] = nuevoISBN13;
      newRow[indicesHamelyn["EAN"]]    = nuevoEAN;
      newRow[indicesHamelyn["Título"]] = nuevoTitulo;
      newRow[indicesHamelyn["Posición"]] = nuevaPosicion;
      newRow[indicesHamelyn["Enlace"]] = nuevoEnlace;
      // Las demás columnas (como Precio, Vendido, Fecha, EnVenta, etc.) se dejan en blanco.
      
      datosHamelyn.push(newRow);
      newRows++;
    }
  }
  
  // Actualizar la hoja "Hamelyn" sin borrar registros (se reescribe todo el rango)
  hojaHamelyn.clearContents();
  hojaHamelyn.getRange(1, 1, datosHamelyn.length, datosHamelyn[0].length).setValues(datosHamelyn);
  
  Logger.log("Actualización completada. Filas actualizadas: " + updatedRows + ", filas nuevas añadidas: " + newRows);
}
