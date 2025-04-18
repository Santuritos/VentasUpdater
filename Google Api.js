function procesarLibrosConGoogleBooksPorISBN() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const apiKey = "AIzaSyDzmBfAnn3KXe6xGdTyvrn5un2bMxcZSXc"; // Clave API válida de Google Books

  for (let i = 2; i <= lastRow; i++) {
    const id = sheet.getRange(i, 1).getValue().toString().trim(); // Columna A
    if (!id) continue; // Saltar si el ID está vacío

    const tipo = sheet.getRange(i, 13).getValue().toString().trim(); // Columna N (Tipo)
    if (tipo.toLowerCase() !== "libro") continue; // Solo procesa libros

    const isbn10 = sheet.getRange(i, 2).getValue().toString().trim(); // Columna B (ISBN-10)
    const isbn13 = sheet.getRange(i, 3).getValue().toString().trim(); // Columna C (ISBN-13)
    const descripcionCell = sheet.getRange(i, 11); // Columna K (Descripción)
    const generoCell = sheet.getRange(i, 12); // Columna L (Género)
    const procesadoCell = sheet.getRange(i, 14); // Columna N

    const isbn = isbn13 || isbn10; // Priorizar ISBN-13 si está disponible
    if (!isbn) {
      procesadoCell.setBackground("#FF9999"); // Destacar celda en rojo claro
      continue;
    }

    try {
      const datosLibro = consultarGoogleBooksPorISBN(isbn, apiKey);
      if (datosLibro) {
        if (datosLibro.description) descripcionCell.setValue(datosLibro.description).setBackground("#CCFFCC");
        if (datosLibro.categories) generoCell.setValue(datosLibro.categories).setBackground("#CCFFCC");
        procesadoCell.setBackground("#CCFFCC"); // Indicar que se procesó correctamente
      } else {
        procesadoCell.setBackground("#FF9999"); // Indicar "No encontrado" con rojo claro
      }
    } catch (e) {
      procesadoCell.setBackground("#FF9999"); // Destacar celda en rojo claro para errores
      Logger.log(`Error procesando fila ${i}: ${e.message}`);
    }

    // Limpiar celdas que contengan "no encontrad"
    [descripcionCell, generoCell].forEach(cell => {
      if (cell.getValue().toString().toLowerCase().includes("no encontrad")) {
        cell.setValue("").setBackground(null);
      }
    });
  }
}

function consultarGoogleBooksPorISBN(isbn, apiKey) {
  try {
    const url = `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(isbn)}&key=${apiKey}&maxResults=1`;
    Logger.log(`Consultando URL: ${url}`); // Registrar la URL consultada
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());

    if (json.items && json.items.length > 0) {
      const libro = json.items[0].volumeInfo;
      return {
        description: libro.description || null,
        categories: libro.categories ? libro.categories.join(", ") : null,
      };
    }
    Logger.log("No se encontraron datos para el ISBN.");
    return null;
  } catch (e) {
    Logger.log(`Error consultando Google Books API: ${e.message}`);
    throw new Error("Error en la consulta a Google Books");
  }
}
