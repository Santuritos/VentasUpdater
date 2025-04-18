function procesarLibrosConOpenLibrary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const id = sheet.getRange(i, 1).getValue().toString().trim(); // Columna A
    if (!id) continue; // Saltar si el ID está vacío

    const tipo = sheet.getRange(i, 13).getValue().toString().trim(); // Columna N (Tipo)
    if (tipo.toLowerCase() !== "libro") continue; // Solo procesa libros

    const isbnCell = sheet.getRange(i, 3); // Columna C (ISBN-13)
    const tituloCell = sheet.getRange(i, 7); // Columna G (Título)
    const autorCell = sheet.getRange(i, 8); // Columna H (Autor)
    const anioCell = sheet.getRange(i, 6); // Columna F (Año)
    const edicionCell = sheet.getRange(i, 9); // Columna I (Edición)
    const editorialCell = sheet.getRange(i, 10); // Columna J (Editorial)
    const descripcionCell = sheet.getRange(i, 11); // Columna K (Descripción)
    const generoCell = sheet.getRange(i, 12); // Columna L (Género)
    const procesadoCell = sheet.getRange(i, 14); // Columna N
    const urlCell = sheet.getRange(i, 31); // Columna AE (URL del libro)

    const isbn = isbnCell.getValue().toString().trim();
    if (!isbn) {
      procesadoCell.setValue("ISBN faltante").setBackground("#FFCCCC");
      continue;
    }

    const datosLibro = consultarOpenLibrary(isbn);
    if (datosLibro) {
      if (datosLibro.title) tituloCell.setValue(datosLibro.title).setBackground("#CCFFCC");
      if (datosLibro.authors) autorCell.setValue(datosLibro.authors).setBackground("#CCFFCC");
      if (datosLibro.publishedDate) anioCell.setValue(datosLibro.publishedDate.match(/\d{4}/)?.[0]).setBackground("#CCFFCC");
      if (datosLibro.edition) edicionCell.setValue(datosLibro.edition).setBackground("#CCFFCC");
      if (datosLibro.publisher) editorialCell.setValue(datosLibro.publisher).setBackground("#CCFFCC");
      if (datosLibro.description) descripcionCell.setValue(datosLibro.description).setBackground("#CCFFCC");
      if (datosLibro.series) generoCell.setValue(datosLibro.series).setBackground("#CCFFCC");
      if (datosLibro.url) urlCell.setValue(datosLibro.url).setBackground("#CCFFCC");
      procesadoCell.setValue("Procesado correctamente").setBackground("#CCFFCC");
    } else {
      procesadoCell.setValue("No encontrado").setBackground("#FFFF99");
    }

    // Limpiar celdas que contengan "no encontrad"
    [tituloCell, autorCell, anioCell, edicionCell, editorialCell, descripcionCell, generoCell].forEach(cell => {
      if (cell.getValue().toString().toLowerCase().includes("no encontrad")) {
        cell.setValue("").setBackground(null);
      }
    });
  }
}

function consultarOpenLibrary(isbn) {
  try {
    const url = `https://openlibrary.org/api/books?bibkeys=ISBN:${isbn}&format=json&jscmd=data`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const key = `ISBN:${isbn}`;

    if (json[key]) {
      const libro = json[key];
      const format = libro.format ? libro.format.match(/- (\d+\w ed\.)/) : null;
      return {
        title: libro.title,
        authors: libro.authors ? libro.authors.map(author => author.name).join(", ") : null,
        publishedDate: libro.publish_date,
        edition: format ? format[1] : null,
        publisher: libro.publishers ? libro.publishers.map(pub => pub.name).join(", ") : null,
        description: libro.notes || "",
        series: libro.series ? libro.series.join(", ") : null,
        url: `https://openlibrary.org${libro.key}`,
      };
    }
    return null;
  } catch (e) {
    Logger.log(`Error consultando Open Library: ${e}`);
    return null;
  }
}
