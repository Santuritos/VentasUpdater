function limpiarISBN() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos_ISBN");
  const lastRow = sheet.getLastRow(); // Última fila con datos

  Logger.log(`Iniciando limpieza de ISBN en ${lastRow - 1} filas.`);

  for (let i = 2; i <= lastRow; i++) { // Iterar desde la fila 2
    for (let col = 2; col <= 4; col++) { // Columnas B (2), C (3), y D (4)
      const cell = sheet.getRange(i, col); // Obtener la celda específica
      let valor = cell.getValue(); // Obtener el valor de la celda

      if (valor && typeof valor === "string") { // Si el valor no está vacío y es cadena
        const valorLimpio = valor.replace(/[-\s]/g, ""); // Quitar guiones y espacios
        if (valor !== valorLimpio) {
          cell.setValue(valorLimpio); // Escribir el valor limpio
          Logger.log(`Fila ${i}, Columna ${col}: "${valor}" → "${valorLimpio}"`);
        }
      } else {
        Logger.log(`Fila ${i}, Columna ${col}: Celda vacía o no es cadena, se omite.`);
      }
    }
  }

  Logger.log("Limpieza de ISBN completada.");
}
