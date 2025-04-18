/**
 * Script: rellenarDatosISBNDesdeTexto
 * Descripción: Procesa la columna Texto en la hoja "Datos_ISBN", rellena las columnas vacías con los datos extraídos, 
 * marca los registros como procesados y limpia los valores de ISBN (BC, ISBN, EAN) eliminando guiones y espacios.
 * Versión: 1.6
 * Fecha de Creación: 10/04/2025
 * Última Modificación: 13/04/2025
 * Autor: Grok 3 (xAI)
 */
function rellenarDatosISBNDesdeTexto() {
  // Obtener la hoja activa
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const datosISBNSheet = spreadsheet.getSheetByName("Datos_ISBN");

  // Validar que la hoja exista
  if (!datosISBNSheet) {
    Logger.log("Hoja 'Datos_ISBN' no encontrada.");
    return;
  }

  // Obtener los datos
  const datosISBNData = datosISBNSheet.getDataRange().getValues();
  if (datosISBNData.length < 1) {
    Logger.log("No hay datos en la hoja Datos_ISBN.");
    return;
  }

  // Obtener los encabezados
  const datosISBNHeaders = datosISBNData[0];
  // Normalizar los encabezados (insensible a mayúsculas y tildes)
  const normalizedHeaders = datosISBNHeaders.map(header => {
    if (!header) return "";
    return header.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  });

  // Registrar los encabezados normalizados para depuración
  Logger.log("Encabezados originales: " + datosISBNHeaders.join(", "));
  Logger.log("Encabezados normalizados: " + normalizedHeaders.join(", "));

  // Índices de las columnas en "Datos_ISBN" (buscar sin tildes)
  const datosISBNIdIndex = normalizedHeaders.indexOf("id");
  const datosISBNBcIndex = normalizedHeaders.indexOf("bc");
  const datosISBNIsbnIndex = normalizedHeaders.indexOf("isbn");
  const datosISBNTituloIndex = normalizedHeaders.indexOf("titulo");
  const datosISBNAutorIndex = normalizedHeaders.indexOf("autor");
  const datosISBNEditorialIndex = normalizedHeaders.indexOf("editorial");
  const datosISBNAñoIndex = normalizedHeaders.indexOf("ano");
  const datosISBNEdicionIndex = normalizedHeaders.indexOf("edicion");
  const datosISBNGeneroIndex = normalizedHeaders.indexOf("genero");
  const datosISBNDescripcionIndex = normalizedHeaders.indexOf("descripcion");
  const datosISBNPrecioIndex = normalizedHeaders.indexOf("precio");
  const datosISBNTextoIndex = normalizedHeaders.indexOf("texto");
  const datosISBNProcesadoIndex = normalizedHeaders.indexOf("procesado");

  // Registrar los índices para depuración
  Logger.log("Índices de columnas: " +
    "ID=" + datosISBNIdIndex + ", " +
    "BC=" + datosISBNBcIndex + ", " +
    "ISBN=" + datosISBNIsbnIndex + ", " +
    "Título=" + datosISBNTituloIndex + ", " +
    "Autor=" + datosISBNAutorIndex + ", " +
    "Editorial=" + datosISBNEditorialIndex + ", " +
    "Año=" + datosISBNAñoIndex + ", " +
    "Edición=" + datosISBNEdicionIndex + ", " +
    "Genero=" + datosISBNGeneroIndex + ", " +
    "Descripcion=" + datosISBNDescripcionIndex + ", " +
    "Precio=" + datosISBNPrecioIndex + ", " +
    "Texto=" + datosISBNTextoIndex + ", " +
    "Procesado=" + datosISBNProcesadoIndex);

  // Validar que las columnas existan
  if (datosISBNIdIndex === -1) Logger.log("Columna 'ID' no encontrada en Datos_ISBN.");
  if (datosISBNBcIndex === -1) Logger.log("Columna 'BC' no encontrada en Datos_ISBN.");
  if (datosISBNIsbnIndex === -1) Logger.log("Columna 'ISBN' no encontrada en Datos_ISBN.");
  if (datosISBNTituloIndex === -1) Logger.log("Columna 'Título' o 'Titulo' no encontrada en Datos_ISBN.");
  if (datosISBNAutorIndex === -1) Logger.log("Columna 'Autor' no encontrada en Datos_ISBN.");
  if (datosISBNEditorialIndex === -1) Logger.log("Columna 'Editorial' no encontrada en Datos_ISBN.");
  if (datosISBNAñoIndex === -1) Logger.log("Columna 'Año' no encontrada en Datos_ISBN.");
  if (datosISBNEdicionIndex === -1) Logger.log("Columna 'Edición' no encontrada en Datos_ISBN.");
  if (datosISBNGeneroIndex === -1) Logger.log("Columna 'Genero' no encontrada en Datos_ISBN.");
  if (datosISBNDescripcionIndex === -1) Logger.log("Columna 'Descripcion' no encontrada en Datos_ISBN.");
  if (datosISBNPrecioIndex === -1) Logger.log("Columna 'Precio' no encontrada en Datos_ISBN.");
  if (datosISBNTextoIndex === -1) Logger.log("Columna 'Texto' no encontrada en Datos_ISBN.");
  if (datosISBNProcesadoIndex === -1) Logger.log("Columna 'Procesado' no encontrada en Datos_ISBN.");

  if (datosISBNIdIndex === -1 || datosISBNBcIndex === -1 || datosISBNIsbnIndex === -1 || datosISBNTituloIndex === -1 ||
      datosISBNAutorIndex === -1 || datosISBNEditorialIndex === -1 || datosISBNAñoIndex === -1 || datosISBNEdicionIndex === -1 ||
      datosISBNGeneroIndex === -1 || datosISBNDescripcionIndex === -1 || datosISBNPrecioIndex === -1 || datosISBNTextoIndex === -1 ||
      datosISBNProcesadoIndex === -1) {
    Logger.log("Faltan columnas necesarias en la hoja Datos_ISBN. Revisa los mensajes anteriores para más detalles.");
    return;
  }

  // Contador para depuración
  let updatedRows = 0;

  // Procesar cada fila
  for (let i = 1; i < datosISBNData.length; i++) {
    // Verificar si el registro ya está marcado como procesado
    let procesado = datosISBNData[i][datosISBNProcesadoIndex];
    if (procesado === true || procesado === "TRUE" || procesado === "true") {
      Logger.log(`Fila ${i + 1} ya está marcada como procesada. Saltando...`);
      continue;
    }

    let texto = datosISBNData[i][datosISBNTextoIndex] ? datosISBNData[i][datosISBNTextoIndex].toString() : "";
    if (!texto) {
      // Si no hay texto, marcar como procesado para evitar procesar esta fila en el futuro
      datosISBNSheet.getRange(i + 1, datosISBNProcesadoIndex + 1).setValue("TRUE");
      Logger.log(`Fila ${i + 1} no tiene texto. Marcada como procesada.`);
      continue;
    }

    // Dividir el texto en líneas
    let lineas = texto.split("\n");
    let datosExtraidos = {};

    // Procesar cada línea para extraer los datos
    for (let j = 0; j < lineas.length; j++) {
      let linea = lineas[j].trim();
      if (!linea) continue;

      // Extraer ISBN 13 (para la columna ISBN)
      if (linea.startsWith("ISBN 13:")) {
        datosExtraidos.isbn = linea.replace("ISBN 13:", "").trim();
      }
      // Extraer ISBN 10 (para la columna BC)
      else if (linea.startsWith("ISBN 10:")) {
        datosExtraidos.bc = linea.replace("ISBN 10:", "").trim();
      }
      // Extraer Título
      else if (linea.startsWith("Título:")) {
        datosExtraidos.titulo = linea.replace("Título:", "").trim();
      }
      // Extraer Autor/es
      else if (linea.startsWith("Autor/es:")) {
        datosExtraidos.autor = linea.replace("Autor/es:", "").trim().replace(" ; tr. [Ver títulos]", "");
      }
      // Extraer Fecha Edición (Año)
      else if (linea.startsWith("Fecha Edición:")) {
        let fecha = linea.replace("Fecha Edición:", "").trim();
        // Extraer el año (formato MM/YYYY)
        let partesFecha = fecha.split("/");
        datosExtraidos.año = partesFecha[1] ? partesFecha[1].trim() : "";
      }
      // Extraer Publicación (Editorial)
      else if (linea.startsWith("Publicación:")) {
        datosExtraidos.editorial = linea.replace("Publicación:", "").trim();
      }
      // Extraer Descripción
      else if (linea.startsWith("Descripción:")) {
        datosExtraidos.descripcion = linea.replace("Descripción:", "").trim();
      }
      // Extraer Materia/s (Género)
      else if (linea.startsWith("Materia/s:")) {
        datosExtraidos.genero = linea.replace("Materia/s:", "").trim();
      }
      // Extraer Precio
      else if (linea.startsWith("Precio:")) {
        datosExtraidos.precio = linea.replace("Precio:", "").trim().replace(" Euros", "");
      }
    }

    // Actualizar solo las columnas vacías
    let updated = false;

    // ISBN (ISBN 13)
    if (!datosISBNData[i][datosISBNIsbnIndex] && datosExtraidos.isbn) {
      datosISBNSheet.getRange(i + 1, datosISBNIsbnIndex + 1).setValue(datosExtraidos.isbn);
      updated = true;
    }

    // BC (ISBN 10)
    if (!datosISBNData[i][datosISBNBcIndex] && datosExtraidos.bc) {
      datosISBNSheet.getRange(i + 1, datosISBNBcIndex + 1).setValue(datosExtraidos.bc);
      updated = true;
    }

    // Título
    if (!datosISBNData[i][datosISBNTituloIndex] && datosExtraidos.titulo) {
      datosISBNSheet.getRange(i + 1, datosISBNTituloIndex + 1).setValue(datosExtraidos.titulo);
      updated = true;
    }

    // Autor
    if (!datosISBNData[i][datosISBNAutorIndex] && datosExtraidos.autor) {
      datosISBNSheet.getRange(i + 1, datosISBNAutorIndex + 1).setValue(datosExtraidos.autor);
      updated = true;
    }

    // Editorial
    if (!datosISBNData[i][datosISBNEditorialIndex] && datosExtraidos.editorial) {
      datosISBNSheet.getRange(i + 1, datosISBNEditorialIndex + 1).setValue(datosExtraidos.editorial);
      updated = true;
    }

    // Año
    if (!datosISBNData[i][datosISBNAñoIndex] && datosExtraidos.año) {
      datosISBNSheet.getRange(i + 1, datosISBNAñoIndex + 1).setValue(datosExtraidos.año);
      updated = true;
    }

    // Descripción
    if (!datosISBNData[i][datosISBNDescripcionIndex] && datosExtraidos.descripcion) {
      datosISBNSheet.getRange(i + 1, datosISBNDescripcionIndex + 1).setValue(datosExtraidos.descripcion);
      updated = true;
    }

    // Género
    if (!datosISBNData[i][datosISBNGeneroIndex] && datosExtraidos.genero) {
      datosISBNSheet.getRange(i + 1, datosISBNGeneroIndex + 1).setValue(datosExtraidos.genero);
      updated = true;
    }

    // Precio
    if (!datosISBNData[i][datosISBNPrecioIndex] && datosExtraidos.precio) {
      datosISBNSheet.getRange(i + 1, datosISBNPrecioIndex + 1).setValue(datosExtraidos.precio);
      updated = true;
    }

    // Si se actualizó alguna columna, marcar como procesado
    if (updated) {
      datosISBNSheet.getRange(i + 1, datosISBNProcesadoIndex + 1).setValue("TRUE");
      updatedRows++;
      Logger.log(`Fila ${i + 1} actualizada y marcada como procesada.`);
    }
  }

  Logger.log(`Se actualizaron ${updatedRows} filas en Datos_ISBN con datos extraídos de la columna Texto.`);

  // Limpieza de ISBN (BC, ISBN, EAN) después de procesar
  Logger.log(`Iniciando limpieza de ISBN en ${datosISBNData.length - 1} filas.`);
  const lastRow = datosISBNSheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) { // Iterar desde la fila 2
    for (let col = datosISBNBcIndex + 1; col <= datosISBNIsbnIndex + 1; col++) { // Columnas BC, ISBN, EAN
      const cell = datosISBNSheet.getRange(i, col);
      let valor = cell.getValue();

      if (valor && typeof valor === "string") {
        const valorLimpio = valor.replace(/[-\s]/g, "");
        if (valor !== valorLimpio) {
          cell.setValue(valorLimpio);
          Logger.log(`Fila ${i}, Columna ${col}: "${valor}" → "${valorLimpio}"`);
        }
      } else {
        Logger.log(`Fila ${i}, Columna ${col}: Celda vacía o no es cadena, se omite.`);
      }
    }
  }

  Logger.log("Limpieza de ISBN completada.");
}