/**
 * Script: Eliminar Duplicados por ISBN/EAN - Versión 1.3
 * Descripción: Elimina duplicados en la hoja activa basándose en ISBN (columna C) o EAN (columna D).
 * Registra los duplicados eliminados en el log y devuelve un mensaje para AppSheet.
 * Fecha: 11 abr 2025
 */
function eliminarDuplicadosISBN_EAN_Sencillo() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hoja.getDataRange().getValues();

  if (datos.length <= 1) {
    Logger.log("No hay suficientes datos para procesar.");
    return "No hay suficientes datos para procesar.";
  }

  const encabezados = datos.shift(); // Extrae la cabecera
  const columnaISBN = 3; // Columna ISBN (C)
  const columnaEAN = 4; // Columna EAN (D)
  const columnaID = 1; // Columna ID (A) - Para el log
  const columnaTitulo = 6; // Columna Título (F) - Para el log (ajusta según tu hoja)

  const clavesVistas = new Set();
  const registrosUnicos = [];
  const duplicadosEliminados = [];

  // Filtrar registros únicos
  datos.forEach((fila, index) => {
    // Verificar que la fila tenga suficientes columnas
    if (fila.length < Math.max(columnaISBN, columnaEAN, columnaID, columnaTitulo)) {
      Logger.log(`Fila ${index + 2} incompleta, se omite: ${fila}`);
      return;
    }

    // Obtener ISBN y EAN, manejando valores vacíos o undefined
    const isbn = fila[columnaISBN - 1] ? fila[columnaISBN - 1].toString().trim() : "";
    const ean = fila[columnaEAN - 1] ? fila[columnaEAN - 1].toString().trim() : "";
    const clave = isbn || ean; // Prioriza ISBN, si no existe toma EAN

    // Obtener ID y Título para el log
    const id = fila[columnaID - 1] ? fila[columnaID - 1].toString().trim() : "Sin ID";
    const titulo = fila[columnaTitulo - 1] ? fila[columnaTitulo - 1].toString().trim() : "Sin Título";

    if (!clave) {
      Logger.log(`Fila ${index + 2} omitida: ISBN y EAN están vacíos. ID: ${id}, Título: ${titulo}`);
      return; // Si no hay ISBN ni EAN, omitir la fila
    }

    if (clavesVistas.has(clave)) {
      // Registrar el duplicado eliminado
      duplicadosEliminados.push(`Fila ${index + 2}: ID=${id}, ISBN/EAN=${clave}, Título=${titulo}`);
      return; // Si la clave ya existe, salta esta fila
    }

    clavesVistas.add(clave); // Marca la clave como vista
    registrosUnicos.push(fila); // Añade la fila a los registros únicos
  });

  // Si no hay registros únicos, detener
  if (registrosUnicos.length === 0) {
    Logger.log("No hay registros únicos para guardar.");
    return "No hay registros únicos para guardar.";
  }

  // Limpiar la hoja y sobrescribir solo los registros únicos
  hoja.clear();
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]); // Escribe la cabecera
  hoja.getRange(2, 1, registrosUnicos.length, encabezados.length).setValues(registrosUnicos); // Escribe los registros únicos

  // Registrar los duplicados eliminados en el log
  if (duplicadosEliminados.length > 0) {
    Logger.log("Duplicados eliminados:");
    duplicadosEliminados.forEach(duplicado => Logger.log(duplicado));
    const mensaje = `Eliminación completada: ${duplicadosEliminados.length} duplicados eliminados.`;
    Logger.log(mensaje);
    return mensaje;
  } else {
    Logger.log("No se encontraron duplicados para eliminar.");
    return "No se encontraron duplicados para eliminar.";
  }
}