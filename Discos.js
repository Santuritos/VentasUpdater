/**
 * Script: Procesar No Libros (Discogs/TMDB)
 * Versión: 1.1
 * Fecha: 21 mar 2025
 *
 * Función:
 *   Recorre la hoja activa y para cada registro cuyo Tipo (columna M) no sea "Libro"
 *   se asegura de que exista un código (EAN). Si falta pero hay ISBN, lo mueve a EAN.
 *   Luego consulta la API de Discogs (para "cd", "lp", "vinyl") o TMDB (para "dvd", "movie", "film", "cd+dvd")
 *   y actualiza (solamente) las columnas: Año (F), Título (G), Autor (H) y Género (L).
 *   Se marca en la columna Procesado (N) con fondo #ffccee si se actualiza correctamente,
 *   o se marca "Sin código" o "No encontrado" con fondo #FFFF99.
 */

function procesarNoLibros() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const discogsApiKey = "nyRSmqWozkBOjOpVcLVqExPHvBtzzdJbcINKXNIZ";
  const tmdbApiKey = "7312d8fb0b4e49c3b0534e6466e47f20";
  
  // Se procesan filas de la 2 a la última
  Logger.log("Iniciando procesamiento en " + (lastRow - 1) + " filas.");
  let processedCount = 0;
  
  for (let i = 2; i <= lastRow; i++) {
    const id = sheet.getRange(i, 1).getValue().toString().trim(); // Columna A: ID
    if (!id) continue; // Si el ID está vacío, saltar la fila.
    
    const tipo = sheet.getRange(i, 13).getValue().toString().trim(); // Columna M: Tipo
    Logger.log("Fila " + i + ": Tipo = '" + tipo + "'");
    // Solo procesar si el tipo no es "libro" (ignorando mayúsculas)
    if (tipo.toLowerCase() === "libro") {
      Logger.log("Fila " + i + ": Tipo 'Libro' no se procesa.");
      continue;
    }
    
    // Obtener código de barras
    let ean = sheet.getRange(i, 4).getValue().toString().trim(); // Columna D: EAN
    const isbn = sheet.getRange(i, 3).getValue().toString().trim(); // Columna C: ISBN

    if (!ean && isbn) {
      // Mover ISBN a EAN si el campo EAN está vacío
      ean = isbn;
      sheet.getRange(i, 4).setValue(isbn); // Escribir en columna D
      sheet.getRange(i, 3).setValue("");    // Limpiar columna C
      Logger.log("Fila " + i + ": ISBN movido a EAN.");
    }
    
    if (!ean) {
      sheet.getRange(i, 14).setValue("Sin código").setBackground("#FFFF99"); // Columna N
      Logger.log("Fila " + i + ": No hay código (EAN), marcado como 'Sin código'.");
      continue;
    }
    
    let datos = null;
    // Si es música: cd, lp, vinyl
    if (["cd", "lp", "vinyl"].includes(tipo.toLowerCase())) {
      datos = consultarDiscogsAPI(ean, discogsApiKey);
    } 
    // Si es película o CD+DVD (o movie, film, dvd)
    else if (["dvd", "movie", "film", "cd+dvd"].includes(tipo.toLowerCase())) {
      const titulo = sheet.getRange(i, 7).getValue().toString().trim(); // Columna G: Título
      datos = consultarTMDBAPI(titulo, tmdbApiKey);
    }
    
    if (datos) {
      // Actualizar campos: 
      // Año (col F), Título (col G), Autor (col H) y Género (col L)
      sheet.getRange(i, 6).setValue(datos.publishedDate || "").setBackground("#ffccee");
      sheet.getRange(i, 7).setValue(datos.title || "Título no encontrado").setBackground("#ffccee");
      sheet.getRange(i, 8).setValue(datos.artist || "").setBackground("#ffccee");
      sheet.getRange(i, 12).setValue(datos.genre || "").setBackground("#ffccee");
      sheet.getRange(i, 14).setValue("Procesado correctamente").setBackground("#ffccee");
      processedCount++;
      Logger.log("Fila " + i + " actualizada correctamente.");
    } else {
      sheet.getRange(i, 14).setValue("No encontrado").setBackground("#FFFF99");
      Logger.log("Fila " + i + ": No se obtuvieron datos de la API.");
    }
  }
  
  Logger.log("Procesamiento completado. Filas procesadas: " + processedCount + ".");
}


// Consultar la API de Discogs
function consultarDiscogsAPI(ean, apiKey) {
  try {
    const url = `https://api.discogs.com/database/search?barcode=${ean}&token=${apiKey}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    
    if (json.results && json.results.length > 0) {
      const disco = json.results[0];
      return {
        title: disco.title || null,
        artist: disco.label ? disco.label.join(", ") : null,
        publishedDate: disco.year || null,
        genre: disco.style ? disco.style.join(", ") : null
      };
    }
    return null;
  } catch (e) {
    Logger.log(`Error consultando Discogs API para EAN ${ean}: ${e}`);
    return null;
  }
}

// Consultar la API de TMDB
function consultarTMDBAPI(titulo, apiKey) {
  try {
    const url = `https://api.themoviedb.org/3/search/movie?query=${encodeURIComponent(titulo)}&api_key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    
    if (json.results && json.results.length > 0) {
      const pelicula = json.results[0];
      return {
        title: pelicula.title || null,
        publishedDate: pelicula.release_date || null,
        genre: pelicula.genre_ids ? pelicula.genre_ids.join(", ") : null
      };
    }
    return null;
  } catch (e) {
    Logger.log(`Error consultando TMDB API para título '${titulo}': ${e}`);
    return null;
  }
}
