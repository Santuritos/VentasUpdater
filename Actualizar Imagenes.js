function actualizarImagenesISBN() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos_ISBN");
  var folderId = "1YIYituaLeeZDPRx61uO4hLfpQ1RrDmLC"; // ID de la carpeta en Google Drive
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  var imagenes = {};
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileUrl = file.getUrl();
    imagenes[fileName] = file;
  }
  
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  
  var columnas = {
    "Portada": ["Portada", "Portada_Link"],
    "Lomo": ["Lomo", "Lomo_Link"],
    "Trasera": ["Trasera", "Trasera_Link"],
    "Dentro": ["Dentro", "Dentro_Link"]
  };
  
  var idsVistos = {};
  
  for (var i = 1; i < datos.length; i++) {
    var idRegistro = datos[i][headers.indexOf("ISBN")];
    if (idsVistos[idRegistro]) {
      var nuevoId = idRegistro + "_" + (idsVistos[idRegistro] + 1);
      idsVistos[idRegistro]++;
    } else {
      var nuevoId = idRegistro;
      idsVistos[idRegistro] = 1;
    }
    
    for (var tipo in columnas) {
      var columnaImagen = headers.indexOf(columnas[tipo][1]);
      
      // Borrar el contenido actual antes de actualizar
      sheet.getRange(i + 1, columnaImagen + 1).setValue("");
      
      var nombreArchivo = datos[i][headers.indexOf(tipo)];
      if (nombreArchivo && imagenes[nombreArchivo]) {
        var nuevoNombre = nuevoId + "_" + tipo + ".jpg";
        imagenes[nombreArchivo].setName(nuevoNombre);
        sheet.getRange(i + 1, columnaImagen + 1).setValue(imagenes[nombreArchivo].getUrl());
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("Se han actualizado los registros duplicados y renombrado las imágenes correctamente.");
}
