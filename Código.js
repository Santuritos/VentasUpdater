function guardarChatManual() {
    // Reemplaza este texto con el chat copiado
    const textoChat = `
    Aquí va el texto completo del chat copiado...
    Puedes pegar el contenido aquí.
    `;
    
    // ID de la carpeta en Google Drive donde se guardarán los archivos
    const folderId = '1TRjMBuHS0PrRirSf4iuCxXG4yjO84jvH'; // Reemplaza con el ID de tu carpeta
    const folder = DriveApp.getFolderById(folderId);

    // Nombre del archivo con fecha y hora actual
    const fileName = `Chat_${new Date().toISOString()}.txt`;

    // Crear el archivo en la carpeta
    folder.createFile(fileName, textoChat);
    Logger.log(`Archivo guardado: ${fileName}`);
}
