function procesarFormulario() {
  // Obtener la hoja activa
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener las hojas
  var formularioSheet = spreadsheet.getSheetByName("Formulario");
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var enviosSheet = spreadsheet.getSheetByName("Envios");
  var pedidosSheet = spreadsheet.getSheetByName("Pedidos");

  // Añadir un pequeño retraso para asegurar que AppSheet sincronice los datos
  Utilities.sleep(2000); // Retraso de 2 segundos

  // Leer los datos de la tabla "Formulario"
  var formularioData = formularioSheet.getDataRange().getValues();
  var formularioHeaders = formularioData[0];
  var proveedorIndex = formularioHeaders.indexOf("Proveedor");

  // Obtener el último registro (el más reciente)
  var lastRow = formularioData.length - 1;
  if (lastRow < 1) return;
  var proveedor = formularioData[lastRow][proveedorIndex];

  // Validar que el campo Proveedor no esté vacío
  if (!proveedor) return;

  // Generar el Número de pedido en el formato PED-AÑO-MES-ALEATORIO
  var today = new Date();
  var year = today.getFullYear();
  var month = ("0" + (today.getMonth() + 1)).slice(-2);
  var day = ("0" + today.getDate()).slice(-2);
  var randomNum = Math.floor(1000 + Math.random() * 9000);
  var numeroPedido = "PED-" + year + "-" + month + "-" + randomNum;

  // Formatear la fecha en DD/MM/YYYY
  var formattedDate = day + "/" + month + "/" + year;

  // Actualizar el último registro en "Formulario" con el Número de pedido
  var numeroPedidoIndex = formularioHeaders.indexOf("Número de pedido");
  if (numeroPedidoIndex === -1) return;
  formularioSheet.getRange(lastRow + 1, numeroPedidoIndex + 1).setValue(numeroPedido);

  // Calcular el total de los registros enviados (suma de los precios)
  var ventasData = ventasSheet.getDataRange().getValues();
  var ventasHeaders = ventasData[0];
  var marcadoIndex = ventasHeaders.indexOf("Marcado");
  var ventasHamelynIndex = ventasHeaders.indexOf("Hamelyn");
  var ventasMomoxIndex = ventasHeaders.indexOf("Momox");
  var ventasAmmarealIndex = ventasHeaders.indexOf("Ammareal");
  var total = 0;

  for (var i = 1; i < ventasData.length; i++) {
    // Verificar si Marcado es "TRUE" (en mayúsculas)
    var marcadoValue = ventasData[i][marcadoIndex];
    var isMarcado = false;
    if (typeof marcadoValue === "string") {
      isMarcado = marcadoValue === "TRUE";
    } else if (typeof marcadoValue === "boolean") {
      isMarcado = marcadoValue === true;
      ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue(isMarcado ? "TRUE" : "FALSE");
    }

    if (isMarcado) {
      var precio = 0;
      if (proveedor === "Hamelyn") {
        precio = Number(ventasData[i][ventasHamelynIndex]) || 0;
      } else if (proveedor === "Momox") {
        precio = Number(ventasData[i][ventasMomoxIndex]) || 0;
      } else if (proveedor === "Ammareal") {
        precio = Number(ventasData[i][ventasAmmarealIndex]) || 0;
      }
      total += precio;
    }
  }

  // Añadir un registro a la tabla "Pedidos"
  pedidosSheet.appendRow([
    numeroPedido,
    numeroPedido,
    formattedDate,
    proveedor,
    total
  ]);

  // Paso 1: Copiar el Número de pedido y el Proveedor a los registros con Marcado = TRUE en "Ventas"
  var numeroPedidoTemporalIndex = ventasHeaders.indexOf("Numero de pedido Temporal");
  var proveedorVentasIndex = ventasHeaders.indexOf("Proveedor");

  for (var i = 1; i < ventasData.length; i++) {
    var marcadoValue = ventasData[i][marcadoIndex];
    var isMarcado = false;
    if (typeof marcadoValue === "string") {
      isMarcado = marcadoValue === "TRUE";
    } else if (typeof marcadoValue === "boolean") {
      isMarcado = marcadoValue === true;
      ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue(isMarcado ? "TRUE" : "FALSE");
    }

    if (isMarcado) {
      ventasSheet.getRange(i + 1, numeroPedidoTemporalIndex + 1).setValue(numeroPedido);
      ventasSheet.getRange(i + 1, proveedorVentasIndex + 1).setValue(proveedor);
    }
  }

  // Paso 2: Enviar los registros con Marcado = TRUE a "Envios"
  var ventasIdIndex = ventasHeaders.indexOf("ID");
  var ventasTituloIndex = ventasHeaders.indexOf("Titulo");

  // Índices fijos para ISBN y EAN
  var ventasIsbnIndex = 1; // Columna B en Ventas
  var ventasEanIndex = 2;  // Columna C en Ventas

  for (var i = 1; i < ventasData.length; i++) {
    var marcadoValue = ventasData[i][marcadoIndex];
    var isMarcado = false;
    if (typeof marcadoValue === "string") {
      isMarcado = marcadoValue === "TRUE";
    } else if (typeof marcadoValue === "boolean") {
      isMarcado = marcadoValue === true;
      ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue(isMarcado ? "TRUE" : "FALSE");
    }

    if (isMarcado) {
      var precio = 0;
      if (proveedor === "Hamelyn") {
        precio = Number(ventasData[i][ventasHamelynIndex]) || 0;
      } else if (proveedor === "Momox") {
        precio = Number(ventasData[i][ventasMomoxIndex]) || 0;
      } else if (proveedor === "Ammareal") {
        precio = Number(ventasData[i][ventasAmmarealIndex]) || 0;
      }

      // Copiar directamente ISBN y EAN de Ventas
      var isbnValue = (ventasData[i][ventasIsbnIndex] || "").toString().trim();
      var eanValue = (ventasData[i][ventasEanIndex] || "").toString().trim();

      // Añadir un nuevo registro en "Envios" escribiendo cada columna individualmente
      var lastEnviosRow = enviosSheet.getLastRow() + 1; // Nueva fila
      enviosSheet.getRange(lastEnviosRow, 1).setValue(numeroPedido + "-" + i); // Columna A: ID
      enviosSheet.getRange(lastEnviosRow, 2).setValue(numeroPedido); // Columna B: Número de pedido
      enviosSheet.getRange(lastEnviosRow, 3).setValue(formattedDate); // Columna C: Fecha
      enviosSheet.getRange(lastEnviosRow, 4).setValue(ventasData[i][ventasIdIndex]); // Columna D: ID Artículo
      enviosSheet.getRange(lastEnviosRow, 5).setValue(proveedor); // Columna E: Proveedor
      enviosSheet.getRange(lastEnviosRow, 6).setValue("Pendiente"); // Columna F: Estado
      enviosSheet.getRange(lastEnviosRow, 7).setValue(ventasData[i][ventasTituloIndex]); // Columna G: Título
      enviosSheet.getRange(lastEnviosRow, 8).setValue(precio); // Columna H: Precio
      enviosSheet.getRange(lastEnviosRow, 9).setValue(isbnValue); // Columna I: ISBN
      enviosSheet.getRange(lastEnviosRow, 10).setValue(eanValue); // Columna J: EAN
    }
  }

  // Paso 3: Actualizar los registros enviados en "Ventas" con Venta = "Enviados" y quitar la marca
  var ventaIndex = ventasHeaders.indexOf("Venta");
  for (var i = 1; i < ventasData.length; i++) {
    var marcadoValue = ventasData[i][marcadoIndex];
    var isMarcado = false;
    if (typeof marcadoValue === "string") {
      isMarcado = marcadoValue === "TRUE";
    } else if (typeof marcadoValue === "boolean") {
      isMarcado = marcadoValue === true;
      ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue(isMarcado ? "TRUE" : "FALSE");
    }

    if (isMarcado) {
      ventasSheet.getRange(i + 1, ventaIndex + 1).setValue("Enviados"); // Marcar como Enviados
      ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue("FALSE"); // Quitar la marca (Marcado = FALSE, en mayúsculas)
    }
  }

  // Añadir un retraso para asegurar que los datos se escriban antes de que AppSheet los lea
  Utilities.sleep(2000); // Retraso de 2 segundos
}

function doPost(e) {
  try {
    procesarFormulario();
    return ContentService.createTextOutput("Script ejecutado con éxito (POST)").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  try {
    procesarFormulario();
    return ContentService.createTextOutput("Script ejecutado con éxito (GET)").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error).setMimeType(ContentService.MimeType.TEXT);
  }
}