/**
 * Script: marcarEnVenta
 * Descripción: Marca los registros de un pedido en la hoja del proveedor correspondiente (Hamelyn, Momox, Ammareal)
 *              en la columna EnVenta (índice 11, columna K) con la fecha del pedido, usando datos de la hoja Envios.
 * Versión: 1.1
 * Fecha de Creación: 10/04/2025
 * Última Modificación: 10/04/2025
 * Autor: [Tu Nombre o "Grok 3 (xAI)"]
 */

function marcarEnVenta() {
  // Obtener la hoja activa
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener las hojas
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var enviosSheet = spreadsheet.getSheetByName("Envios");

  // Paso 1: Obtener el último pedido de la hoja "Envios"
  var enviosData = enviosSheet.getDataRange().getValues();
  if (enviosData.length < 2) {
    Logger.log("No hay datos en la hoja Envios (menos de 2 filas).");
    return;
  }

  var enviosHeaders = enviosData[0];
  // Mostrar todos los encabezados para depuración
  Logger.log("Encabezados en la hoja Envios: " + enviosHeaders.join(", "));

  // Buscar los índices de las columnas, ignorando mayúsculas/minúsculas
  var numeroPedidoIndex = -1;
  var fechaIndex = -1;
  var proveedorIndex = -1;

  for (var i = 0; i < enviosHeaders.length; i++) {
    var header = enviosHeaders[i].toString().toLowerCase().trim();
    if (header === "número de pedido" || header === "numero de pedido") {
      numeroPedidoIndex = i;
    } else if (header === "fecha" || header === "date" || header === "fecha pedido") {
      fechaIndex = i;
    } else if (header === "proveedor") {
      proveedorIndex = i;
    }
  }

  // Verificar si se encontraron todas las columnas
  if (numeroPedidoIndex === -1) Logger.log("Columna 'Número de pedido' no encontrada en Envios.");
  if (fechaIndex === -1) Logger.log("Columna 'Fecha' no encontrada en Envios.");
  if (proveedorIndex === -1) Logger.log("Columna 'Proveedor' no encontrada en Envios.");

  if (numeroPedidoIndex === -1 || fechaIndex === -1 || proveedorIndex === -1) {
    Logger.log("No se encontraron todas las columnas necesarias en la hoja Envios.");
    return;
  }

  // Obtener el último registro (el pedido más reciente)
  var lastRow = enviosData.length - 1;
  var numeroPedido = enviosData[lastRow][numeroPedidoIndex];
  var fechaPedido = enviosData[lastRow][fechaIndex];
  var proveedor = enviosData[lastRow][proveedorIndex];

  // Depuración: Mostrar los valores obtenidos
  Logger.log("Último envío encontrado - Número de pedido: " + numeroPedido + ", Fecha: " + fechaPedido + ", Proveedor: " + proveedor);

  // Validar que los datos del envío sean válidos
  if (!numeroPedido || !fechaPedido || !proveedor) {
    Logger.log("Datos del envío incompletos: Número de pedido: " + numeroPedido + ", Fecha: " + fechaPedido + ", Proveedor: " + proveedor);
    return;
  }

  // Paso 2: Obtener los registros de "Ventas" que corresponden a este pedido
  var ventasData = ventasSheet.getDataRange().getValues();
  if (ventasData.length < 2) {
    Logger.log("No hay datos en la hoja Ventas (menos de 2 filas).");
    return;
  }

  var ventasHeaders = ventasData[0];
  var numeroPedidoTemporalIndex = ventasHeaders.indexOf("Numero de pedido Temporal");
  var idIndex = ventasHeaders.indexOf("ID");

  if (numeroPedidoTemporalIndex === -1) {
    Logger.log("Columna 'Numero de pedido Temporal' no encontrada en la hoja Ventas.");
    return;
  }
  if (idIndex === -1) {
    Logger.log("Columna 'ID' no encontrada en la hoja Ventas.");
    return;
  }

  // Lista para almacenar los IDs de los registros enviados en este pedido
  var idsEnviados = [];

  for (var i = 1; i < ventasData.length; i++) {
    var numeroPedidoTemporal = ventasData[i][numeroPedidoTemporalIndex];
    if (numeroPedidoTemporal === numeroPedido) {
      var id = ventasData[i][idIndex];
      idsEnviados.push(id);
    }
  }

  if (idsEnviados.length === 0) {
    Logger.log("No se encontraron registros en Ventas para el pedido " + numeroPedido);
    return;
  }

  // Paso 3: Actualizar la hoja del proveedor correspondiente
  var proveedorSheet = spreadsheet.getSheetByName(proveedor);
  if (!proveedorSheet) {
    Logger.log("Hoja del proveedor " + proveedor + " no encontrada.");
    return;
  }

  // Obtener los datos de la hoja del proveedor
  var proveedorData = proveedorSheet.getDataRange().getValues();
  if (proveedorData.length < 2) {
    Logger.log("No hay datos en la hoja del proveedor " + proveedor + " (menos de 2 filas).");
    return;
  }

  var proveedorHeaders = proveedorData[0];
  var proveedorIdIndex = proveedorHeaders.indexOf("ID");

  if (proveedorIdIndex === -1) {
    Logger.log("Columna 'ID' no encontrada en la hoja del proveedor " + proveedor);
    return;
  }

  // Índice de la columna EnVenta (índice 11, columna K)
  var enVentaIndex = 10; // Índice 11 en base 1 es 10 en base 0 (columna K)

  // Actualizar la columna EnVenta con la fecha del pedido
  var updatedRows = 0;
  for (var i = 1; i < proveedorData.length; i++) {
    var proveedorId = proveedorData[i][proveedorIdIndex];
    if (idsEnviados.includes(proveedorId)) {
      proveedorSheet.getRange(i + 1, enVentaIndex + 1).setValue(fechaPedido);
      updatedRows++;
    }
  }

  Logger.log("Se actualizaron " + updatedRows + " filas en la hoja " + proveedor + " para el pedido " + numeroPedido);
}

function marcarEnVentaPorNumeroPedido(numeroPedidoEspecifico) {
  // Obtener la hoja activa
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener las hojas
  var ventasSheet = spreadsheet.getSheetByName("Ventas");
  var enviosSheet = spreadsheet.getSheetByName("Envios");

  // Paso 1: Buscar el pedido en la hoja "Envios"
  var enviosData = enviosSheet.getDataRange().getValues();
  var enviosHeaders = enviosData[0];
  var numeroPedidoIndex = -1;
  var fechaIndex = -1;
  var proveedorIndex = -1;

  for (var i = 0; i < enviosHeaders.length; i++) {
    var header = enviosHeaders[i].toString().toLowerCase().trim();
    if (header === "número de pedido" || header === "numero de pedido") {
      numeroPedidoIndex = i;
    } else if (header === "fecha" || header === "date" || header === "fecha pedido") {
      fechaIndex = i;
    } else if (header === "proveedor") {
      proveedorIndex = i;
    }
  }

  if (numeroPedidoIndex === -1 || fechaIndex === -1 || proveedorIndex === -1) {
    Logger.log("No se encontraron todas las columnas necesarias en la hoja Envios.");
    return;
  }

  var fechaPedido = null;
  var proveedor = null;
  for (var i = 1; i < enviosData.length; i++) {
    if (enviosData[i][numeroPedidoIndex] === numeroPedidoEspecifico) {
      fechaPedido = enviosData[i][fechaIndex];
      proveedor = enviosData[i][proveedorIndex];
      break;
    }
  }

  if (!fechaPedido || !proveedor) {
    Logger.log("Pedido " + numeroPedidoEspecifico + " no encontrado o datos incompletos en Envios.");
    return;
  }

  // Paso 2: Obtener los registros de "Ventas" que corresponden a este pedido
  var ventasData = ventasSheet.getDataRange().getValues();
  var ventasHeaders = ventasData[0];
  var numeroPedidoTemporalIndex = ventasHeaders.indexOf("Numero de pedido Temporal");
  var idIndex = ventasHeaders.indexOf("ID");

  var idsEnviados = [];
  for (var i = 1; i < ventasData.length; i++) {
    var numeroPedidoTemporal = ventasData[i][numeroPedidoTemporalIndex];
    if (numeroPedidoTemporal === numeroPedidoEspecifico) {
      var id = ventasData[i][idIndex];
      idsEnviados.push(id);
    }
  }

  if (idsEnviados.length === 0) {
    Logger.log("No se encontraron registros en Ventas para el pedido " + numeroPedidoEspecifico);
    return;
  }

  // Paso 3: Actualizar la hoja del proveedor correspondiente
  var proveedorSheet = spreadsheet.getSheetByName(proveedor);
  if (!proveedorSheet) {
    Logger.log("Hoja del proveedor " + proveedor + " no encontrada.");
    return;
  }

  var proveedorData = proveedorSheet.getDataRange().getValues();
  var proveedorHeaders = proveedorData[0];
  var proveedorIdIndex = proveedorHeaders.indexOf("ID");

  var enVentaIndex = 10; // Índice 11 en base 1 es 10 en base 0 (columna K)

  var updatedRows = 0;
  for (var i = 1; i < proveedorData.length; i++) {
    var proveedorId = proveedorData[i][proveedorIdIndex];
    if (idsEnviados.includes(proveedorId)) {
      proveedorSheet.getRange(i + 1, enVentaIndex + 1).setValue(fechaPedido);
      updatedRows++;
    }
  }

  Logger.log("Se actualizaron " + updatedRows + " filas en la hoja " + proveedor + " para el pedido " + numeroPedidoEspecifico);
}