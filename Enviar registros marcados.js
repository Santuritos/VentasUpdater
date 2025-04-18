function enviarRegistrosMarcados(proveedor) {
  try {
    // Obtener la hoja de cálculo usando el Spreadsheet ID
    var spreadsheet = SpreadsheetApp.openById("1r9gpYhGLZDUt5NI0rCMMsqA8RosJBk3IHBy6y6Yovy0");

    // Obtener las hojas de cálculo usando Sheet IDs
    var ventasSheet = spreadsheet.getSheetById(537746285); // Sheet ID de "Ventas"
    var enviosSheet = spreadsheet.getSheetById(1253902922); // Sheet ID de "Envíos"
    var pedidosSheet = spreadsheet.getSheetById(1156952776); // Sheet ID de "Pedidos"

    // Verificar que las hojas existan
    if (!ventasSheet) throw new Error("No se encontró la hoja 'Ventas' con Sheet ID 537746285");
    if (!enviosSheet) throw new Error("No se encontró la hoja 'Envíos' con Sheet ID 1253902922");
    if (!pedidosSheet) throw new Error("No se encontró la hoja 'Pedidos' con Sheet ID 1156952776");

    // Obtener los datos de las hojas
    var ventasData = ventasSheet.getDataRange().getValues();
    var enviosData = enviosSheet.getDataRange().getValues();
    var pedidosData = pedidosSheet.getDataRange().getValues();

    // Generar un "Número de pedido" único
    var numeroPedido = "PED-" + Utilities.getUuid().substring(0, 8);

    // Encontrar los índices de las columnas en "Ventas"
    var ventasHeaders = ventasData[0];
    var idIndex = ventasHeaders.indexOf("ID");
    var marcadoIndex = ventasHeaders.indexOf("Marcado");
    var ventaIndex = ventasHeaders.indexOf("Venta");
    var tituloIndex = ventasHeaders.indexOf("Titulo");
    var posicionIndex = ventasHeaders.indexOf("Posicion");
    var hamelynIndex = ventasHeaders.indexOf("Hamelyn");
    var momoxIndex = ventasHeaders.indexOf("Momox");
    var ammarealIndex = ventasHeaders.indexOf("Ammareal");

    // Verificar que las columnas existan en "Ventas"
    if (idIndex === -1) throw new Error("No se encontró la columna 'ID' en 'Ventas'");
    if (marcadoIndex === -1) throw new Error("No se encontró la columna 'Marcado' en 'Ventas'");
    if (ventaIndex === -1) throw new Error("No se encontró la columna 'Venta' en 'Ventas'");
    if (tituloIndex === -1) throw new Error("No se encontró la columna 'Titulo' en 'Ventas'");
    if (posicionIndex === -1) throw new Error("No se encontró la columna 'Posicion' en 'Ventas'");
    if (hamelynIndex === -1) throw new Error("No se encontró la columna 'Hamelyn' en 'Ventas'");
    if (momoxIndex === -1) throw new Error("No se encontró la columna 'Momox' en 'Ventas'");
    if (ammarealIndex === -1) throw new Error("No se encontró la columna 'Ammareal' en 'Ventas'");

    // Determinar el índice de la columna de precio según el proveedor
    var precioIndex;
    if (proveedor === "Hamelyn") {
      precioIndex = hamelynIndex;
    } else if (proveedor === "Momox") {
      precioIndex = momoxIndex;
    } else if (proveedor === "Ammareal") {
      precioIndex = ammarealIndex;
    } else {
      throw new Error("Proveedor no válido: " + proveedor + ". Debe ser 'Hamelyn', 'Momox' o 'Ammareal'.");
    }

    // Encontrar los índices de las columnas en "Envíos"
    var enviosHeaders = enviosData[0];
    var enviosIdIndex = enviosHeaders.indexOf("id");
    var enviosNumeroPedidoIndex = enviosHeaders.indexOf("Número de pedido");
    var enviosFechaIndex = enviosHeaders.indexOf("Fecha");
    var enviosIdArticuloIndex = enviosHeaders.indexOf("ID Artículo");
    var enviosProveedorIndex = enviosHeaders.indexOf("Proveedor");
    var enviosEstadoIndex = enviosHeaders.indexOf("Estado");
    var enviosTituloIndex = enviosHeaders.indexOf("Título");
    var enviosPrecioProveedorIndex = enviosHeaders.indexOf("Precio");

    // Verificar que las columnas existan en "Envíos"
    if (enviosIdIndex === -1) throw new Error("No se encontró la columna 'id' en 'Envíos'");
    if (enviosNumeroPedidoIndex === -1) throw new Error("No se encontró la columna 'Número de pedido' en 'Envíos'");
    if (enviosFechaIndex === -1) throw new Error("No se encontró la columna 'Fecha' en 'Envíos'");
    if (enviosIdArticuloIndex === -1) throw new Error("No se encontró la columna 'ID Artículo' en 'Envíos'");
    if (enviosProveedorIndex === -1) throw new Error("No se encontró la columna 'Proveedor' en 'Envíos'");
    if (enviosEstadoIndex === -1) throw new Error("No se encontró la columna 'Estado' en 'Envíos'");
    if (enviosTituloIndex === -1) throw new Error("No se encontró la columna 'Título' en 'Envíos'");
    if (enviosPrecioProveedorIndex === -1) throw new Error("No se encontró la columna 'Precio' en 'Envíos'");

    // Calcular el Total y procesar los registros marcados
    var total = 0;
    var registrosEnviados = 0;
    var fechaActual = new Date();

    for (var i = 1; i < ventasData.length; i++) {
      var row = ventasData[i];
      var marcado = row[marcadoIndex];
      var venta = row[ventaIndex];

      // Verificar si el registro está marcado y pendiente
      if (marcado === true && venta === "Pendiente") {
        // Añadir el registro a "Envíos"
        var precio = parseFloat(row[precioIndex]);
        enviosSheet.appendRow([
          Utilities.getUuid(), // id
          numeroPedido, // Número de pedido
          fechaActual, // Fecha
          row[idIndex], // ID Artículo
          proveedor, // Proveedor
          "Pendiente", // Estado
          row[tituloIndex], // Título
          precio // Precio (según el proveedor)
        ]);

        // Actualizar el registro en "Ventas"
        ventasSheet.getRange(i + 1, marcadoIndex + 1).setValue(false); // Marcar como FALSE
        ventasSheet.getRange(i + 1, ventaIndex + 1).setValue("Enviado"); // Venta = "Enviado"

        // Sumar al total
        if (!isNaN(precio)) {
          total += precio;
        }

        registrosEnviados++;
      }
    }

    // Añadir el pedido a la tabla "Pedidos"
    if (registrosEnviados > 0) {
      pedidosSheet.appendRow([
        numeroPedido, // ID
        numeroPedido, // Número de pedido
        fechaActual, // Fecha actual
        proveedor, // Proveedor
        total // Total
      ]);
    }

    return {
      success: true,
      message: "Se enviaron " + registrosEnviados + " registros a Envíos con Número de pedido: " + numeroPedido + ", Total: " + total
    };
  } catch (error) {
    return {
      success: false,
      message: "Error: " + error.message
    };
  }
}

// Función para probar el script desde el editor
function probarEnviarRegistrosMarcados() {
  var result = enviarRegistrosMarcados("Hamelyn");
  Logger.log(result);
}