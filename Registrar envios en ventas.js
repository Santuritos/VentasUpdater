/**
 * Script: Registrar Envíos en Ventas
 * Versión: 1.1
 * Fecha: 13 abr 2025
 * Descripción: Actualiza la hoja "Ventas" basándose en "Envios", estableciendo "Venta" = "Enviado" para registros coincidentes.
 */
function registrarEnviosEnVentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEnvios = ss.getSheetByName("Envios");
  const hojaVentas = ss.getSheetByName("Ventas");

  if (!hojaEnvios || !hojaVentas) {
    throw new Error("Verifica que existan las hojas 'Envios' y 'Ventas'.");
  }

  const enviosData = hojaEnvios.getDataRange().getValues();
  const ventasData = hojaVentas.getDataRange().getValues();

  if (enviosData.length < 2) {
    throw new Error("La hoja Envios no contiene registros.");
  }
  if (ventasData.length < 2) {
    throw new Error("La hoja Ventas no contiene registros para actualizar.");
  }

  const headerEnvios = enviosData[0];
  const headerVentas = ventasData[0];

  const numPedidoEnviosIndex = headerEnvios.indexOf("Número de pedido");
  const idArticuloEnviosIndex = headerEnvios.indexOf("ID Artículo");
  const proveedorEnviosIndex = headerEnvios.indexOf("Proveedor");

  if (numPedidoEnviosIndex === -1 || idArticuloEnviosIndex === -1 || proveedorEnviosIndex === -1) {
    throw new Error("Faltan las columnas 'Número de pedido', 'ID Artículo' o 'Proveedor' en la hoja Envios.");
  }

  const idVentasIndex = headerVentas.indexOf("ID");
  const ventaIndex = headerVentas.indexOf("Venta");
  const proveedorVentasIndex = headerVentas.indexOf("Proveedor");
  const estadoTemporalVentasIndex = headerVentas.indexOf("Estado Temporal");
  const marcadoVentasIndex = headerVentas.indexOf("Marcado");
  const numPedidoTempVentasIndex = headerVentas.indexOf("Numero de pedido Temporal");

  if ([idVentasIndex, ventaIndex, proveedorVentasIndex, estadoTemporalVentasIndex, marcadoVentasIndex, numPedidoTempVentasIndex].some(index => index === -1)) {
    throw new Error("Faltan columnas requeridas en Ventas: 'ID', 'Venta', 'Proveedor', 'Estado Temporal', 'Marcado' o 'Numero de pedido Temporal'.");
  }

  let ventasMap = {};
  for (let i = 1; i < ventasData.length; i++) {
    const row = ventasData[i];
    const id = row[idVentasIndex] ? row[idVentasIndex].toString().trim() : "";
    if (id) {
      ventasMap[id] = i;
    }
  }

  for (let i = 1; i < enviosData.length; i++) {
    const rowEnvios = enviosData[i];
    const idArticulo = rowEnvios[idArticuloEnviosIndex] ? rowEnvios[idArticuloEnviosIndex].toString().trim() : "";
    if (!idArticulo) continue;

    if (ventasMap.hasOwnProperty(idArticulo)) {
      const rowIndex = ventasMap[idArticulo];
      ventasData[rowIndex][ventaIndex] = "Enviado";
      ventasData[rowIndex][proveedorVentasIndex] = rowEnvios[proveedorEnviosIndex];
      ventasData[rowIndex][marcadoVentasIndex] = "TRUE";
      ventasData[rowIndex][numPedidoTempVentasIndex] = rowEnvios[numPedidoEnviosIndex];
      // No tocamos "Estado Temporal" para mantener "Pendiente" (ajustar si se necesita otro valor)
    }
  }

  hojaVentas.getRange(1, 1, ventasData.length, ventasData[0].length).setValues(ventasData);

  Logger.log("Registro de envíos en Ventas completado.");
}