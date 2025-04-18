/**
 * Script: Colorear Precios y Marcar Enviados
 * Versión: 1.3
 * Fecha: 13 abr 2025
 * Descripción: Colorea filas en "Ventas" con rojo claro (#ffcccc) si "Venta" = "Enviado" y aplica colores a "Hamelyn", "Momox", "Ammareal" según el mayor precio.
 */
function colorearPreciosVentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVentas = ss.getSheetByName("Ventas");
  if (!hojaVentas) {
    Logger.log("Error: La hoja 'Ventas' no existe.");
    return;
  }

  const datos = hojaVentas.getDataRange().getValues();
  if (datos.length < 2) {
    Logger.log("No hay registros para procesar en 'Ventas'.");
    return;
  }

  const header = datos[0];
  const colIndexHamelyn = header.indexOf("Hamelyn");
  const colIndexMomox = header.indexOf("Momox");
  const colIndexAmmareal = header.indexOf("Ammareal");
  const colIndexVenta = header.indexOf("Venta");

  if (colIndexHamelyn === -1 || colIndexMomox === -1 || colIndexAmmareal === -1 || colIndexVenta === -1) {
    Logger.log("Error: No se encontraron las columnas 'Hamelyn', 'Momox', 'Ammareal' o 'Venta'. Encabezados: " + header.join(", "));
    return;
  }

  const colorMayorValor = "#f4ed7c"; // Dorado para el único valor máximo
  const colorEmpate = "#b3e5fc"; // Celeste para empate
  const sinColor = null; // Sin color
  const colorVentaEnviado = "#ffcccc"; // Rojo suave para "Enviado"

  Logger.log(`Procesando ${datos.length - 1} filas en 'Ventas'...`);

  // Colorear filas con "Venta" = "Enviado"
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    let valorVenta = row[colIndexVenta] ? row[colIndexVenta].toString().trim().toLowerCase() : "";
    Logger.log(`Fila ${i + 1}: Valor en 'Venta' = "${valorVenta}"`);

    if (valorVenta === "enviado") {
      Logger.log(`Fila ${i + 1}: Aplicando color rojo claro (${colorVentaEnviado})`);
      hojaVentas.getRange(i + 1, 1, 1, header.length).setBackground(colorVentaEnviado);
    } else {
      Logger.log(`Fila ${i + 1}: Restableciendo fondo (sin color)`);
      hojaVentas.getRange(i + 1, 1, 1, header.length).setBackground(sinColor);
    }
  }

  // Colorear celdas de precios para filas no enviadas
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    let valorVenta = row[colIndexVenta] ? row[colIndexVenta].toString().trim().toLowerCase() : "";
    if (valorVenta === "enviado") {
      Logger.log(`Fila ${i + 1}: Omitiendo coloreado de precios porque 'Venta' = 'Enviado'`);
      continue;
    }

    const precioHamelyn = parseFloat(row[colIndexHamelyn]) || 0;
    const precioMomox = parseFloat(row[colIndexMomox]) || 0;
    const precioAmmareal = parseFloat(row[colIndexAmmareal]) || 0;

    Logger.log(`Fila ${i + 1}: Precios - Hamelyn: ${precioHamelyn}, Momox: ${precioMomox}, Ammareal: ${precioAmmareal}`);

    if (precioHamelyn === 0 && precioMomox === 0 && precioAmmareal === 0) {
      Logger.log(`Fila ${i + 1}: Todos los precios son 0, no se colorea`);
      continue;
    }

    const precios = [
      { columna: colIndexHamelyn, valor: precioHamelyn },
      { columna: colIndexMomox, valor: precioMomox },
      { columna: colIndexAmmareal, valor: precioAmmareal }
    ];

    const maxValor = Math.max(precioHamelyn, precioMomox, precioAmmareal);
    const ganadores = precios.filter(obj => obj.valor === maxValor);

    const colorAAplicar = ganadores.length > 1 ? colorEmpate : colorMayorValor;
    Logger.log(`Fila ${i + 1}: Máximo valor = ${maxValor}, Color a aplicar = ${colorAAplicar}`);

    ganadores.forEach(obj => {
      hojaVentas.getRange(i + 1, obj.columna + 1).setBackground(colorAAplicar);
    });
  }

  Logger.log("Coloreado completado en 'Ventas'.");
}