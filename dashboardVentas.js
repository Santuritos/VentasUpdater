/**
 * Script: dashboardVentas.gs
 * Versión: 1.0.8
 * Autor: Grok (xAI)
 * Fecha: 12-abr-2025
 * Descripción: Script para filtrar y mostrar datos de ventas en un dashboard.
 */

/**
 * Muestra el panel lateral con el formulario de filtros.
 */
function mostrarPanel() {
  const html = HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Dashboard de Ventas")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Obtiene las posiciones únicas de la hoja "Ventas" (columna G).
 * @returns {string[]} Lista de posiciones únicas.
 */
function obtenerPosiciones() {
  const hojaVentas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ventas");
  if (!hojaVentas) throw new Error("No se encontró la hoja 'Ventas'.");

  const ultimaFila = hojaVentas.getLastRow();
  if (ultimaFila < 2) return [];

  const datos = hojaVentas.getRange(2, 7, ultimaFila - 1, 1).getValues(); // Columna G (Posición)
  const posiciones = [...new Set(datos.flat().filter(pos => pos))].sort();
  return posiciones;
}

/**
 * Limpia la hoja "Dashboard" para prepararla para nuevos datos, preservando los estados de los checkboxes.
 * @returns {Object} Mapa de IDs a estados de checkboxes.
 */
function limpiarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDashboard = ss.getSheetByName("Dashboard");

  if (!hojaDashboard) {
    hojaDashboard = ss.insertSheet("Dashboard");
  }

  // Guardar los estados de los checkboxes antes de limpiar
  const ultimaFila = hojaDashboard.getLastRow();
  const checkboxStates = {};
  if (ultimaFila >= 2) {
    const datos = hojaDashboard.getRange(2, 1, ultimaFila - 1, 2).getValues(); // Columnas A:B (✔, ID)
    datos.forEach(row => {
      const id = row[1]; // Columna B (ID)
      const checked = row[0]; // Columna A (✔)
      if (id) {
        checkboxStates[id] = checked;
      }
    });
  }

  // Eliminar cualquier filtro existente
  const filtroExistente = hojaDashboard.getFilter();
  if (filtroExistente) {
    filtroExistente.remove();
  }

  // Limpiar contenido e imágenes
  hojaDashboard.getRange("A2:K" + hojaDashboard.getLastRow()).clearContent();
  hojaDashboard.getImages().forEach(img => img.remove());
  hojaDashboard.getRange("L1").setFormula("=SUM(K:K)");

  return checkboxStates;
}

/**
 * Procesa las ventas marcadas: genera un PDF y crea una nueva hoja con la fecha y el comprador.
 */
function procesarVentaMarcada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDashboard = ss.getSheetByName("Dashboard");
  if (!hojaDashboard) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Dashboard'.");
    return;
  }

  const ultimaFila = hojaDashboard.getLastRow();
  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert("No hay datos en la hoja 'Dashboard'.");
    return;
  }

  // Obtener la columna seleccionada (comprador)
  const scriptProperties = PropertiesService.getScriptProperties();
  const comprador = scriptProperties.getProperty("COLUMNA_SELECCIONADA") || "Hamelyn";

  // Obtener las filas marcadas
  const datos = hojaDashboard.getRange(2, 1, ultimaFila - 1, 11).getValues(); // Columnas A:K
  const filasMarcadas = [];
  datos.forEach(row => {
    if (row[0] === true) { // Columna A (✔)
      filasMarcadas.push(row);
    }
  });

  if (filasMarcadas.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas marcadas para procesar.");
    return;
  }

  // Crear una nueva hoja con la fecha y el comprador
  const fecha = new Date();
  const nombreHoja = `Venta_${fecha.toISOString().split('T')[0]}_${comprador}`;
  let hojaVenta = ss.getSheetByName(nombreHoja);
  if (hojaVenta) {
    hojaVenta.clear();
  } else {
    hojaVenta = ss.insertSheet(nombreHoja);
  }

  // Preparar los datos para la nueva hoja y el PDF
  const encabezados = ["ID", "ISBN", "Título", "Precio", "Posición", "Hamelyn", "Momox", "Ammareal", "Portada", "Subtotal", "Comprador", "Fecha"];
  const datosVenta = filasMarcadas.map(row => [
    row[1], // ID
    row[2], // ISBN
    row[3], // Título
    row[4], // Precio
    row[5], // Posición
    row[6], // Hamelyn
    row[7], // Momox
    row[8], // Ammareal
    row[9], // Portada
    row[10], // Subtotal
    comprador,
    fecha.toISOString().split('T')[0]
  ]);

  // Volcar datos en la nueva hoja
  hojaVenta.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hojaVenta.getRange(2, 1, datosVenta.length, encabezados.length).setValues(datosVenta);

  // Generar el PDF
  const pdfContent = `
    <h1>Venta procesada - ${comprador}</h1>
    <p><strong>Fecha:</strong> ${fecha.toISOString().split('T')[0]}</p>
    <table border="1" style="border-collapse: collapse; width: 100%;">
      <thead>
        <tr>
          ${encabezados.map(enc => `<th>${enc}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${datosVenta.map(row => `
          <tr>
            ${row.map(cell => `<td>${cell}</td>`).join('')}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  const htmlBlob = HtmlService.createHtmlOutput(pdfContent).getBlob().setName(`Venta_${fecha.toISOString().split('T')[0]}_${comprador}.pdf`);
  const pdfFile = DriveApp.createFile(htmlBlob);

  // Mostrar enlace al PDF
  SpreadsheetApp.getUi().alert(`PDF generado: ${pdfFile.getUrl()}`);

  // Desmarcar las filas procesadas en Dashboard
  datos.forEach((row, index) => {
    if (row[0] === true) {
      hojaDashboard.getRange(index + 2, 1).setValue(false);
    }
  });
}

/**
 * Marca todas las casillas de verificación en la columna A.
 */
function marcarTodos() {
  const hojaDashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!hojaDashboard) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Dashboard'.");
    return;
  }

  const ultimaFila = hojaDashboard.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const rango = hojaDashboard.getRange("A2:A" + ultimaFila);
  rango.setValue(true); // Marca todas las casillas
}

/**
 * Desmarca todas las casillas de verificación en la columna A.
 */
function desmarcarTodos() {
  const hojaDashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!hojaDashboard) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Dashboard'.");
    return;
  }

  const ultimaFila = hojaDashboard.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const rango = hojaDashboard.getRange("A2:A" + ultimaFila);
  rango.setValue(false); // Desmarca todas las casillas
}

/**
 * Marca las filas que han sido procesadas pero rechazadas (Hamelyn, Momox, Ammareal = 0.00).
 */
function marcarRechazoManual() {
  const hojaDashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!hojaDashboard) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Dashboard'.");
    return;
  }

  const ultimaFila = hojaDashboard.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const datos = hojaDashboard.getRange(2, 7, ultimaFila - 1, 3).getValues(); // Columnas G:I (Hamelyn, Momox, Ammareal)
  const checkboxes = hojaDashboard.getRange(2, 1, ultimaFila - 1, 1); // Columna A (✔)

  datos.forEach((fila, index) => {
    const hamelyn = Number(fila[0]);
    const momox = Number(fila[1]);
    const ammareal = Number(fila[2]);

    // Si todos los valores son 0.00, marcar la casilla (asumiendo que esto indica un rechazo)
    if (hamelyn === 0 && momox === 0 && ammareal === 0) {
      checkboxes.getCell(index + 1, 1).setValue(true);
    }
  });
}

/**
 * Marca las filas donde la columna seleccionada en el panel tiene valor 0.00 (rechazo desde panel).
 */
function marcarRechazoDesdePanel() {
  const hojaDashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if (!hojaDashboard) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Dashboard'.");
    return;
  }

  // Obtener la columna seleccionada almacenada en las propiedades del script
  const scriptProperties = PropertiesService.getScriptProperties();
  const columnaSeleccionada = scriptProperties.getProperty("COLUMNA_SELECCIONADA") || "Hamelyn";

  const ultimaFila = hojaDashboard.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const colLetras = { Hamelyn: 7, Momox: 8, Ammareal: 9 }; // Columnas G, H, I
  const colIndex = colLetras[columnaSeleccionada] - 7; // Índice relativo para datos (0, 1, 2)

  const datos = hojaDashboard.getRange(2, 7, ultimaFila - 1, 3).getValues(); // Columnas G:I
  const checkboxes = hojaDashboard.getRange(2, 1, ultimaFila - 1, 1); // Columna A (✔)

  datos.forEach((fila, index) => {
    const valor = Number(fila[colIndex]);
    if (valor === 0) {
      checkboxes.getCell(index + 1, 1).setValue(true);
    }
  });
}

/**
 * Filtra y vuelca los datos en la hoja "Dashboard" según los filtros proporcionados.
 * @param {Object} filtro - Objeto con los filtros: { posiciones: string[], columnas: string[], color: string }
 */
function filtrarYSumar(filtro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVentas = ss.getSheetByName("Ventas");
  if (!hojaVentas) throw new Error("No se encontró la hoja 'Ventas'.");

  // Almacenar la columna seleccionada en las propiedades del script
  const scriptProperties = PropertiesService.getScriptProperties();
  if (filtro.columnas.length > 0) {
    scriptProperties.setProperty("COLUMNA_SELECCIONADA", filtro.columnas[0]);
  }

  // Limpiar la hoja Dashboard y obtener los estados de los checkboxes
  const checkboxStates = limpiarDashboard();

  // Obtener datos de la hoja Ventas
  const ultimaFila = hojaVentas.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const datos = hojaVentas.getRange(2, 1, ultimaFila - 1, 16).getValues(); // Columnas A:P
  const colores = hojaVentas.getRange(2, 14, ultimaFila - 1, 3).getBackgrounds(); // Columnas N:P (Hamelyn, Momox, Ammareal)

  // Preparar la hoja Dashboard
  const hojaDashboard = ss.getSheetByName("Dashboard");
  const encabezados = ["✔", "ID", "ISBN", "Título", "Precio", "Posición", "Hamelyn", "Momox", "Ammareal", "Portada", "Subtotal"];
  hojaDashboard.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  // Mapear IDs a enlaces de portada
  const idToLinkMap = {};
  datos.forEach(row => {
    const id = row[0]; // Columna A (ID)
    const link = row[9]; // Columna J (Portada)
    if (id && link) idToLinkMap[id] = link;
  });

  // Filtrar y construir filas para el Dashboard
  const filasResultado = [];
  const colLetras = { Hamelyn: "G", Momox: "H", Ammareal: "I" };
  const colIndices = { Hamelyn: 13, Momox: 14, Ammareal: 15 }; // Índices en datos (N, O, P)
  const colColorIndices = { Hamelyn: 0, Momox: 1, Ammareal: 2 }; // Índices en colores

  datos.forEach((fila, index) => {
    const posicion = fila[6]; // Columna G (Posición)
    if (filtro.posiciones.length > 0 && !filtro.posiciones.includes(posicion)) return;

    const valores = {
      Hamelyn: fila[13], // Columna N
      Momox: fila[14],   // Columna O
      Ammareal: fila[15] // Columna P
    };
    const coloresFila = {
      Hamelyn: colores[index][0].toLowerCase(),
      Momox: colores[index][1].toLowerCase(),
      Ammareal: colores[index][2].toLowerCase()
    };

    let incluir = false;
    for (const proveedor of filtro.columnas) {
      const colorMatch = !filtro.color || filtro.color.toLowerCase() === coloresFila[proveedor];
      if (Number(valores[proveedor]) > 0 && colorMatch) {
        incluir = true;
        break;
      }
    }

    if (incluir) {
      const idFila = fila[0];
      const linkPortada = idToLinkMap[idFila] || "";
      const filaNum = filasResultado.length + 2;
      const colLetra = colLetras[filtro.columnas[0]] || "G"; // Usar la primera columna seleccionada

      // Restaurar el estado del checkbox si existe
      const checkboxState = checkboxStates[idFila] || false;

      filasResultado.push([
        checkboxState, // Checkbox (restaurado)
        fila[0],      // ID
        fila[2],      // ISBN
        fila[3],      // Título
        fila[12],     // Precio (columna M)
        fila[6],      // Posición (columna G)
        fila[13],     // Hamelyn
        fila[14],     // Momox
        fila[15],     // Ammareal
        linkPortada,  // Portada (texto plano)
        `=IF(A${filaNum};${colLetra}${filaNum};"")` // Subtotal (usando ; como separador)
      ]);
    }
  });

  // Volcar datos en la hoja Dashboard
  if (filasResultado.length > 0) {
    const numFilas = filasResultado.length;
    hojaDashboard.getRange(2, 1, numFilas, encabezados.length).setValues(filasResultado);
    hojaDashboard.getRange(2, 1, numFilas).insertCheckboxes();

    // Aplicar formato
    hojaDashboard.setRowHeights(2, numFilas, 20);
    hojaDashboard.getRange(2, 1, numFilas, 2).setHorizontalAlignment("center"); // ✔, ID
    hojaDashboard.getRange(2, 3, numFilas).setHorizontalAlignment("center"); // ISBN
    hojaDashboard.getRange(2, 4, numFilas).setHorizontalAlignment("left"); // Título
    hojaDashboard.getRange(2, 5, numFilas).setHorizontalAlignment("center"); // Precio
    hojaDashboard.getRange(2, 6, numFilas).setHorizontalAlignment("center"); // Posición
    hojaDashboard.getRange(2, 7, numFilas, 4).setNumberFormat("#,##0.00"); // Hamelyn, Momox, Ammareal
    hojaDashboard.getRange(2, 11, numFilas).setNumberFormat("#,##0.00"); // Subtotal
    hojaDashboard.getRange(2, 11, numFilas).setHorizontalAlignment("right"); // Subtotal

    // Agregar filtro
    hojaDashboard.getRange(1, 1, numFilas + 1, encabezados.length).createFilter();
  }

  SpreadsheetApp.flush();
}