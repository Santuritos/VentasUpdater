/**
 * Script: Actualizador de Ventas - Versión 1.5 (Reconstrucción total - Precio desde Datos_ISBN)
 * Fecha: 16 mar 2025
 *
 * Descripción:
 * Se reconstruye la hoja "Ventas" a partir de los datos de "Datos_ISBN" y se sobrescriben
 * de forma especial algunos campos:
 *
 *   - "Precio": se copia el valor que viene en Datos_ISBN (no se calcula por prioridad).
 *   - "Hamelyn", "Momox" y "Ammareal": se asignan sus respectivos precios (incluso si son 0),
 *     obtenidos de las hojas correspondientes.
 *   - "Marcado": se asigna "FALSE" si está vacío.
 *
 * La estructura (encabezado) de la hoja Ventas es:
 * ID | ISBN | EAN | Titulo | Tipo | Situacion | Posicion | Estado | Portada | Portada_Link | 
 * Portada_Imagen | Fecha entrada | Precio | Hamelyn | Momox | Ammareal | Precio_eBay | 
 * Precio_Todocoleccion | Precio_amazon | Venta | Proveedor | Estado Temporal | Marcado | Numero de pedido Temporal
 *
 * Se copia de Datos_ISBN en aquellas columnas que tengan el mismo nombre. Las que no se encuentren
 * en Datos_ISBN se dejarán en blanco.
 */
function actualizarVentasDesdeDatosISBN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVentas  = ss.getSheetByName("Ventas");
  const hojaISBN    = ss.getSheetByName("Datos_ISBN");
  const hojaHamelyn = ss.getSheetByName("Hamelyn");
  const hojaMomox   = ss.getSheetByName("Momox");
  const hojaAmmareal = ss.getSheetByName("Ammareal");

  if (!hojaVentas || !hojaISBN || !hojaHamelyn || !hojaMomox || !hojaAmmareal) {
    throw new Error("Verifica que existan las hojas 'Ventas', 'Datos_ISBN', 'Hamelyn', 'Momox' y 'Ammareal'.");
  }
  
  // 1. Leer los datos de las hojas (incluyendo cabeceras)
  const datosISBN = hojaISBN.getDataRange().getValues();
  const cabeceraISBN = datosISBN[0];
  
  const datosHamelyn = hojaHamelyn.getDataRange().getValues();
  const cabeceraHamelyn = datosHamelyn[0];
  
  const datosMomox = hojaMomox.getDataRange().getValues();
  const cabeceraMomox = datosMomox[0];
  
  const datosAmmareal = hojaAmmareal.getDataRange().getValues();
  const cabeceraAmmareal = datosAmmareal[0];
  
  const datosVentasOld = hojaVentas.getDataRange().getValues();
  const cabeceraVentas = datosVentasOld[0];
  
  // 2. Mapear los nombres de columna de Ventas y de Datos_ISBN en objetos (nombre -> índice).
  let indicesVentas = {};
  cabeceraVentas.forEach((col, i) => { indicesVentas[col] = i; });
  
  let indicesISBN = {};
  cabeceraISBN.forEach((col, i) => { indicesISBN[col] = i; });
  
  // 3. Crear mapas de precios desde Hamelyn, Momox y Ammareal (se incluyen incluso los 0)
  let preciosHamelyn = {};
  for (let i = 1; i < datosHamelyn.length; i++) {
    const id = datosHamelyn[i][ cabeceraHamelyn.indexOf("ID") ]?.toString().trim();
    const precio = datosHamelyn[i][ cabeceraHamelyn.indexOf("Precio") ];
    if (id && (precio || precio === 0) && precio.toString().trim() !== "") {
      preciosHamelyn[id] = precio;
    }
  }
  
  let preciosMomox = {};
  for (let i = 1; i < datosMomox.length; i++) {
    const id = datosMomox[i][ cabeceraMomox.indexOf("ID") ]?.toString().trim();
    const precio = datosMomox[i][ cabeceraMomox.indexOf("Precio") ];
    if (id && (precio || precio === 0) && precio.toString().trim() !== "") {
      preciosMomox[id] = precio;
    }
  }
  
  let preciosAmmareal = {};
  for (let i = 1; i < datosAmmareal.length; i++) {
    const id = datosAmmareal[i][ cabeceraAmmareal.indexOf("ID") ]?.toString().trim();
    const precio = datosAmmareal[i][ cabeceraAmmareal.indexOf("Precio") ];
    if (id && (precio || precio === 0) && precio.toString().trim() !== "") {
      preciosAmmareal[id] = precio;
    }
  }
  
  // 4. Reconstruir la nueva tabla para Ventas. La primera fila será el encabezado actual.
  let ventasNew = [];
  ventasNew.push(cabeceraVentas);
  
  // 5. Por cada registro de Datos_ISBN (excluyendo la cabecera), se crea una nueva fila para Ventas.
  // Se copia el valor de cada columna si el nombre existe en Datos_ISBN; de lo contrario, se deja en blanco.
  for (let i = 1; i < datosISBN.length; i++) {
    const filaISBN = datosISBN[i];
    const id = filaISBN[ indicesISBN["ID"] ]?.toString().trim();
    if (!id) continue;
    
    // Crear una fila con la misma longitud que el encabezado de Ventas.
    let newRow = new Array(cabeceraVentas.length).fill("");
    
    // Para cada columna del encabezado de Ventas, copiar el valor si existe en Datos_ISBN.
    for (let j = 0; j < cabeceraVentas.length; j++) {
      const colName = cabeceraVentas[j];
      if (indicesISBN.hasOwnProperty(colName)) {
        newRow[j] = filaISBN[ indicesISBN[colName] ];
      } // Si no existe, dejamos en blanco.
    }
    
    // 6. Fijar "Precio" con el valor de Datos_ISBN (sin cálculos de prioridad)
    // Se asume que la columna "Precio" existe en Datos_ISBN (y se copió en el paso anterior).
    // Así, no se sobreescribe.
    
    // 7. Asignar los precios individuales para las columnas "Hamelyn", "Momox" y "Ammareal"
    newRow[ indicesVentas["Hamelyn"] ] = preciosHamelyn.hasOwnProperty(id) ? preciosHamelyn[id] : "";
    newRow[ indicesVentas["Momox"] ]   = preciosMomox.hasOwnProperty(id) ? preciosMomox[id] : "";
    newRow[ indicesVentas["Ammareal"] ]  = preciosAmmareal.hasOwnProperty(id) ? preciosAmmareal[id] : "";
    
    // 8. Si "Marcado" está vacío, asignar "FALSE" por defecto.
    if (!newRow[ indicesVentas["Marcado"] ] || newRow[ indicesVentas["Marcado"] ].toString().trim() === "") {
      newRow[ indicesVentas["Marcado"] ] = "FALSE";
    }
    
    ventasNew.push(newRow);
  }
  
  // 9. Reemplazar todo el contenido de la hoja Ventas con la nueva tabla.
  hojaVentas.clearContents();
  hojaVentas.getRange(1, 1, ventasNew.length, ventasNew[0].length).setValues(ventasNew);
  
  Logger.log("Actualización completa de Ventas.");
}
