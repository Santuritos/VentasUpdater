/**
 * Script: menu.gs
 * Descripción: Crea un menú personalizado en Google Sheets para ejecutar las funciones del proyecto.
 * Versión: 1.8
 * Fecha de Creación: 11/04/2025
 * Última Modificación: 11/04/2025
 * Autor: Grok 3 (xAI)
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📚 Gestión de Libros');

  // Submenú: Actualizar Precios e Inventario
  var actualizarMenu = ui.createMenu('Actualizar Precios e Inventario')
    .addItem('Sincronizar Todas las Hojas', 'sincronizarHojas')
    .addItem('Configurar Disparador Automático (Hora)', 'configurarDisparadorSincronizacion')
    .addItem('Configurar Disparador onChange (AppSheet)', 'configurarDisparadoresInstalables')
    .addItem('Actualizar Precios en Plataformas', 'actualizarPreciosPlataformas');

  // Submenú: Procesar Datos
  var procesarMenu = ui.createMenu('Procesar Datos')
    .addItem('Rellenar Datos desde Texto', 'rellenarDatosISBNDesdeTexto')
    .addItem('Eliminar Duplicados', 'eliminarDuplicadosISBN_EAN_Sencillo')
    .addItem('Enviar Registros Marcados', 'enviarRegistrosMarcados')
    .addItem('Procesar Formulario', 'procesarFormulario');

  // Submenú: Imágenes y PDFs
  var imagenesMenu = ui.createMenu('Imágenes y PDFs')
    .addItem('Actualizar Imágenes', 'actualizarImagenes')
    .addItem('Generar PDF para Hamelyn', 'generarPDFHamelyn')
    .addItem('Generar PDF para Momox', 'generarPDFMomox');

  // Submenú: Dashboards y Reportes
  var reportesMenu = ui.createMenu('Dashboards y Reportes')
    .addItem('Mostrar Dashboard de Ventas', 'mostrarPanel');

  // Submenú: Otros
  var otrosMenu = ui.createMenu('Otros')
    .addItem('Comparador', 'comparador')
    .addItem('En Venta', 'enVenta')
    .addItem('Consultar Open Library', 'consultarOpenLibrary');

  // Construir el menú principal
  menu
    .addSubMenu(actualizarMenu)
    .addSubMenu(procesarMenu)
    .addSubMenu(imagenesMenu)
    .addSubMenu(reportesMenu)
    .addSubMenu(otrosMenu)
    .addToUi();
}

/**
 * Funciones de redirección para evitar errores si los nombres de las funciones no coinciden.
 */
function sincronizarHojas() {
  try {
    sincronizarHojas();
  } catch (e) {
    Logger.log("Error al sincronizar todas las hojas: " + e.message);
    registrarError("sincronizarHojas", e.message);
  }
}

function configurarDisparadorSincronizacion() {
  try {
    configurarDisparadorSincronizacion();
  } catch (e) {
    Logger.log("Error al configurar el disparador automático: " + e.message);
    registrarError("configurarDisparadorSincronizacion", e.message);
  }
}

function configurarDisparadoresInstalables() {
  try {
    configurarDisparadoresInstalables();
  } catch (e) {
    Logger.log("Error al configurar el disparador onChange: " + e.message);
    registrarError("configurarDisparadoresInstalables", e.message);
  }
}

function actualizarPreciosPlataformas() {
  try {
    ActualizarPreciosPlataformas();
  } catch (e) {
    Logger.log("Error al actualizar precios en plataformas: " + e.message);
    registrarError("actualizarPreciosPlataformas", e.message);
  }
}

function enviarRegistrosMarcados() {
  try {
    EnviarRegistrosMarcados();
  } catch (e) {
    Logger.log("Error al enviar registros marcados: " + e.message);
    registrarError("enviarRegistrosMarcados", e.message);
  }
}

function procesarFormulario() {
  try {
    ProcesarFormulario();
  } catch (e) {
    Logger.log("Error al procesar formulario: " + e.message);
    registrarError("procesarFormulario", e.message);
  }
}

function actualizarImagenes() {
  try {
    ActualizarImagenes();
  } catch (e) {
    Logger.log("Error al actualizar imágenes: " + e.message);
    registrarError("actualizarImagenes", e.message);
  }
}

function generarPDFHamelyn() {
  try {
    GenerarPDF("Hamelyn");
  } catch (e) {
    Logger.log("Error al generar PDF para Hamelyn: " + e.message);
    registrarError("generarPDFHamelyn", e.message);
  }
}

function generarPDFMomox() {
  try {
    GenerarPDF("Momox");
  } catch (e) {
    Logger.log("Error al generar PDF para Momox: " + e.message);
    registrarError("generarPDFMomox", e.message);
  }
}

function mostrarPanel() {
  try {
    mostrarPanel();
  } catch (e) {
    Logger.log("Error al mostrar el panel de ventas: " + e.message);
    registrarError("mostrarPanel", e.message);
  }
}

function comparador() {
  try {
    Comparador();
  } catch (e) {
    Logger.log("Error al ejecutar comparador: " + e.message);
    registrarError("comparador", e.message);
  }
}

function enVenta() {
  try {
    EnVenta();
  } catch (e) {
    Logger.log("Error al ejecutar En Venta: " + e.message);
    registrarError("enVenta", e.message);
  }
}

function consultarOpenLibrary() {
  try {
    ConsultarOpenLibrary();
  } catch (e) {
    Logger.log("Error al consultar Open Library: " + e.message);
    registrarError("consultarOpenLibrary", e.message);
  }
}

/**
 * Función auxiliar para registrar errores en una hoja.
 */
function registrarError(funcion, mensaje) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaErrores = ss.getSheetByName("Errores");
  if (!hojaErrores) {
    hojaErrores = ss.insertSheet("Errores");
    hojaErrores.getRange(1, 1, 1, 3).setValues([["Fecha", "Función", "Error"]]);
  }
  const fecha = new Date().toLocaleString();
  hojaErrores.appendRow([fecha, funcion, mensaje]);
}