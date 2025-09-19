const CONFIG = {
  folderId: "10KUuLGKley2fEyRgPYXZ6smjqZ1D72zw",
  sheetName: "Gestión de registro",
  docTemplateId: "10VgbuoK-rmvl0OVRv7w1OUCNtn30PNuHlmo0pwVOc24",
  docTemplateName: "Perfilamiento",
  colAccion: "Instrucción",
  claveDuplicado: "Número de identificación",
  claveAgendamiento: "Número de agendamiento",
  camposObligatorios: [
    // "Número de agendamiento",
    // "Número de identificación",
    "Nombres completos",
    // "Contacto",
    // "Ubicación",
    "Campaña",
    "Pagaduría",
    "Entidad",
    "Modalidad",
    "Monto",
    // "Cuota",
    "Plazo",
    "Tasa",
    "Fecha de agendamiento",
    "Responsable de venta",
    "Documento",
    "Estado",
    "Instrucción",
    // "Observación",
  ],
  hojaLogs: "Data",
  logSheetId: "1VlhDd5zrmWs6FNnk_SR_lUoxVQBAXn1fD5VdEUr1rXQ",

  gestion: "Documentación",
  area: "Ventas",
};
// Ejecución individual: Depuración
function ejecutarDepuracion() {
  if (typeof Coretools?.main === "function") {
    Logger.log("Ejecución: depuración");
    Coretools.main();
  } else {
    Logger.log("Error: Coretools.main() no está definida.");
  }
}

// Ejecución individual: Generación documental
function ejecutarGeneracion(origen = "Manual") {
  if (typeof Coretools?.generacionDocumento === "function") {
    Logger.log(`Ejecución: documentación (${origen})`);
    const configConOrigen = { ...CONFIG, origenEjecucion: origen };
    Coretools.generacionDocumento(configConOrigen);
  } else {
    Logger.log("Error: Coretools.generacionDocumento() no está definida.");
  }
}

// Ejecución activador automático
function ejecucionActivador() {
  ejecutarDepuracion();
  ejecutarGeneracion("Automático");
}

// Activador automático.
function creacionActivador() {
  const usuario = Session.getActiveUser().getEmail();
  const autorizado = "angel.arciniegas@cooasefin.com.co";

  if (usuario !== autorizado) {
    throw new Error("No tienes permiso para crear activadores.");
  }

  const horas = [6, 12, 19];
  const funcionObjetivo = "ejecucionActivador";

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === funcionObjetivo) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  horas.forEach((hora) => {
    ScriptApp.newTrigger(funcionObjetivo)
      .timeBased()
      .atHour(hora)
      .everyDays(1)
      .create();
  });
}

// Ejecución completa
function ejecutarTodo() {
  Logger.log("Ejecución: general");
  ejecutarDepuracion();
  ejecutarGeneracion();
  Logger.log("Ejecución finalizada");
}

// Menú personalizado.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Cooasefin");
  const usuario = Session.getActiveUser().getEmail();

  menu.addItem("Compilación", "ejecutarTodo");

  if (usuario === "angel.arciniegas@cooasefin.com.co") {
    menu.addItem("Ejecución periódica", "creacionActivador");
  }

  menu.addSeparator();

  menu.addItem("Depuración formato", "ejecutarDepuracion");
  menu.addItem("Generación documental", "ejecutarGeneracion");

  menu.addToUi();
}
