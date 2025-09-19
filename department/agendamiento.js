const CONFIG = {
  folderId: "1JmFoWklvJWh8TZfr8QBQ629WHrCLIabi",
  sheetName: "Gestión de registro",
  docTemplateId: "1TQ9khFJtkEOSj64RtaNCFniY02WNibV8LvF0oe4SfeA",
  docTemplateName: "Agendamiento",
  colAccion: "Instrucción",
  claveDuplicado: "Número de identificación",
  claveAgendamiento: "Número de agendamiento",
  camposObligatorios: [
    "Número de identificación",
    // "Número de agendamiento",
    "Nombres completos",
    // "Contacto",
    // "Ubicación",
    // "Pagaduría",
    "Entidad",
    "Modalidad",
    "Monto",
    "Cuota",
    "Plazo",
    "Tasa",
    "Fecha de agendamiento",
    "Documento",
    // "Responsable comercial",
    "Estado",
    "Instrucción",
    // "Observación",
  ],
  hojaLogs: "Data",
  logSheetId: "18l_HeoG4UA7m2xXLU7Rnwg1LT_PzkhpAWO2O5anWLyo",

  gestion: "Documentación",
  area: "Administración",
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
