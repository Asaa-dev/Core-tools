const CONFIG = {
  folderId: "1UvTBG0zf1j2b1Rm6ESolABpy7zA42PRb",
  sheetName: "Gestión de registro",
  docTemplateId: "1ryV_2Q9J6dcak9_op3Cjmu06tmFom7rqD84vSyrd4rY",
  docTemplateName: "Consulta",
  colAccion: "Instrucción",
  claveDuplicado: "Número de identificación",
  camposObligatorios: [
    "Número de identificación",
    "Nombres completos",
    // "Contacto",
    // "Ubicación",
    "Pagaduría",
    "Entidad",
    "Modalidad",
    "Monto",
    "Fecha de consulta",
    "Documento",
    // "Responsable comercial",
    "Estado",
    "Instrucción",
    // "Observación",
  ],
  hojaLogs: "Data",
  logSheetId: "1umafGdWyTm-AazD5pOulpo96SGDQEfaqi5TMhrtQF5U",

  gestion: "Documentación",
  area: "Crédito",
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
