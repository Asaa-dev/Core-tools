const CONFIG = {
  // Reservados para futuras funciones de generación documental:
  //folderId: "1j8UTyKhhDdxqrs8XRxWbDH-bPCFg98Dd",
  //docTemplateId: "10VgbuoK-rmvl0OVRv7w1OUCNtn30PNuHlmo0pwVOc24",
  //hojaLogs: "Data",
  //logSheetId: "1VlhDd5zrmWs6FNnk_SR_lUoxVQBAXn1fD5VdEUr1rXQ",
  //gestion: "Documentación",
  //area: "Corporativo",

  sheetName: "Gestión de registro",
  // docTemplateName: "Perfilamiento",
  // colAccion: "Instrucción",
  // claveDuplicado: "Identificador",
  // claveAgendamiento: "Número de agendamiento",
  camposObligatorios: [
    //"Identificador",
    "Contacto",
    "Campaña",
    "Anuncio",
    "Fecha de inicio",
    "Fecha de finalización",
    "Confirmación de tratamientos",
    "Escala",
    "Estado",
    "Observación",
  ],
};

// Ejecución individual: Depuración
function ejecutarDepuracion() {
  if (typeof Coretools?.main === "function") {
    Logger.log("Ejecución: depuración");
    Coretools.main();
  } else {
    Logger.log("Error: Coretools.main() no está definida.");
    SpreadsheetApp.getUi().alert("Función depuración no disponible");
  }
}

// Función ejecución generación no disponible
function ejecutarGeneracion(origen = "Manual") {
  SpreadsheetApp.getUi().alert("Función generación no disponible.");

  /*
  // Habilitar función generación
  if (typeof Coretools?.generacionDocumento === "function") {
    Logger.log(`Ejecución: documentación (${origen})`);
    const configConOrigen = { ...CONFIG, origenEjecucion: origen };
    Coretools.generacionDocumento(configConOrigen);
  } else {
    Logger.log("Error: Coretools.generacionDocumento() no está definida.");
    SpreadsheetApp.getUi().alert("Función generación no disponible.");
  }
  */
}

// Ejecución activador automático.
function ejecucionActivador() {
  ejecutarDepuracion();
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

// Ejecución compiplación inhabilitada.
function ejecutarTodo() {
  SpreadsheetApp.getUi().alert("Función compilación no disponible.");

  /*
  // Habilitar función generación.
  Logger.log("Ejecución: general");
  ejecutarDepuracion();
  ejecutarGeneracion();
  Logger.log("Ejecución finalizada");
  */
}

// Menú personalizado.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Cooasefin");
  const usuario = Session.getActiveUser().getEmail();

  // menu.addItem("Compilación", "ejecutarTodo");

  if (usuario === "angel.arciniegas@cooasefin.com.co") {
    menu.addItem("Ejecución periódica", "creacionActivador");
  }

  // menu.addSeparator();

  menu.addItem("Depuración formato", "ejecutarDepuracion");
  // menu.addItem("Generación documental", "ejecutarGeneracion");

  menu.addToUi();
}
