/**
 * Función principal: Generación documental
 */
function generacionDocumento(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    registrarLog(`Hoja inválida: ${config.sheetName}`, config);
    return;
  }

  const [headers, ...rows] = sheet.getDataRange().getValues();
  const colIndex = {
    accion: headers.indexOf(config.colAccion),
    clave: headers.indexOf(config.claveDuplicado),
  };

  if (config.claveAgendamiento) {
    colIndex.agendamiento = headers.indexOf(config.claveAgendamiento);
    if (colIndex.agendamiento === -1) {
      registrarLog("Error: columna de agendamiento no encontrada", config);
      return;
    }
  }

  rows.forEach((fila, idx) => {
    const rowIndex = idx + 2;
    try {
      const datos = formateoFila(headers, fila);
      const identificacion = datos[config.claveDuplicado];
      const agendamiento = config.claveAgendamiento
        ? datos[config.claveAgendamiento]
        : "";
      const accion = fila[colIndex.accion]?.toString().trim().toLowerCase();

      if (!accion) {
        registrarLog(
          `Instrucción no registrada: Fila ${rowIndex} (Identificación: ${identificacion})`,
          config
        );
        return;
      }

      if (accion === "revisión") {
        registrarLog(
          `Información en revisión: Fila ${rowIndex} (Identificación: ${identificacion})`,
          config
        );
        return;
      }

      const faltantes = validacionCampos(datos, config);
      if (faltantes.length > 0) {
        registrarLog(
          `Información no registrada: Fila ${rowIndex}: ${faltantes.join(
            ", "
          )} (Identificación: ${identificacion})`,
          config
        );
        return;
      }

      const nombreArchivo = `${config.docTemplateName} ${identificacion}${
        agendamiento ? " " + agendamiento : ""
      }`;
      const archivo = buscarDocumento(nombreArchivo, config);

      let url = "";
      switch (accion) {
        case "generar":
          if (archivo) {
            registrarLog(
              `Documentación duplicada: ${nombreArchivo} (Identificación: ${identificacion})`,
              config
            );
            return;
          }
          url = generacionPlantilla(datos, nombreArchivo, config);
          registrarLog(
            `Documentación generada: ${url} (Identificación: ${identificacion})`,
            config
          );
          break;

        case "modificar":
          if (archivo) archivo.setTrashed(true);
          url = generacionPlantilla(datos, nombreArchivo, config);
          registrarLog(
            `Documentación modificada: ${url} (Identificación: ${identificacion})`,
            config
          );
          break;

        case "eliminar":
          if (!archivo) {
            registrarLog(
              `Documentación no registrada: ${nombreArchivo} (Identificación: ${identificacion})`,
              config
            );
            return;
          }
          archivo.setTrashed(true);
          registrarLog(
            `Documentación eliminada: ${nombreArchivo} (Identificación: ${identificacion})`,
            config
          );
          break;

        default:
          registrarLog(
            `Instrucción inválida: Fila ${rowIndex}: ${accion} (Identificación: ${identificacion})`,
            config
          );
      }
    } catch (error) {
      registrarLog(
        `Error fila ${rowIndex}: (Identificación: ${identificacion}) ${error.message}`,
        config
      );
    }
  });
}

function validacionCampos(datos, config) {
  return config.camposObligatorios.filter((campo) => {
    const valor = datos[campo];
    return !valor || valor.toString().trim() === "";
  });
}

function formateoFila(headers, values) {
  return headers.reduce((obj, key, i) => {
    obj[key] = values[i];
    return obj;
  }, {});
}

function generacionPlantilla(datos, nombreArchivo, config) {
  const folder = DriveApp.getFolderById(config.folderId);
  const copia = DriveApp.getFileById(config.docTemplateId).makeCopy(
    nombreArchivo,
    folder
  );
  const doc = DocumentApp.openById(copia.getId());
  reemplazoMarcador(doc, datos);
  doc.saveAndClose();
  return doc.getUrl();
}

function reemplazoMarcador(doc, datos) {
  const body = doc.getBody();
  const limitesPorCampo = {
    "Número de identificación": 10,
    "Número de agendamiento": 10,
    "Número de proceso": 15,
    "Nombres completos": 55,
    Contacto: 10,
    Ubicación: 25,
    "Dirección de residencia": 20,
    "Dirección de correo electrónico": 35,
    Pagaduría: 35,
    Campaña: 20,
    Anuncio: 15,
    Entidad: 15,
    Modalidad: 20,
    Monto: 15,
    Cuota: 15,
    Plazo: 3,
    Tasa: 5,
    "Fecha de agendamiento": 10,
    "Fecha de consulta": 10,
    "Fecha de recepción": 10,
    "Fecha de radicación": 10,
    "Fecha de devolución": 10,
    "Fecha de desembolso": 10,
    "Responsable comercial": 35,
    "Responsable de captación": 35,
    "Responsable de venta": 35,
    Estado: 32,
    Observación: 200,
  };

  for (let campo in datos) {
    const marcador = `{{${campo}}}`;
    let valor = datos[campo];

    const camposFecha = [
      "Fecha de agendamiento",
      "Fecha de consulta",
      "Fecha de recepción",
      "Fecha de radicación",
      "Fecha de devolución",
      "Fecha de desembolso",
    ];

    if (camposFecha.includes(campo)) {
      valor = formatearFecha(valor);
    } else if (campo === "Tasa") {
      const numero = parseFloat(valor);
      if (!isNaN(numero)) {
        valor = (numero * 100).toFixed(2) + "%";
      }
    } else if (campo === "Monto" || campo === "Cuota") {
      const numero = parseFloat(valor);
      if (!isNaN(numero)) {
        valor = new Intl.NumberFormat("es-CO", {
          style: "currency",
          currency: "COP",
          minimumFractionDigits: 0,
        }).format(numero);
      }
    }

    const limite = limitesPorCampo[campo] || 200;
    const textoAjustado = limitarTexto(valor, limite);
    body.replaceText(marcador, textoAjustado);
  }
}

function limitarTexto(texto, limite = 200) {
  if (!texto) return "";
  return texto.toString().length > limite
    ? texto.toString().substring(0, limite) + "…"
    : texto.toString();
}

function formatearFecha(fecha) {
  if (Object.prototype.toString.call(fecha) !== "[object Date]") {
    const posibleFecha = new Date(fecha);
    if (!isNaN(posibleFecha)) {
      fecha = posibleFecha;
    } else {
      return fecha;
    }
  }

  const dia = String(fecha.getDate()).padStart(2, "0");
  const mes = String(fecha.getMonth() + 1).padStart(2, "0");
  const anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

function buscarDocumento(nombreArchivo, config) {
  const folder = DriveApp.getFolderById(config.folderId);
  const archivos = folder.getFilesByName(nombreArchivo);
  return archivos.hasNext() ? archivos.next() : null;
}

function registrarLog(mensaje, config) {
  try {
    const logSpreadsheet = SpreadsheetApp.openById(config.logSheetId);
    let hojaLog = logSpreadsheet.getSheetByName(config.hojaLogs);

    if (!hojaLog) {
      hojaLog = logSpreadsheet.insertSheet(config.hojaLogs);
      hojaLog.appendRow([
        "Fecha",
        "Gestión",
        "Área",
        "Usuario",
        "Origen",
        "Estado",
        "Mensaje",
        "Enlace",
      ]);
    }

    const fecha = new Date();
    const gestion = config.gestion || "General";
    const area = config.area || "No especificada";
    const usuario = Session.getActiveUser().getEmail() || "Desconocido";
    const origen = config.origenEjecucion || "Manual";
    const estado = determinarEstado(mensaje);

    const urlRegex = /(https?:\/\/[^\s)]+)/;
    const urlMatch = mensaje.match(urlRegex);
    const url = urlMatch ? urlMatch[1] : "";

    const mensajeSinUrl = mensaje.replace(urlRegex, "").trim();
    const enlace = url ? `=HYPERLINK("${url}", "Documento generado")` : "";

    hojaLog.appendRow([
      fecha,
      gestion,
      area,
      usuario,
      origen,
      estado,
      mensajeSinUrl,
      enlace,
    ]);
  } catch (e) {
    Logger.log(`Error registro log: ${e.message}`);
  }

  const gestion = config.gestion ? `[${config.gestion}] ` : "";
  const usuario = Session.getActiveUser().getEmail() || "Desconocido";
  Logger.log(`${gestion}${mensaje} (Usuario: ${usuario})`);
}

function determinarEstado(mensaje) {
  if (mensaje.includes("generada")) return "Generado";
  if (mensaje.includes("modificada")) return "Modificado";
  if (mensaje.includes("eliminada")) return "Eliminado";
  if (mensaje.includes("duplicada")) return "Omitido";
  if (mensaje.includes("no registrada")) return "Omitido";
  if (mensaje.includes("Instrucción no registrada")) return "Omitido";
  if (mensaje.includes("Información no registrada")) return "Omitido";
  if (mensaje.includes("Información en revisión")) return "Información";
  if (mensaje.includes("Instrucción inválida")) return "Error";
  if (mensaje.includes("columna de agendamiento no encontrada")) return "Error";
  if (mensaje.includes("Hoja inválida")) return "Error";
  if (mensaje.includes("Error fila")) return "Error";
  return "Información";
}
