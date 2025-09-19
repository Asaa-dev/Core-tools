/**
 * Función principal: Unificación.
 */
const config_unificacion = {
  // Hojas fuente.
  sheetSettingName: "Setting",
  sheetDataName: "Data",

  // Hojas logs.
  logSheetId: "1j8mkYwH9ylKPYtHwJ-LdsaGb4pNnna9H4TE69sOcpBU",
  hojaLogs: "Data",

  // Información setting.
  colSetting: {
    identificador: "Identificador",
    fuente: "Fuente",
    ubicacion: "Ubicación",
    hoja: "Hoja",
    area: "Área",
    estado: "Estado",
    observacion: "Observación",
  },

  // Información data.
  colData: {
    identificador: "Identificador",
    fuente: "Fuente",
    area: "Área",
    numeroIdentificacion: "Número de identificación",
    numeroAgendamiento: "Número de agendamiento",
    nombresCompletos: "Nombres completos",
    contacto: "Contacto",
    ubicacion: "Ubicación",
    pagaduria: "Pagaduría",
    entidad: "Entidad",
    modalidad: "Modalidad",
    monto: "Monto",
    cuota: "Cuota",
    plazo: "Plazo",
    tasa: "Tasa",
    estado: "Estado",
    observacion: "Observación",
    responsableComercial: "Responsable comercial",
    responsableVenta: "Responsable de venta",
    responsableCaptacion: "Responsable de captación",
    campaña: "Campaña",
    fechaAgendamiento: "Fecha de agendamiento",
    fechaConsulta: "Fecha de consulta",
    fechaRecepcion: "Fecha de recepción",
    fechaRadicacion: "Fecha de radicación",
    fechaDesembolso: "Fecha de desembolso",
  },

  // Parámetros data.
  camposObligatorios: ["Identificador", "Fuente", "Número de identificación"],

  // Condicional destino.
  exportConfigs: [
    {
      id: "1Jmsdm6OunpInPdmsW-EsVmRUdHojGJVkz3GuY7zWH_0",
      sheet: "Gestión de registro",
      rangoIds: { inicio: "DA-001", fin: "DA-014" },
    },
    {
      id: "1kdsS1w-N_iSngiaBBFbj35jr-Vfl1Mni1-_1eRefBg0",
      sheet: "Gestión de registro",
      rangoIds: { inicio: "DA-014", fin: "DA-024" },
    },
  ],
};

/**
 * Función normalización.
 */
function condicionFila(row, len) {
  if (!Array.isArray(row)) return Array(len).fill("");
  if (row.length === len) return row;
  if (row.length === 0) return Array(len).fill("");
  if (row.length < len) return [...row, ...Array(len - row.length).fill("")];
  return row.slice(0, len);
}

function normalizacion2D(rows, len) {
  return rows
    .map((r) => condicionFila(r, len))
    .filter((row) => row.length === len);
}

/**
 * Función principal.
 */
function ejecucionGeneral() {
  procesarFuentes();
  exportarDatos();
}

/**
 * Proceso fuentes.
 */
function procesarFuentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSetting = ss.getSheetByName(config_unificacion.sheetSettingName);
  const sheetData = ss.getSheetByName(config_unificacion.sheetDataName);
  if (!sheetSetting || !sheetData)
    throw new Error("Información: Hojas parámetros");

  const [headersSetting, ...rowsSetting] = sheetSetting
    .getDataRange()
    .getValues();
  const colEstadoIndex =
    headersSetting.indexOf(config_unificacion.colSetting.estado) + 1;

  const headerDataRow = sheetData
    .getRange(1, 1, 1, sheetData.getLastColumn())
    .getValues()[0];

  if (!headerDataRow || headerDataRow.length === 0) {
    throw new Error("Información: Hoja, data columnas no definidas.");
  }

  const colMapData = {};
  headerDataRow.forEach((colName, index) => (colMapData[colName] = index));

  const lastDataRow = sheetData.getLastRow();
  let dataExistente = [];
  if (lastDataRow > 1) {
    dataExistente = sheetData
      .getRange(2, 1, lastDataRow - 1, sheetData.getLastColumn())
      .getValues();
  }

  const keyMap = {};
  dataExistente.forEach((fila, i) => {
    const numeroId =
      fila[colMapData[config_unificacion.colData.numeroIdentificacion]];
    const entidad = fila[colMapData[config_unificacion.colData.entidad]] || "";
    const pagaduria =
      fila[colMapData[config_unificacion.colData.pagaduria]] || "";
    const key = numeroId + "|" + entidad + "|" + pagaduria;
    keyMap[key] = i;
  });

  const filasNuevas = [];

  rowsSetting.forEach((fila, rowIndex) => {
    const filaNumber = rowIndex + 2;
    let filaSettingObj = {};

    try {
      headersSetting.forEach((h, i) => {
        filaSettingObj[h] = fila[i];
      });

      const estado = filaSettingObj[config_unificacion.colSetting.estado];
      if (estado !== "Scheduled" && estado !== "Retry") return;

      sheetSetting.getRange(filaNumber, colEstadoIndex).setValue("Running");

      const dataFuente = obtenerDataFuente(filaSettingObj);

      if (!Array.isArray(dataFuente) || dataFuente.length === 0) {
        sheetSetting.getRange(filaNumber, colEstadoIndex).setValue("Succeeded");
        return;
      }

      dataFuente.forEach((filaData) => {
        const numeroId =
          filaData[config_unificacion.colData.numeroIdentificacion];
        const entidad = filaData[config_unificacion.colData.entidad] || "";
        const pagaduria = filaData[config_unificacion.colData.pagaduria] || "";
        const key = numeroId + "|" + entidad + "|" + pagaduria;

        const filaEscribir = new Array(headerDataRow.length).fill("");
        Object.keys(filaData).forEach((colName) => {
          if (colMapData.hasOwnProperty(colName)) {
            filaEscribir[colMapData[colName]] = filaData[colName];
          }
        });

        if (keyMap.hasOwnProperty(key)) {
          dataExistente[keyMap[key]] = filaEscribir;
        } else {
          filasNuevas.push(filaEscribir);
          keyMap[key] = dataExistente.length + filasNuevas.length - 1;
        }
      });

      sheetSetting.getRange(filaNumber, colEstadoIndex).setValue("Succeeded");

      registrarLog({
        gestion: "Unificación fuentes",
        area:
          filaSettingObj[config_unificacion.colSetting.area] ||
          "No especificado",
        estado: "Succeeded",
        mensaje: `Fuente ${
          filaSettingObj[config_unificacion.colSetting.identificador]
        } Proceso exitoso`,
        usuario: Session.getActiveUser().getEmail() || "Desconocido",
        origen: "Manual",
        enlace: "",
      });
    } catch (error) {
      try {
        sheetSetting.getRange(filaNumber, colEstadoIndex).setValue("Retry");
      } catch (e) {}

      if (!String(error).toLowerCase().includes("No datos")) {
        registrarLog({
          gestion: "Unificación fuentes",
          area:
            filaSettingObj[config_unificacion.colSetting.area] ||
            "No especifico",
          estado: "Failed",
          mensaje: `Error proceso fuente. ${
            filaSettingObj[config_unificacion.colSetting.identificador]
          }: ${error}`,
          usuario: Session.getActiveUser().getEmail() || "Desconocido",
          origen: "Manual",
          enlace: "",
        });
      }

      Logger.log(
        `Error proceso fuente ${
          filaSettingObj[config_unificacion.colSetting.identificador]
        }: ${error}`
      );
    }
  });

  if (dataExistente.length > 0) {
    let dataExistenteNormalizada = normalizacion2D(
      dataExistente,
      headerDataRow.length
    );
    dataExistenteNormalizada = dataExistenteNormalizada.filter((r) =>
      r.some((c) => c !== "" && c !== null && c !== undefined)
    );

    if (dataExistenteNormalizada.length > 0) {
      sheetData
        .getRange(2, 1, dataExistenteNormalizada.length, headerDataRow.length)
        .setValues(dataExistenteNormalizada);
    }
  }

  if (filasNuevas.length > 0) {
    const filasSoloArrays = filasNuevas.filter((f) => Array.isArray(f));
    let filasNormalizadas = normalizacion2D(
      filasSoloArrays,
      headerDataRow.length
    );
    filasNormalizadas = filasNormalizadas.filter((r) =>
      r.some((c) => c !== "" && c !== null && c !== undefined)
    );

    if (filasNormalizadas.length > 0) {
      const startRow = sheetData.getLastRow() + 1 || 2;
      sheetData
        .getRange(startRow, 1, filasNormalizadas.length, headerDataRow.length)
        .setValues(filasNormalizadas);
    }
  }
}

/**
 * Obtener información parámetros setting.
 */
function obtenerDataFuente(filaSetting) {
  const sheetOrigen = SpreadsheetApp.openById(
    filaSetting[config_unificacion.colSetting.ubicacion]
  ).getSheetByName(filaSetting[config_unificacion.colSetting.hoja]);

  if (!sheetOrigen)
    throw new Error(
      `Hoja ${filaSetting[config_unificacion.colSetting.hoja]} no disponible`
    );

  const [headers, ...rows] = sheetOrigen.getDataRange().getValues();
  if (!rows || rows.length === 0) return [];

  return rows.map((row) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    obj[config_unificacion.colData.identificador] =
      filaSetting[config_unificacion.colSetting.identificador];
    obj[config_unificacion.colData.fuente] =
      filaSetting[config_unificacion.colSetting.fuente];
    obj[config_unificacion.colData.area] =
      filaSetting[config_unificacion.colSetting.area];
    return obj;
  });
}

/**
 * Exportación de datos
 */
function exportarDatos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetData = ss.getSheetByName(config_unificacion.sheetDataName);
    const [headersData, ...rowsData] = sheetData.getDataRange().getValues();

    function idToNumber(id) {
      if (!id || typeof id !== "string") return NaN;
      const m = id.match(/DA-0*([0-9]+)/);
      return m ? parseInt(m[1], 10) : NaN;
    }

    config_unificacion.exportConfigs.forEach((destino) => {
      try {
        const ssDestino = SpreadsheetApp.openById(destino.id);
        const sheetDestino = ssDestino.getSheetByName(destino.sheet);
        if (!sheetDestino)
          throw new Error(
            `Información: ubicación no disponible ${destino.sheet}`
          );

        const headersDestino = sheetDestino
          .getRange(1, 1, 1, sheetDestino.getLastColumn())
          .getValues()[0];

        const inicioNum = idToNumber(destino.rangoIds.inicio);
        const finNum = idToNumber(destino.rangoIds.fin);

        const registrosFiltrados = rowsData.filter((row) => {
          const id =
            row[headersData.indexOf(config_unificacion.colData.identificador)];
          const num = idToNumber(id);
          return !isNaN(num) && num >= inicioNum && num <= finNum;
        });

        const esPerfilamiento =
          destino.rangoIds.inicio === "DA-014" &&
          destino.rangoIds.fin === "DA-024";

        const registrosTransformados = registrosFiltrados.map((row) => {
          const obj = {};
          headersData.forEach((h, i) => (obj[h] = row[i]));
          if (esPerfilamiento && headersDestino.includes("Responsable")) {
            obj["Responsable"] =
              obj[config_unificacion.colData.responsableVenta] ||
              obj[config_unificacion.colData.responsableComercial] ||
              obj[config_unificacion.colData.responsableCaptacion] ||
              "";
          }
          return headersDestino.map((h) => obj[h] || "");
        });

        if (sheetDestino.getLastRow() > 1) {
          sheetDestino
            .getRange(
              2,
              1,
              sheetDestino.getLastRow() - 1,
              headersDestino.length
            )
            .clearContent();
        }
        if (registrosTransformados.length > 0) {
          const registrosOk = normalizacion2D(
            registrosTransformados,
            headersDestino.length
          );
          sheetDestino
            .getRange(2, 1, registrosOk.length, headersDestino.length)
            .setValues(registrosOk);
        }

        registrarLog({
          gestion: "Exportación",
          area: "Corporativo",
          estado: "Succeeded",
          mensaje: `Exportación ${registrosTransformados.length} información a ${destino.sheet}`,
          usuario: Session.getActiveUser().getEmail() || "Desconocido",
          origen: "Automático",
          enlace: ssDestino.getUrl(),
        });
      } catch (e) {
        registrarLog({
          gestion: "Exportación",
          area: "Corporativo",
          estado: "Failed",
          mensaje: `Error exportación de datos ${destino.sheet}: ${e}`,
          usuario: Session.getActiveUser().getEmail() || "Desconocido",
          origen: "Automático",
          enlace: destino.id,
        });
      }
    });
  } catch (e) {
    Logger.log(`Información: error de exportación ${e}`);
  }
}

/**
 * Registro log.
 */
function registrarLog(info) {
  try {
    const logSpreadsheet = SpreadsheetApp.openById(
      config_unificacion.logSheetId
    );
    let hojaLog = logSpreadsheet.getSheetByName(config_unificacion.hojaLogs);

    if (!hojaLog) {
      hojaLog = logSpreadsheet.insertSheet(config_unificacion.hojaLogs);
      hojaLog.appendRow([
        "Fecha de registro",
        "Gestión",
        "Área",
        "Usuario",
        "Origen",
        "Estado",
        "Mensaje",
        "Documentación",
      ]);
    }

    const fecha = new Date();
    hojaLog.appendRow([
      fecha,
      info.gestion || "",
      info.area || "",
      info.usuario || "",
      info.origen || "",
      info.estado || "",
      info.mensaje || "",
      info.enlace || "",
    ]);
  } catch (e) {
    Logger.log(`Información: error de registro ${e.message}`);
  }
}
