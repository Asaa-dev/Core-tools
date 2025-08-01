/**
 * Función principal: Lectura de información general.
 */

function main() {
  const hojasObjetivo = ["Gestión de registro"];
  hojasObjetivo.forEach((nombreHoja) => depuracionFormato(nombreHoja));
}

/**
 * Función principal: Función depuración.
 */
function depuracionFormato(nombreHoja) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) {
    Logger.log(`Hoja inválida: ${nombreHoja}`);
    return;
  }

  const numFilas = hoja.getLastRow();
  const numColumnas = hoja.getLastColumn();
  const rango = hoja.getDataRange();

  estiloBase(rango);

  const encabezado = hoja.getRange(1, 1, 1, numColumnas);
  estiloBase(encabezado);
  encabezado.setHorizontalAlignment("center").setVerticalAlignment("middle");

  rango.setBorder(false, false, false, false, false, false);

  // Tamaño columna.
  const anchoMaximo = 200;
  for (let col = 1; col <= numColumnas; col++) {
    hoja.autoResizeColumn(col);
    const anchoActual = hoja.getColumnWidth(col);
    if (anchoActual > anchoMaximo) {
      hoja.setColumnWidth(col, anchoMaximo);
    }
  }

  // Tamaño filas.
  const alturaEstandar = 21;
  for (let fila = 1; fila <= numFilas; fila++) {
    hoja.setRowHeight(fila, alturaEstandar);
  }

  capitalizarTexto(hoja, numFilas, numColumnas);
  alinearDatos(hoja, numFilas, numColumnas);
  ajusteEspacio(hoja, numFilas, numColumnas);
  formateoColor(hoja, numFilas, numColumnas);
}

function estiloBase(rango) {
  rango
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontStyle("normal")
    .setFontLine("none")
    .setBackground(null)
    .setFontColor(null);
}

function capitalizarTexto(hoja, numFilas, numColumnas) {
  if (numFilas < 2) return;
  const encabezados = hoja.getRange(1, 1, 1, numColumnas).getValues()[0];
  const rango = hoja.getRange(2, 1, numFilas - 1, numColumnas);
  const valores = rango.getValues();

  // Creamos nuevo array columna por columna
  for (let col = 0; col < numColumnas; col++) {
    const encabezado = encabezados[col];
    if (["Ubicación", "Documento"].includes(encabezado)) continue;

    const nuevaColumna = valores.map((fila) => {
      const celda = fila[col];
      if (typeof celda === "string") {
        return [
          celda
            .toLowerCase()
            .split(" ")
            .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
            .join(" "),
        ];
      }
      return [celda];
    });

    hoja.getRange(2, col + 1, numFilas - 1, 1).setValues(nuevaColumna);
  }
}

function alinearDatos(hoja, numFilas, numColumnas) {
  if (numFilas < 2) return;
  const encabezados = hoja.getRange(1, 1, 1, numColumnas).getValues()[0];
  const numFilasDatos = numFilas - 1;
  const datos = hoja.getRange(2, 1, numFilasDatos, numColumnas).getValues();

  for (let col = 0; col < numColumnas; col++) {
    const encabezado = encabezados[col];
    const columnaDatos = datos
      .map((fila) => fila[col])
      .filter((celda) => celda !== "" && celda !== null);
    if (columnaDatos.length === 0) continue;

    const tipos = columnaDatos.map((celda) => typeof celda);
    const conteoTipos = tipos.reduce((acc, tipo) => {
      acc[tipo] = (acc[tipo] || 0) + 1;
      return acc;
    }, {});

    const tipoDominante = Object.entries(conteoTipos).sort(
      (a, b) => b[1] - a[1]
    )[0][0];

    let alineacion;
    switch (tipoDominante) {
      case "number":
      case "object":
        alineacion = "right";
        break;
      case "boolean":
        alineacion = "center";
        break;
      default:
        alineacion = "left";
    }

    const rango = hoja.getRange(2, col + 1, numFilasDatos);
    rango.setHorizontalAlignment(alineacion);
    rango.setVerticalAlignment("middle");
  }

  hoja
    .getRange(1, 1, 1, numColumnas)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function ajusteEspacio(hoja, numFilas, numColumnas) {
  if (numFilas < 2) return;
  const encabezados = hoja.getRange(1, 1, 1, numColumnas).getValues()[0];
  const rango = hoja.getRange(2, 1, numFilas - 1, numColumnas);
  const valores = rango.getValues();

  for (let col = 0; col < numColumnas; col++) {
    const encabezado = encabezados[col];
    if (["Ubicación", "Documento"].includes(encabezado)) continue;

    const nuevaColumna = valores.map((fila) => {
      const celda = fila[col];
      if (typeof celda === "string") {
        return [celda.trim().replace(/\s+/g, " ")];
      }
      return [celda];
    });

    hoja.getRange(2, col + 1, numFilas - 1, 1).setValues(nuevaColumna);
  }
}

function formateoColor(hoja, numFilas, numColumnas) {
  if (numFilas === 0 || numColumnas === 0) return;

  hoja.getRange(1, 1, 1, numColumnas).setFontColor("#212121");

  if (numFilas > 1) {
    hoja.getRange(2, 1, numFilas - 1, numColumnas).setFontColor(null);
  }
}

function programarEjecucion() {
  ScriptApp.newTrigger("main").timeBased().atHour(7).everyDays(1).create();
}
