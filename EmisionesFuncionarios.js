function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Automatizaciones")
    .addItem("Emisiones", "iniciarNormalizacion")
    .addToUi();
}

function iniciarNormalizacion() {
  const ui = SpreadsheetApp.getUi();

  const respuesta = ui.prompt(
    "Iniciar Normalización",
    "Ingresa el nombre de la hoja donde se ejecutará la automatización:",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() === ui.Button.CANCEL) return;

  const nombreHoja = respuesta.getResponseText().trim();
  if (!nombreHoja) {
    ui.alert("El nombre de la hoja no puede estar vacío.");
    return;
  }

  ejecutarNormalizacion(nombreHoja);
}

function ejecutarNormalizacion(nombreHojaUsuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const hojaUsuario = ss.getSheetByName(nombreHojaUsuario);
  const hojaLeads = ss.getSheetByName("LEADS TOTAL");

  if (!hojaUsuario) {
    ui.alert(`La hoja "${nombreHojaUsuario}" no existe.`);
    return;
  }

  if (!hojaLeads) {
    ui.alert(`La hoja "LEADS TOTAL" no existe.`);
    return;
  }

  const ultimaFilaUsuario = hojaUsuario.getLastRow();
  if (ultimaFilaUsuario < 2) {
    ui.alert("No hay datos para procesar.");
    return;
  }

  const datosUsuario = hojaUsuario.getRange(2, 1, ultimaFilaUsuario - 1, 14).getValues();

  const ultimaFilaLeads = hojaLeads.getLastRow();
  if (ultimaFilaLeads < 2) {
    ui.alert("LEADS TOTAL no tiene datos.");
    return;
  }

  const datosLeads = hojaLeads.getRange(2, 1, ultimaFilaLeads - 1, 9).getValues();

  const mapa = crearMapaNormalizado(datosLeads);

  const resultados = [];

  datosUsuario.forEach(fila => {
    const cc = normalizarCC(fila[9]);
    const mail1 = normalizarCorreo(fila[11]);
    const mail2 = normalizarCorreo(fila[12]);
    const cel = normalizarCelular(fila[13]);
    const fechaBase = fila[7];

    let P = 0, Q = 0, R = 0, S = 0;
    if (cc && mapa.cc.has(cc)) P = 1;
    if (mail1 && mapa.correo.has(mail1)) Q = 1;
    if (mail2 && mapa.correo.has(mail2)) R = 1;
    if (cel && mapa.celular.has(cel)) S = 1;

    const T = P + Q + R + S;

    let match = null;

    if (P === 1) { 
      match = mapa.cc.get(cc);
    } else if (Q === 1) { 
      match = mapa.correo.get(mail1);
    } else if (R === 1) { 
      match = mapa.correo.get(mail2);
    } else if (S === 1) {
      match = mapa.celular.get(cel);
    }
    
    let fuente = "";
    let medio = "";
    let campana = "";
    let producto = "";
    let fechaForm = "";

    if (match) {
      fuente = match[4];
      medio = match[5];
      campana = match[6];
      producto = match[7];
      fechaForm = match[8];
    }

    let difDias = "";
    if (fechaBase instanceof Date && fechaForm instanceof Date) {
      difDias = Math.floor((fechaBase - fechaForm) / (1000 * 60 * 60 * 24));
    }

    resultados.push([
      P, Q, R, S, T,
      fuente, medio, campana, producto, fechaForm,
      difDias
    ]);
  });

  hojaUsuario
    .getRange(2, 16, resultados.length, 11)
    .setValues(resultados);

  ui.alert(`Proceso finalizado. Filas procesadas: ${resultados.length}`);
}

function crearMapaNormalizado(data) {
  const mapaCC = new Map();
  const mapaCorreo = new Map();
  const mapaCelular = new Map();

  data.forEach(fila => {
    const cc = normalizarCC(fila[1]);
    const correo = normalizarCorreo(fila[2]);
    const cel = normalizarCelular(fila[3]);

    if (cc && !mapaCC.has(cc)) mapaCC.set(cc, fila);
    if (correo && !mapaCorreo.has(correo)) mapaCorreo.set(correo, fila);
    if (cel && !mapaCelular.has(cel)) mapaCelular.set(cel, fila);
  });

  return { cc: mapaCC, correo: mapaCorreo, celular: mapaCelular };
}


function normalizarCC(valor) {
  if (!valor) return null;
  return String(valor).replace(/[^0-9]/g, "");
}

function normalizarCorreo(valor) {
  if (!valor) return null;
  return String(valor)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalizarCelular(valor) {
  if (!valor) return null;
  return String(valor).replace(/[^0-9]/g, "");
}