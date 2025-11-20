function SaludAsuMedida(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPrincipal = ss.getSheetByName(nombreHoja);

  if (!hojaPrincipal) {
    SpreadsheetApp.getUi().alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }

  const ULTIMA_FILA = hojaPrincipal.getLastRow();
  if (ULTIMA_FILA < 2) {
    Logger.log("No hay datos para procesar en la hoja principal.");
    return;
  }

  const COL_POLIZA = 0; // columna E (índice 0 relativo al rango seleccionado)
  const COL_CORREO = 6; // columna K
  const COL_CC1 = 7;     // columna L
  const COL_CC2 = 9;     // columna N
  const NUM_COLUMNAS_RANGO = 10;
  const COL_INICIO_RESULTADOS = 16;

  const rangoDatos = hojaPrincipal.getRange(2, 5, ULTIMA_FILA - 1, NUM_COLUMNAS_RANGO).getValues();

  const limpiarCC = d => String(d || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').replace(/\s/g, '').trim().toLowerCase();
  const limpiarPoliza = p => String(p || '').trim();

  const limpiarValorCondicional = valor => {
    const s = String(valor || '');
    if (s.includes('@')) return limpiarCorreo(s);
    return limpiarCC(s);
  };

  const formatearFecha = fecha => {
    if (!fecha) return '';
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || 'GMT-5', "yyyy-MM-dd HH:mm:ss");
  };

  function cargarDatosYMapa(nombreHoja, idColumnas, fechaColumna) {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) return { map: new Map(), data: [] };

    const data = hoja.getDataRange().getValues().slice(1);

    if (fechaColumna !== null && nombreHoja !== "Revision duplicados") {
      data.sort((a, b) => (b[fechaColumna] ? new Date(b[fechaColumna]).getTime() : 0) - (a[fechaColumna] ? new Date(a[fechaColumna]).getTime() : 0));
    }

    const mapa = new Map();
    data.forEach(row => {
      idColumnas.forEach(colId => {
        let valor = row[colId];
        if (nombreHoja === "Revision duplicados") valor = limpiarPoliza(valor);
        else valor = limpiarValorCondicional(valor);

        if (valor && !mapa.has(valor)) mapa.set(valor, row);
      });
    });

    return { map: mapa, data: data };
  }

  const { map: duplicadosMap } = cargarDatosYMapa("Revision duplicados", [12], null);
  const esPolizaDuplicada = poliza => poliza && duplicadosMap.has(limpiarPoliza(poliza));

  const config = {
    leadsSalud: { name: "TOTAL LEADS SALUD LIGERO", ids: [6, 4], fecha: 12, infoCols: { fuente: 9, medio: 10, campaña: 11 } },
    leads322: { name: "Leads 322 - salud", ids: [11], fecha: 27, infoCols: { medio: 8, campaña: 24 }, fuente: "322" },
    referidos: { name: "Referidos Salud a su medida", ids: [0, 2], fecha: 6, infoCols: { medio: 7 }, fuente: "Referido" },
    bases: { name: "BASES SALUD A SU MEDIDA", ids: [0, 1], fecha: 2, infoCols: { fuente: 7, medio: 3, campaña: 6 } }
  };

  const { map: leadsSaludMap } = cargarDatosYMapa(config.leadsSalud.name, config.leadsSalud.ids, config.leadsSalud.fecha);
  const { map: leads322Map } = cargarDatosYMapa(config.leads322.name, config.leads322.ids, config.leads322.fecha);
  const { map: referidosMap } = cargarDatosYMapa(config.referidos.name, config.referidos.ids, config.referidos.fecha);
  const { map: basesMap } = cargarDatosYMapa(config.bases.name, config.bases.ids, config.bases.fecha);

  const encabezados = [
    "cc - LIGERO", "cc2 - LIGERO", "correo - LIGERO", "S", "T",
    "322", "Referidos", "Base CC", "Base mail",
    "ventas", "test", "fuente", "medio", "campaña", "fecha lead"
  ];
  hojaPrincipal.getRange(1, COL_INICIO_RESULTADOS, 1, encabezados.length).setValues([encabezados]);

  const resultados = [];

  rangoDatos.forEach(fila => {
    const poliza = limpiarPoliza(fila[COL_POLIZA]);
    const correo = limpiarCorreo(fila[COL_CORREO]);
    const cc1 = limpiarCC(fila[COL_CC1]);
    const cc2 = limpiarCC(fila[COL_CC2]);

    let testValue = "-";
    let skipLeadsSearch = false;

    if (esPolizaDuplicada(poliza)) {
      testValue = "DUPLICADO";
      skipLeadsSearch = true;
    }

    if (skipLeadsSearch) {
      resultados.push([0, 0, 0, "", "", 0, 0, 0, 0, 0, testValue, "", "", "", ""]);
      return;
    }

    const matchSaludCC1 = (cc1 && leadsSaludMap.has(cc1)) ? 1 : 0;
    const matchSaludCC2 = (cc2 && leadsSaludMap.has(cc2)) ? 1 : 0;
    const matchSaludCorreo = (correo && leadsSaludMap.has(correo)) ? 1 : 0;
    const match322CC = ((cc1 && leads322Map.has(cc1)) || (cc2 && leads322Map.has(cc2))) ? 1 : 0;
    const matchReferidos = ((cc1 && referidosMap.has(cc1)) || (cc2 && referidosMap.has(cc2)) || (correo && referidosMap.has(correo))) ? 1 : 0;
    const matchBaseCC = ((cc1 && basesMap.has(cc1)) || (cc2 && basesMap.has(cc2))) ? 1 : 0;
    const matchBaseMail = (correo && basesMap.has(correo)) ? 1 : 0;

    const ventas = matchSaludCC1 + matchSaludCC2 + matchSaludCorreo +
                   match322CC + matchReferidos + matchBaseCC + matchBaseMail;

    let fuente = "", medio = "", campana = "", fechaLead = null;
    let registro = null;

    if (matchSaludCC1 || matchSaludCC2 || matchSaludCorreo) {
      registro = leadsSaludMap.get(cc1) || leadsSaludMap.get(cc2) || leadsSaludMap.get(correo);
      if (registro) {
        fuente = registro[config.leadsSalud.infoCols.fuente] || '';
        medio = registro[config.leadsSalud.infoCols.medio] || '';
        campana = registro[config.leadsSalud.infoCols.campaña] || '';
        fechaLead = registro[config.leadsSalud.fecha];
      }
    } else if (match322CC) {
      registro = leads322Map.get(cc1) || leads322Map.get(cc2);
      if (registro) {
        fuente = config.leads322.fuente;
        medio = registro[config.leads322.infoCols.medio] || '';
        campana = registro[config.leads322.infoCols.campaña] || '';
        fechaLead = registro[config.leads322.fecha];
      }
    } else if (matchReferidos) {
      registro = referidosMap.get(cc1) || referidosMap.get(cc2) || referidosMap.get(correo);
      if (registro) {
        fuente = config.referidos.fuente;
        medio = registro[config.referidos.infoCols.medio] || '';
        campana = '';
        fechaLead = registro[config.referidos.fecha];
      }
    } else if (matchBaseCC || matchBaseMail) {
      registro = basesMap.get(cc1) || basesMap.get(correo) || basesMap.get(cc2);
      if (registro) {
        fuente = registro[config.bases.infoCols.fuente] || '';
        medio = registro[config.bases.infoCols.medio] || '';
        campana = registro[config.bases.infoCols.campaña] || '';
        fechaLead = registro[config.bases.fecha];
      }
    }

    resultados.push([
      matchSaludCC1, matchSaludCC2, matchSaludCorreo, "", "",
      match322CC, matchReferidos, matchBaseCC, matchBaseMail,
      ventas, testValue, fuente, medio, campana, formatearFecha(fechaLead)
    ]);
  });

  if (resultados.length > 0) {
    hojaPrincipal.getRange(2, COL_INICIO_RESULTADOS, resultados.length, resultados[0].length).setValues(resultados);
  }
}
