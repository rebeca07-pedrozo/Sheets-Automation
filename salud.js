function verificarLeadsIntegral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPrincipal = ss.getSheetByName("Copy of Emisiones Integral 22 sep");
  
  if (!hojaPrincipal) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'Copy of Emisiones Integral 14 sep'.");
    return;
  }

  const ultimaFila = hojaPrincipal.getLastRow();
  if (ultimaFila < 2) {
    Logger.log("No hay datos para procesar en la hoja principal.");
    return;
  }
  
  const rangoDatos = hojaPrincipal.getRange(2, 11, ultimaFila - 1, 4).getValues(); 

  const encabezados = [
    "cc1 LEADS TOTAL INTEGRAL", "correo LEADS TOTAL INTEGRAL", "cc2 LEADS TOTAL INTEGRAL",
    "322 CC1", "322 CC2", "Base CC1", "Base Mail", "Base CC2", "322 otros", "Referidos",
    "ventas", "fuente", "medio", "campaña", "fecha lead"
  ];
  hojaPrincipal.getRange(1, 16, 1, encabezados.length).setValues([encabezados]);

  function cargarDatosYMapa(nombreHoja, idColumnas, fechaColumna, fuenteFija = null) {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) {
      Logger.log(`Advertencia: La hoja '${nombreHoja}' no existe o no tiene datos.`);
      return { map: new Map(), info: null };
    }

    const data = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
    
    if (nombreHoja !== "LEADS TOTAL INTEGRAL" && fechaColumna !== null) {
      data.sort((a, b) => {
        const dateA = a[fechaColumna] ? new Date(a[fechaColumna]).getTime() : 0;
        const dateB = b[fechaColumna] ? new Date(b[fechaColumna]).getTime() : 0;
        return dateB - dateA;
      });
    }

    const mapa = new Map();
    data.forEach(row => {
      idColumnas.forEach(colId => {
        const valor = row[colId] ? String(row[colId]).trim() : "";
        if (valor && !mapa.has(valor)) {
          mapa.set(valor, row);
        }
      });
    });

    const info = {
      fuente: (fuenteFija !== null) ? fuenteFija : data[0] ? data[0][idColumnas[0]] : "",
      medio: "",
      campaña: "",
      fecha: fechaColumna
    };

    return { map: mapa, info };
  }

  const config = {
    integral: { name: "LEADS TOTAL INTEGRAL", ids: [2, 15], fecha: 8, infoCols: { fuente: 4, medio: 5, campaña: 6 } }, 
    leads322: { name: "Leads 322", ids: [11], fecha: 27, infoCols: { medio: 8, campaña: 25 }, fuente: "322" }, 
    referidos: { name: "Leads Referidos", ids: [0, 6], fecha: 1, fechaAlternativa: 7, fuente: "Referido" }, 
    bases: { name: "BASES INTEGRAL", ids: [0, 1], fecha: 2, infoCols: { fuente: 7, medio: "Prioridad", campaña: 6 } } 
  };

  const { map: integralMap } = cargarDatosYMapa(config.integral.name, config.integral.ids, config.integral.fecha);
  const { map: leads322Map } = cargarDatosYMapa(config.leads322.name, config.leads322.ids, config.leads322.fecha, config.leads322.fuente);
  const { map: referidosMap } = cargarDatosYMapa(config.referidos.name, config.referidos.ids, config.referidos.fecha, config.referidos.fuente);
  const { map: basesMap } = cargarDatosYMapa(config.bases.name, config.bases.ids, config.bases.fecha);

  const resultados = [];

  rangoDatos.forEach(fila => {
    const correo = fila[0] ? String(fila[0]).trim().toLowerCase() : ""; 
    const cc1 = fila[1] ? String(fila[1]).trim() : ""; 
    const cc2 = fila[3] ? String(fila[3]).trim() : ""; 

    const matchIntegralCC1 = (cc1 && integralMap.has(cc1)) ? 1 : 0;
    const matchIntegralCorreo = (correo && integralMap.has(correo)) ? 1 : 0;
    const matchIntegralCC2 = (cc2 && integralMap.has(cc2)) ? 1 : 0;
    const match322CC1 = (cc1 && leads322Map.has(cc1)) ? 1 : 0;
    const match322CC2 = (cc2 && leads322Map.has(cc2)) ? 1 : 0;
    const matchBaseCC1 = (cc1 && basesMap.has(cc1)) ? 1 : 0;
    const matchBaseMail = (correo && basesMap.has(correo)) ? 1 : 0;
    const matchBaseCC2 = (cc2 && basesMap.has(cc2)) ? 1 : 0;
    const match322Otros = (cc1 && referidosMap.has(cc1) && referidosMap.get(cc1)[config.referidos.ids[0]] !== cc1) ? 1 : 0;
    const matchReferidos = (cc1 && referidosMap.has(cc1) && referidosMap.get(cc1)[config.referidos.ids[0]] === cc1) ? 1 : 0;
    
    const ventas = matchIntegralCC1 + matchIntegralCorreo + matchIntegralCC2 +
                   match322CC1 + match322CC2 + matchBaseCC1 + matchBaseMail + matchBaseCC2 +
                   match322Otros + matchReferidos;

    let fuente = "", medio = "", campaña = "", fechaLead = "";
    let registro = null;

    if (matchIntegralCC1 || matchIntegralCorreo || matchIntegralCC2) {
      registro = integralMap.get(cc1) || integralMap.get(correo) || integralMap.get(cc2);
      if (registro) {
        fuente = registro[config.integral.infoCols.fuente];
        medio = registro[config.integral.infoCols.medio];
        campaña = registro[config.integral.infoCols.campaña];
        fechaLead = registro[config.integral.fecha];
      }
    } else if (match322CC1 || match322CC2) {
      registro = leads322Map.get(cc1) || leads322Map.get(cc2);
      if (registro) {
        fuente = config.leads322.fuente;
        medio = registro[config.leads322.infoCols.medio];
        campaña = registro[config.leads322.infoCols.campaña];
        fechaLead = registro[config.leads322.fecha];
      }
    } else if (matchReferidos || match322Otros) {
      registro = referidosMap.get(cc1);
      if (registro) {
        fuente = config.referidos.fuente;
        if (match322Otros) {
            fuente = "Referidos";
            fechaLead = registro[config.referidos.fechaAlternativa];
        } else {
            fechaLead = registro[config.referidos.fecha];
        }
        medio = "";
        campaña = "";
      }
    } else if (matchBaseCC1 || matchBaseMail || matchBaseCC2) {
      registro = basesMap.get(cc1) || basesMap.get(correo) || basesMap.get(cc2);
      if (registro) {
        fuente = registro[config.bases.infoCols.fuente];
        medio = config.bases.infoCols.medio;
        campaña = registro[config.bases.infoCols.campaña];
        fechaLead = registro[config.bases.fecha];
      }
    }
    
    let fechaFormateada = "";
    if (fechaLead instanceof Date) {
      fechaFormateada = Utilities.formatDate(fechaLead, "GMT-5", "yyyy-MM-dd HH:mm:ss");
    }

    resultados.push([
      matchIntegralCC1, matchIntegralCorreo, matchIntegralCC2,
      match322CC1, match322CC2,
      matchBaseCC1, matchBaseMail, matchBaseCC2,
      match322Otros, matchReferidos,
      ventas, fuente, medio, campaña, fechaFormateada
    ]);
  });

  if (resultados.length > 0) {
    hojaPrincipal.getRange(2, 16, resultados.length, resultados[0].length).setValues(resultados);
  }
}

// ACA EMPIEZA LA OTRA, SALUD A SU MEDIDA!!


function SaludAsuMedida() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPrincipal = ss.getSheetByName("Copy of Emisiones A su medida 22 sep"); 
  
  if (!hojaPrincipal) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'Copy of Emisiones A su medida 22 sep'.");
    return;
  }

  const ultimaFila = hojaPrincipal.getLastRow();
  if (ultimaFila < 2) {
    Logger.log("No hay datos para procesar en la hoja principal.");
    return;
  }
  
  const rangoDatos = hojaPrincipal.getRange(2, 5, ultimaFila - 1, 10).getValues(); 


  const limpiarCC = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').trim();
  const limpiarCorreo = c => String(c || '').replace(/\s/g, '').trim().toLowerCase(); 
  const limpiarPoliza = p => String(p || '').trim();

  const limpiarValorCondicional = valor => {
      const s = String(valor || '');
      if (s.includes('@')) {
          return limpiarCorreo(s);
      } else {
          return limpiarCC(s);
      }
  };

  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || 'GMT-5', "yyyy-MM-dd HH:mm:ss");
  }

  function cargarDatosYMapa(nombreHoja, idColumnas, fechaColumna) {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) {
      Logger.log(`Advertencia: La hoja '${nombreHoja}' no existe o no tiene datos.`);
      return { map: new Map(), data: [] };
    }

    const data = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
    
    if (fechaColumna !== null && nombreHoja !== "Revision duplicados") {
      data.sort((a, b) => {
        const dateA = a[fechaColumna] ? new Date(a[fechaColumna]).getTime() : 0;
        const dateB = b[fechaColumna] ? new Date(b[fechaColumna]).getTime() : 0;
        return dateB - dateA; 
      });
    }

    const mapa = new Map();
    data.forEach(row => {
      idColumnas.forEach(colId => {
        let valor = row[colId];
        
        if (nombreHoja === "Revision duplicados") {
            valor = limpiarPoliza(valor); 
        } else {
            valor = limpiarValorCondicional(valor);
        }
        
        if (valor && !mapa.has(valor)) {
          mapa.set(valor, row);
        }
      });
    });

    return { map: mapa, data: data };
  }


  const { map: duplicadosMap } = cargarDatosYMapa("Revision duplicados", [12], null);
  const esPolizaDuplicada = (poliza) => poliza && duplicadosMap.has(limpiarPoliza(poliza));

  const config = {
    leadsSalud: { name: "TOTAL LEADS SALUD LIGERO", ids: [6, 4], fecha: 12, infoCols: { fuente: 9, medio: 10, campaña: 11 } }, 
    leads322: { name: "Leads 322 - salud ", ids: [11], fecha: 27, infoCols: { medio: 8, campaña: 24 }, fuente: "322" }, 
    referidos: { name: "Referidos Salud a su medida", ids: [0, 2], fecha: 6, infoCols: { medio: 7 }, fuente: "Referido" }, 
    bases: { name: "BASES SALUD A SU MEDIDA", ids: [0, 1], fecha: 2, infoCols: { fuente: 7, medio: 3, campaña: 6 } } 
  };

  const { map: leadsSaludMap } = cargarDatosYMapa(config.leadsSalud.name, config.leadsSalud.ids, config.leadsSalud.fecha);
  const { map: leads322Map } = cargarDatosYMapa(config.leads322.name, config.leads322.ids, config.leads322.fecha);
  const { map: referidosMap } = cargarDatosYMapa(config.referidos.name, config.referidos.ids, config.referidos.fecha);
  const { map: basesMap } = cargarDatosYMapa(config.bases.name, config.bases.ids, config.bases.fecha);

  const encabezados = [
    "cc", "cc2", "correo", "S", "T",
    "322", "Referidos", "Bases", "Base mail", 
    "ventas", "test", "fuente", "medio", "campaña", "fecha lead"
  ];
  const colInicioEscritura = 16; 
  hojaPrincipal.getRange(1, colInicioEscritura, 1, encabezados.length).setValues([encabezados]);

  const resultados = [];

  rangoDatos.forEach(fila => {
    const poliza = limpiarPoliza(fila[0]); 
    
    const correo = limpiarCorreo(fila[6]); 
    const cc1 = limpiarCC(fila[7]);     
    const cc2 = limpiarCC(fila[9]);     

    let testValue = "-";
    let skipLeadsSearch = false;

    if (esPolizaDuplicada(poliza)) {
      testValue = "DUPLICADO";
      skipLeadsSearch = true;
    }
    
    if (skipLeadsSearch) {
      resultados.push([0, 0, 0, "", "", 0, 0, 0, 0, 0, testValue, "", "", "", "" ]);
      return; 
    }


    const matchSaludCC1   = (cc1 && leadsSaludMap.has(cc1)) ? 1 : 0;
    const matchSaludCC2   = (cc2 && leadsSaludMap.has(cc2)) ? 1 : 0;
    const matchSaludCorreo = (correo && leadsSaludMap.has(correo)) ? 1 : 0;
    
    const match322CC     = (cc1 && leads322Map.has(cc1)) ? 1 : 0; 
    const matchReferidos = (cc1 && referidosMap.has(cc1)) ? 1 : 0;
    const matchBaseCC     = (cc1 && basesMap.has(cc1)) ? 1 : 0;
    const matchBaseMail   = (correo && basesMap.has(correo)) ? 1 : 0;

    const ventas = matchSaludCC1 + matchSaludCC2 + matchSaludCorreo +
                   match322CC + matchReferidos + matchBaseCC + matchBaseMail;

    let fuente = "", medio = "", campana = "", fechaLead = null;
    let registro = null;


    if (matchSaludCC1 || matchSaludCC2 || matchSaludCorreo) {
      registro = leadsSaludMap.get(cc1) || leadsSaludMap.get(cc2) || leadsSaludMap.get(correo);
      if (registro) {
        fuente = registro[config.leadsSalud.infoCols.fuente];
        medio = registro[config.leadsSalud.infoCols.medio];
        campana = registro[config.leadsSalud.infoCols.campaña];
        fechaLead = registro[config.leadsSalud.fecha];
      }
    } 
    else if (match322CC) {
      registro = leads322Map.get(cc1);
      if (registro) {
        fuente = config.leads322.fuente; 
        medio = registro[config.leads322.infoCols.medio];
        campana = registro[config.leads322.infoCols.campaña];
        fechaLead = registro[config.leads322.fecha];
      }
    } 
    else if (matchReferidos) {
      registro = referidosMap.get(cc1);
      if (registro) {
        fuente = config.referidos.fuente; 
        medio = registro[config.referidos.infoCols.medio];
        campana = ""; 
        fechaLead = registro[config.referidos.fecha];
      }
    } 
    else if (matchBaseCC || matchBaseMail) {
      registro = basesMap.get(cc1) || basesMap.get(correo);
      if (registro) {
        fuente = registro[config.bases.infoCols.fuente];
        medio = registro[config.bases.infoCols.medio];
        campana = registro[config.bases.infoCols.campaña];
        fechaLead = registro[config.bases.fecha];
      }
    }
    
    const fechaFormateada = formatearFecha(fechaLead);

    resultados.push([
      matchSaludCC1,     
      matchSaludCC2,      
      matchSaludCorreo,   
      "", "",             
      match322CC,         
      matchReferidos,     
      matchBaseCC,        
      matchBaseMail,      
      ventas,             
      testValue,          
      fuente,             
      medio,              
      campana,           
      fechaFormateada    
    ]);
  });

  if (resultados.length > 0) {
    hojaPrincipal.getRange(2, colInicioEscritura, resultados.length, resultados[0].length).setValues(resultados);
  }
}
