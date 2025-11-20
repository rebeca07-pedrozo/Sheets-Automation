function CruceDatosSaludIntegral(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPrincipal = ss.getSheetByName(nombreHoja);
  const COLUMNA_INICIO_RESULTADOS = 15;

  if (!hojaPrincipal) {
    SpreadsheetApp.getUi().alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }

  const ultimaFila = hojaPrincipal.getLastRow();
  if (ultimaFila < 2) {
    Logger.log("No hay datos para procesar en la hoja principal.");
    return;
  }

  const limpiarCC = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();

  const COL_CORREO_PRINCIPAL = 9;
  const COL_CC1_PRINCIPAL = 10;
  const COL_CC2_PRINCIPAL = 12;
  const NUM_COLUMNAS_PRINCIPALES = 13;

  const datosPrincipal = hojaPrincipal.getRange(2, 1, ultimaFila - 1, NUM_COLUMNAS_PRINCIPALES).getValues();

  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    const timeZone = ss.getSpreadsheetTimeZone() || 'GMT-5';
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd HH:mm:ss");
  }


  const LEADS_TOTAL_INTEGRAL = ss.getSheetByName("LEADS TOTAL INTEGRAL");
  const map_LEADS_INTEGRAL_CC = new Map();
  const map_LEADS_INTEGRAL_Correo = new Map();

  if (LEADS_TOTAL_INTEGRAL) {
    LEADS_TOTAL_INTEGRAL.getDataRange().getValues().slice(1).forEach(r => {
      const cc = r[2] ? limpiarCC(r[2]) : '';
      const correo = r[0] ? limpiarCorreo(r[0]) : '';
      if (cc && !map_LEADS_INTEGRAL_CC.has(cc)) map_LEADS_INTEGRAL_CC.set(cc, r);
      if (correo && !map_LEADS_INTEGRAL_Correo.has(correo)) map_LEADS_INTEGRAL_Correo.set(correo, r);
    });
  }

  const LEADS_322 = ss.getSheetByName("Leads 322");
  const map_LEADS_322_CC = new Map();

  if (LEADS_322) {
    LEADS_322.getDataRange().getValues().slice(1).forEach(r => {
      const cc = r[11] ? limpiarCC(r[11]) : '';
      if (cc && !map_LEADS_322_CC.has(cc)) map_LEADS_322_CC.set(cc, r);
    });
  }

  const REFERIDOS = ss.getSheetByName("Leads Referidos");
  const map_REFERIDOS_CC = new Map();
  const map_REFERIDOS_Correo = new Map();

  if (REFERIDOS) {
    REFERIDOS.getDataRange().getValues().slice(1).forEach(r => {
      const cc = r[0] ? limpiarCC(r[0]) : '';
      const correo = r[2] ? limpiarCorreo(r[2]) : '';
      if (cc && !map_REFERIDOS_CC.has(cc)) map_REFERIDOS_CC.set(cc, r);
      if (correo && !map_REFERIDOS_Correo.has(correo)) map_REFERIDOS_Correo.set(correo, r);
    });
  }

  const NOMBRE_HOJA_BASES = "BASES INTEGRAL";
  const BASES = ss.getSheetByName(NOMBRE_HOJA_BASES);
  const basesDatos = BASES ? BASES.getDataRange().getValues().slice(1) : [];


  basesDatos.sort((a, b) => {
    const fechaA = new Date(a[2]);
    const fechaB = new Date(b[2]);
    if (isNaN(fechaB)) return -1;
    if (isNaN(fechaA)) return 1;
    return fechaB.getTime() - fechaA.getTime();
  });

  const map_BASES_CC = new Map();
  const map_BASES_Correo = new Map();

  basesDatos.forEach(r => {
    const cc = r[0] ? limpiarCC(r[0]) : '';
    const correo = r[1] ? limpiarCorreo(r[1]) : '';
    if (cc && !map_BASES_CC.has(cc)) map_BASES_CC.set(cc, r);
    if (correo && !map_BASES_Correo.has(correo)) map_BASES_Correo.set(correo, r);
  });


  const resultadosFinales = datosPrincipal.map(r => {
    const correo = limpiarCorreo(r[COL_CORREO_PRINCIPAL]);
    const cc1 = limpiarCC(r[COL_CC1_PRINCIPAL]);
    const cc2 = limpiarCC(r[COL_CC2_PRINCIPAL]);

    let count_LTI_CC1 = 0;
    let count_LTI_CC2 = 0;
    let count_LTI_Correo = 0;
    let count_322_CC1 = 0;
    let count_322_CC2 = 0;
    let count_BASES_CC1 = 0;
    let count_BASES_Correo = 0;
    let count_BASES_CC2 = 0;
    let count_322_Otros = 0;
    let count_Referidos = 0;

    let fuenteFinal = '', medioFinal = '', campañaFinal = '', fechaFinal = '';
    let foundRow = null;
    let foundSheet = null; 

    if (cc1 && map_LEADS_INTEGRAL_CC.has(cc1)) {
      foundRow = map_LEADS_INTEGRAL_CC.get(cc1);
      foundSheet = 'LTI';
    } else if (cc2 && map_LEADS_INTEGRAL_CC.has(cc2)) {
      foundRow = map_LEADS_INTEGRAL_CC.get(cc2);
      foundSheet = 'LTI';
    } else if (correo && map_LEADS_INTEGRAL_Correo.has(correo)) {
      foundRow = map_LEADS_INTEGRAL_Correo.get(correo);
      foundSheet = 'LTI';
    }

    if (cc1 && map_LEADS_INTEGRAL_CC.has(cc1)) count_LTI_CC1 = 1;
    if (cc2 && map_LEADS_INTEGRAL_CC.has(cc2)) count_LTI_CC2 = 1;
    if (correo && map_LEADS_INTEGRAL_Correo.has(correo)) count_LTI_Correo = 1;
    
    if (!foundRow && cc1 && map_LEADS_322_CC.has(cc1)) {
      foundRow = map_LEADS_322_CC.get(cc1);
      foundSheet = '322';
    } else if (!foundRow && cc2 && map_LEADS_322_CC.has(cc2)) {
      foundRow = map_LEADS_322_CC.get(cc2);
      foundSheet = '322';
    }

    if (cc1 && map_LEADS_322_CC.has(cc1)) count_322_CC1 = 1;
    if (cc2 && map_LEADS_322_CC.has(cc2)) count_322_CC2 = 1;

    if (!foundRow && cc1 && map_BASES_CC.has(cc1)) {
      foundRow = map_BASES_CC.get(cc1);
      foundSheet = 'BASES';
    } else if (!foundRow && correo && map_BASES_Correo.has(correo)) {
      foundRow = map_BASES_Correo.get(correo);
      foundSheet = 'BASES';
    } else if (!foundRow && cc2 && map_BASES_CC.has(cc2)) {
      foundRow = map_BASES_CC.get(cc2);
      foundSheet = 'BASES';
    }

    if (cc1 && map_BASES_CC.has(cc1)) count_BASES_CC1 = 1;
    if (correo && map_BASES_Correo.has(correo)) count_BASES_Correo = 1;
    if (cc2 && map_BASES_CC.has(cc2)) count_BASES_CC2 = 1;
    
    if (!foundRow && ((cc1 && map_REFERIDOS_CC.has(cc1)) || (cc2 && map_REFERIDOS_CC.has(cc2)) || (correo && map_REFERIDOS_Correo.has(correo)))) {
      foundRow = (cc1 && map_REFERIDOS_CC.get(cc1)) || (cc2 && map_REFERIDOS_CC.get(cc2)) || (correo && map_REFERIDOS_Correo.get(correo));
      foundSheet = 'REFERIDOS';
    }

    if ((cc1 && map_REFERIDOS_CC.has(cc1)) || (cc2 && map_REFERIDOS_CC.has(cc2)) || (correo && map_REFERIDOS_Correo.has(correo))) {
      count_Referidos = 1;
    }

    if (cc1 && map_LEADS_322_CC.has(cc1) && count_322_CC1 === 0) count_322_Otros = 1;
    if (cc2 && map_LEADS_322_CC.has(cc2) && count_322_CC2 === 0) count_322_Otros = 1;


    if (foundRow) {
      if (foundSheet === 'LTI' && foundRow.length > 7) {
        fuenteFinal = foundRow[4] || '';
        medioFinal = foundRow[5] || '';
        campañaFinal = foundRow[6] || '';
        fechaFinal = formatearFecha(foundRow[8]);
      } 
      else if (foundSheet === '322' && foundRow.length > 27) {
        fuenteFinal = "322"; 
        medioFinal = foundRow[8] || '';
        campañaFinal = foundRow[24] || '';
        fechaFinal = formatearFecha(foundRow[0]); 
      } 
      else if (foundSheet === 'BASES' && foundRow.length > 6) {
        fuenteFinal = foundRow[8] || 'BASES';
        medioFinal = foundRow[7] || '';
        campañaFinal = foundRow[6] || '';
        fechaFinal = formatearFecha(foundRow[2]);
      } else if (foundSheet === 'REFERIDOS' && foundRow.length > 2) {
        fuenteFinal = "Referidos";
        medioFinal = '';
        campañaFinal = '';
        fechaFinal = formatearFecha(foundRow[1]);
      }
    }

    const conteosNumericos = [
      count_LTI_CC1, count_LTI_CC2, count_LTI_Correo,
      count_322_CC1, count_322_CC2,
      count_BASES_CC1, count_BASES_Correo, count_BASES_CC2,
      count_322_Otros,
      count_Referidos,
    ];
    
    const totalVentas = conteosNumericos.reduce((a, b) => a + b, 0);

    const conteosString = conteosNumericos.map(String);
    
    return [
      ...conteosString, 
      String(totalVentas), 
      fuenteFinal,
      medioFinal,
      campañaFinal,
      fechaFinal,
    ];
  });

  const nuevosEncabezados = [
    "cc - TOTAL LEADS SALUD LIGERO", "cc2 - TOTAL LEADS SALUD LIGERO", "correo - TOTAL LEADS SALUD LIGERO",
    "322 CC1 - Leads 322", "322 CC2 - Leads 322",
    "Base CC1 - BASES INTEGRAL", "Base Mail - BASES INTEGRAL", "Base CC2 - BASES INTEGRAL",
    "322 otros - Leads 322",
    "Referidos - Leads Referidos",
    "ventas",
    "fuente",
    "medio",
    "campaña",
    "fecha lead",
  ];

  if (resultadosFinales.length > 0) {
    hojaPrincipal.getRange(1, COLUMNA_INICIO_RESULTADOS, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hojaPrincipal.getRange(2, COLUMNA_INICIO_RESULTADOS, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }
}

//Aca empieza medida
/***Poliza : E
 * Correo: K
 * CC1: L 
 * CC2: N 
***/
function SaludAsuMedida(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPrincipal = ss.getSheetByName(nombreHoja);

  if (!hojaPrincipal) {
    SpreadsheetApp.getUi().alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }

  const ultimaFila = hojaPrincipal.getLastRow();
  if (ultimaFila < 2) {
    Logger.log("No hay datos para procesar en la hoja principal.");
    return;
  }

  const numColumnas = 10;
  const rangoDatos = hojaPrincipal.getRange(2, 5, ultimaFila - 1, numColumnas).getValues(); 

  const limpiarCC = d => String(d || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
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

    const data = hoja.getDataRange().getValues().slice(1);
    
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
    "cc - LIGERO", "cc2 - LIGERO", "correo - LIGERO", "S", "T",
    "322", "Referidos", "Base CC", "Base mail", 
    "ventas", "test", "fuente", "medio", "campaña", "fecha lead"
  ];
  const colInicioEscritura = 16; 
  hojaPrincipal.getRange(1, colInicioEscritura, 1, encabezados.length).setValues([encabezados]);

  const resultados = [];

  rangoDatos.forEach(fila => {
    const poliza = limpiarPoliza(fila[0]); // E
    const correo = limpiarCorreo(fila[6]); // K
    const cc1 = limpiarCC(fila[7]);       // L
    const cc2 = limpiarCC(fila[9]);       // N

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
    
    const match322CC     = ((cc1 && leads322Map.has(cc1)) || (cc2 && leads322Map.has(cc2))) ? 1 : 0; 
    
    const matchReferidos = ((cc1 && referidosMap.has(cc1)) || (cc2 && referidosMap.has(cc2)) || (correo && referidosMap.has(correo))) ? 1 : 0;

    const matchBaseCC     = ((cc1 && basesMap.has(cc1)) || (cc2 && basesMap.has(cc2))) ? 1 : 0;
    const matchBaseMail   = (correo && basesMap.has(correo)) ? 1 : 0;

    const ventas = matchSaludCC1 + matchSaludCC2 + matchSaludCorreo +
                     match322CC + matchReferidos + matchBaseCC + matchBaseMail;

    let fuente = "", medio = "", campana = "", fechaLead = null;
    let registro = null;


    if (matchSaludCC1 || matchSaludCC2 || matchSaludCorreo) {
        if (matchSaludCC1) {
            registro = leadsSaludMap.get(cc1);
        } else if (matchSaludCC2) {
            registro = leadsSaludMap.get(cc2);
        } else if (matchSaludCorreo) {
            registro = leadsSaludMap.get(correo);
        }
        
        if (registro) {
            fuente = registro[config.leadsSalud.infoCols.fuente] || '';
            medio = registro[config.leadsSalud.infoCols.medio] || '';
            campana = registro[config.leadsSalud.infoCols.campaña] || '';
            fechaLead = registro[config.leadsSalud.fecha];
        }
    } 
    else if (match322CC) {
        registro = leads322Map.get(cc1) || leads322Map.get(cc2);
        if (registro) {
            fuente = config.leads322.fuente; 
            medio = registro[config.leads322.infoCols.medio] || '';
            campana = registro[config.leads322.infoCols.campaña] || '';
            fechaLead = registro[config.leads322.fecha];
        }
    } 
    else if (matchReferidos) {
        registro = referidosMap.get(cc1) || referidosMap.get(cc2) || referidosMap.get(correo);
        if (registro) {
            fuente = config.referidos.fuente; 
            medio = registro[config.referidos.infoCols.medio] || '';
            campana = ""; 
            fechaLead = registro[config.referidos.fecha];
        }
    } 
    else if (matchBaseCC || matchBaseMail) {
        registro = basesMap.get(cc1) || basesMap.get(correo) || basesMap.get(cc2);
        if (registro) {
            fuente = registro[config.bases.infoCols.fuente] || '';
            medio = registro[config.bases.infoCols.medio] || '';
            campana = registro[config.bases.infoCols.campaña] || '';
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