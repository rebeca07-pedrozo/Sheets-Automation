
function EmisionesAutosCruzados(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    ui.alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const limpiarDoc = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').toLowerCase();
  const limpiarPlaca = p => String(p || '').replace(/\s/g, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();

  const datosEmisiones = hoja.getRange(2, 2, ultimaFila - 1, 14).getValues();

  const hojaLeads = ss.getSheetByName("TOTAL LEADS - CRUZAR");
  if (!hojaLeads) {
      ui.alert("Error: No se encontró la hoja 'TOTAL LEADS - CRUZAR'.");
      return;
  }
  const ultimaFilaLeads = hojaLeads.getLastRow();
  const ultimaColumnaLeads = hojaLeads.getLastColumn();
  
  const totalLeadsCruzarDatos = (ultimaFilaLeads > 1 && ultimaColumnaLeads > 0) 
      ? hojaLeads.getRange(2, 1, ultimaFilaLeads - 1, ultimaColumnaLeads).getValues() 
      : [];

  const mapTLC_CC = new Map(), mapTLC_Placa = new Map(), mapTLC_Correo = new Map();

  totalLeadsCruzarDatos.forEach(r => {
    const cedula = r[1] ? limpiarDoc(r[1]) : '';
    const correo = r[2] ? limpiarCorreo(r[2]) : '';
    const placa = r[4] ? limpiarPlaca(r[4]) : '';

    if (cedula && !mapTLC_CC.has(cedula)) mapTLC_CC.set(cedula, r);
    if (correo && !mapTLC_Correo.has(correo)) mapTLC_Correo.set(correo, r);
    if (placa && !mapTLC_Placa.has(placa)) mapTLC_Placa.set(placa, r);
  });

  const basesDatos = ss.getSheetByName("BASES").getDataRange().getValues().slice(1);

  basesDatos.sort((a, b) => {
    const fechaA = new Date(a[2]);
    const fechaB = new Date(b[2]);
    if (isNaN(fechaB)) return -1;
    if (isNaN(fechaA)) return 1;
    return fechaB - fechaA;
  });

  const mapBASES_CC = new Map(), mapBASES_Correo = new Map();

  basesDatos.forEach(r => {
    const cedula = r[0] ? limpiarDoc(r[0]) : '';
    const correo = r[1] ? limpiarCorreo(r[1]) : '';
    if (cedula && !mapBASES_CC.has(cedula)) mapBASES_CC.set(cedula, r);
    if (correo && !mapBASES_Correo.has(correo)) mapBASES_Correo.set(correo, r);
  });

  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    const timeZone = ss.getSpreadsheetTimeZone() || 'GMT-5';
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd HH:mm:ss");
  }

  const resultadosFinales = datosEmisiones.map(r => {
    const doc = limpiarDoc(r[10]); 
    const placa = limpiarPlaca(r[0]); 
    const correo = limpiarCorreo(r[13]); 

    let countTLC_CC = 0, countTLC_Placa = 0, countTLC_Correo = 0;
    let countBASES_CC = 0, countBASES_Correo = 0;
    let fuenteFinal = '', medioFinal = '', campañaFinal = '', adnameFinal = '', fechaFinal = '';

    let foundRowTLC = null;
    let foundRowBASES = null;

    if (doc && mapTLC_CC.has(doc)) {
      foundRowTLC = mapTLC_CC.get(doc);
    }
    if (placa && !foundRowTLC && mapTLC_Placa.has(placa)) {
      foundRowTLC = mapTLC_Placa.get(placa);
    }
    if (correo && !foundRowTLC && mapTLC_Correo.has(correo)) {
      foundRowTLC = mapTLC_Correo.get(correo);
    }

    if (doc && mapBASES_CC.has(doc)) {
      foundRowBASES = mapBASES_CC.get(doc);
    }
    if (correo && !foundRowBASES && mapBASES_Correo.has(correo)) {
      foundRowBASES = mapBASES_Correo.get(correo);
    }

    if(foundRowTLC){
        if (doc && mapTLC_CC.get(doc) === foundRowTLC) countTLC_CC = 1;
        if (placa && mapTLC_Placa.get(placa) === foundRowTLC) countTLC_Placa = 1;
        if (correo && mapTLC_Correo.get(correo) === foundRowTLC) countTLC_Correo = 1;
    }
    if(foundRowBASES){
        if (doc && mapBASES_CC.get(doc) === foundRowBASES) countBASES_CC = 1;
        if (correo && mapBASES_Correo.get(correo) === foundRowBASES) countBASES_Correo = 1;
    }

    if (foundRowTLC) {
      fuenteFinal = foundRowTLC[8] || '';
      medioFinal = foundRowTLC[12] || '';
      campañaFinal = foundRowTLC[9] || '';
      adnameFinal = foundRowTLC[13] || ''; 
      fechaFinal = formatearFecha(foundRowTLC[11]);
    } else if (foundRowBASES) {
      fuenteFinal = foundRowBASES[7] || 'BASES';
      medioFinal = foundRowBASES[3] || '';
      campañaFinal = foundRowBASES[6] || '';
      adnameFinal = ''; 
      fechaFinal = formatearFecha(foundRowBASES[2]);
    }

    const totalConteo = countTLC_CC + countTLC_Placa + countTLC_Correo + countBASES_CC + countBASES_Correo;
    const conteos = [countTLC_CC, countTLC_Placa, countTLC_Correo, countBASES_CC, countBASES_Correo];

    return [
      ...conteos,
      totalConteo, 
      totalConteo, 
      fuenteFinal,
      medioFinal,
      campañaFinal,
      adnameFinal, 
      fechaFinal,
    ];
  });

  const nuevosEncabezados = [
    "TOTAL LEADS CC", "TOTAL LEADS Placa", "TOTAL LEADS Correo",
    "Bases CC", "Bases Correo", "Total Leads", "Ventas",
    "Fuente", "Medio", "Campaña", "Adname", "Fecha Lead" 
  ];

  if (resultadosFinales.length > 0) {
    const columnaInicioResultados = 17;
    hoja.getRange(1, columnaInicioResultados, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hoja.getRange(2, columnaInicioResultados, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }

  
}




function cruzarLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); 

  const NOMBRE_HOJA_CRUZAR = 'TOTAL LEADS - CRUZAR';
  const NOMBRE_HOJA_ACUMULADO = 'Acumulado leads fb';
  
  const hojaCruzar = ss.getSheetByName(NOMBRE_HOJA_CRUZAR);
  const hojaAcumulado = ss.getSheetByName(NOMBRE_HOJA_ACUMULADO);

  if (!hojaCruzar || !hojaAcumulado) {
    const missingSheet = !hojaCruzar ? NOMBRE_HOJA_CRUZAR : NOMBRE_HOJA_ACUMULADO;
    Logger.log(`Error: No se encontró la hoja '${missingSheet}'.`);
    ui.alert(`Error: Asegúrate de que los nombres de las hojas '${NOMBRE_HOJA_CRUZAR}' y '${NOMBRE_HOJA_ACUMULADO}' sean correctos.`);
    return;
  }

  const COL_CC_CRUZAR = 1;      
  const COL_CORREO_CRUZAR = 2;  
  const COL_UTM_SOURCE = 8;     
  const COL_DESTINO = 13;       
  
  const COL_CC_ACUMULADO = 1;   
  const COL_CORREO_ACUMULADO = 2; 
  const COL_DATO_O = 14;          

  const datosCruzar = hojaCruzar.getDataRange().getValues();
  const datosAcumulado = hojaAcumulado.getDataRange().getValues();

  const limpiarDato = d => String(d || '').trim();

  const mapaAcumulado = new Map();
  for (let i = 1; i < datosAcumulado.length; i++) {
    const fila = datosAcumulado[i];
    const cc = limpiarDato(fila[COL_CC_ACUMULADO]);
    const correo = limpiarDato(fila[COL_CORREO_ACUMULADO]);
    const datoO = fila[COL_DATO_O];
    
    if (cc) {
      if (!mapaAcumulado.has(`CC_${cc}`)) {
        mapaAcumulado.set(`CC_${cc}`, datoO);
      }
    }
    if (correo) {
      if (!mapaAcumulado.has(`CORREO_${correo}`)) {
        mapaAcumulado.set(`CORREO_${correo}`, datoO);
      }
    }
  }

  const resultados = [];

  for (let i = 1; i < datosCruzar.length; i++) {
    const fila = datosCruzar[i];
    const utmSource = limpiarDato(fila[COL_UTM_SOURCE]).toLowerCase();
    let valorEncontrado = '';

    if (utmSource === 'ig' || utmSource === 'fb') {
      const cc = limpiarDato(fila[COL_CC_CRUZAR]);
      const correo = limpiarDato(fila[COL_CORREO_CRUZAR]);

      if (cc && mapaAcumulado.has(`CC_${cc}`)) {
        valorEncontrado = mapaAcumulado.get(`CC_${cc}`);
      } else if (correo && mapaAcumulado.has(`CORREO_${correo}`)) {
        valorEncontrado = mapaAcumulado.get(`CORREO_${correo}`);
      } else {
        valorEncontrado = ' ';
      }
    }
    
    resultados.push([valorEncontrado]);
  }

  if (resultados.length > 0) {
    hojaCruzar.getRange(2, COL_DESTINO + 1, resultados.length, 1).setValues(resultados);
    Logger.log('Proceso de cruce de leads completado con éxito.');
  }
}