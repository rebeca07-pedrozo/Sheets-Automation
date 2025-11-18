
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

  const hojaBases = ss.getSheetByName("BASES");
  if (!hojaBases) {
    ui.alert("Error: No se encontró la hoja 'BASES'.");
    return;
  }
  const basesDatos = hojaBases.getDataRange().getValues().slice(1);

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
    const placa = limpiarPlaca(r[0]);
    const doc = limpiarDoc(r[10]);
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
    
    if (!foundRowTLC) {
      if (doc && mapBASES_CC.has(doc)) {
        foundRowBASES = mapBASES_CC.get(doc);
      }
      if (correo && !foundRowBASES && mapBASES_Correo.has(correo)) {
        foundRowBASES = mapBASES_Correo.get(correo);
      }
    }
    
    if (doc && mapTLC_CC.has(doc)) countTLC_CC = 1;
    if (placa && mapTLC_Placa.has(placa)) countTLC_Placa = 1;
    if (correo && mapTLC_Correo.has(correo)) countTLC_Correo = 1;

    if (doc && mapBASES_CC.has(doc)) countBASES_CC = 1;
    if (correo && mapBASES_Correo.has(correo)) countBASES_Correo = 1;

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


    const conteos = [countTLC_CC, countTLC_Placa, countTLC_Correo, countBASES_CC, countBASES_Correo];
    const totalConteo = conteos.reduce((sum, current) => sum + current, 0);

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
    "Bases CC", "Bases Correo",
    "Total Leads", "Ventas", 
    "Fuente", "Medio", "Campaña", "Adname", "Fecha Lead" 
  ];

  if (resultadosFinales.length > 0) {
    const columnaInicioResultados = 17; 
    hoja.getRange(1, columnaInicioResultados, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hoja.getRange(2, columnaInicioResultados, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }
}