function SaludPREPERFILADOR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Copia de Emisiones 31 ago"); 
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja de destino.');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const config = {
    hojaDestino: {
      ccCol: 11,      
      correoCol: 10,  
      cc2Col: 13,     
      resultadosCol: 15,
      extrasCol: 19      
    },
    leads: {
      sheet: "Leads",
      ccCol: 1,        
      correoCol: 2,   
      productoCol: 4,
      medioCol: 5,     
      fuenteCol: 7,   
      campañaCol: 8,   
      fechaCol: 9     
    }
  };

  const cedulas = hoja.getRange(2, config.hojaDestino.ccCol, ultimaFila - 1, 1).getValues()
                      .map(r => r[0] ? String(r[0]).trim() : "");
  const correos = hoja.getRange(2, config.hojaDestino.correoCol, ultimaFila - 1, 1).getValues()
                      .map(r => r[0] ? String(r[0]).trim().toLowerCase() : "");
  const cedulas2 = hoja.getRange(2, config.hojaDestino.cc2Col, ultimaFila - 1, 1).getValues()
                       .map(r => r[0] ? String(r[0]).trim() : "");

  const leadsSheet = ss.getSheetByName(config.leads.sheet);
  if (!leadsSheet) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja Leads.');
    return;
  } 
  const leadsData = leadsSheet.getDataRange().getValues();
  leadsData.shift(); 

  const leadsCedulaMap = new Map();
  const leadsEmailMap = new Map();

  leadsData.forEach(row => {
    const cc = row[config.leads.ccCol] ? String(row[config.leads.ccCol]).trim() : "";
    const email = row[config.leads.correoCol] ? String(row[config.leads.correoCol]).trim().toLowerCase() : "";
    if (cc) leadsCedulaMap.set(cc, row);
    if (email) leadsEmailMap.set(email, row);
  });

  function formatearFecha(fecha) {
    if (!fecha) return "";
    const d = new Date(fecha);
    if (isNaN(d)) return "";
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }

  const resultados = [];

  cedulas.forEach((cc, i) => {
    const correo = correos[i];
    const cc2 = cedulas2[i];

    let matchCC = cc && leadsCedulaMap.has(cc) ? 1 : 0;
    let matchCorreo = correo && leadsEmailMap.has(correo) ? 1 : 0;
    let matchCC2 = cc2 && leadsCedulaMap.has(cc2) ? 1 : 0;

    let ventas = matchCC + matchCorreo + matchCC2;

    let fuente = "", medio = "", campaña = "", fechaLead = "", producto = "";

    let foundRow = null;
    if (matchCC) {
      foundRow = leadsCedulaMap.get(cc);
    } else if (matchCorreo) {
      foundRow = leadsEmailMap.get(correo);
    } else if (matchCC2) {
      foundRow = leadsCedulaMap.get(cc2);
    }

    if (foundRow) {
      fuente    = foundRow[config.leads.fuenteCol]   ? String(foundRow[config.leads.fuenteCol])   : "";
      medio     = foundRow[config.leads.medioCol]    ? String(foundRow[config.leads.medioCol])    : "";
      campaña   = foundRow[config.leads.campañaCol]  ? String(foundRow[config.leads.campañaCol])  : "";
      fechaLead = formatearFecha(foundRow[config.leads.fechaCol]);
      producto  = foundRow[config.leads.productoCol] ? String(foundRow[config.leads.productoCol]) : "";
    }

    resultados.push([matchCC, matchCorreo, matchCC2, ventas, fuente, medio, campaña, fechaLead, producto]);
  });

  hoja.getRange(2, config.hojaDestino.resultadosCol, resultados.length, 4)
      .setValues(resultados.map(r => r.slice(0, 4)))
      .setHorizontalAlignment("right");

  hoja.getRange(2, config.hojaDestino.extrasCol, resultados.length, 5)
      .setValues(resultados.map(r => r.slice(4)))
      .setHorizontalAlignment("right");

  hoja.getRange(1, config.hojaDestino.resultadosCol, 1, 4)
      .setValues([["cc", "correo", "cc2", "ventas"]])
      .setHorizontalAlignment("right");
  hoja.getRange(1, config.hojaDestino.extrasCol, 1, 5)
      .setValues([["fuente", "Medio", "campaña", "fecha lead", "Producto lead"]])
      .setHorizontalAlignment("right");
}
