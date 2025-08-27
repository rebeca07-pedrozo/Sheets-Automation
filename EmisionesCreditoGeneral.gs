function EmisionesCreditoPractica_Todas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Sheet18");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;
  const encabezados = [
    "venta",         
    "fuente",        
    "med",           
    "campaña",       
    "fecha lead",    
    "prod",          
    "cruce cami",    
    "Prioridad",     
    "Base",          
    "Cruce Email"    
  ];
  hoja.getRange(1, 19, 1, encabezados.length).setValues([encabezados]).setHorizontalAlignment("right");
  hoja.getRange(1, 19, 1, encabezados.length).setValues([encabezados]).setHorizontalAlignment("right");

  hoja.getRange(2, 19, ultimaFila - 1, encabezados.length).clearContent();

  const cedulas = hoja.getRange(2, 9, ultimaFila - 1, 1).getValues().map(r => r[0] ? String(r[0]).trim() : "");
  const correos = hoja.getRange(2, 11, ultimaFila - 1, 1).getValues().map(r => r[0] ? String(r[0]).trim().toLowerCase() : "");

  const totalLeadsSheet = ss.getSheetByName("TOTAL LEADS");
  const totalLeadsData = totalLeadsSheet.getRange(2, 1, totalLeadsSheet.getLastRow() - 1, 14).getValues();
  const totalLeadsMap = new Map();
  totalLeadsData.forEach(row => {
    const cedula = row[4] ? String(row[4]).trim() : "";
    const email = row[2] ? String(row[2]).trim().toLowerCase() : "";
    if (cedula) totalLeadsMap.set(cedula, row);
    if (email) totalLeadsMap.set(email, row);
  });
  const leads322Sheet = ss.getSheetByName("Leads 322");
  const leads322Data = leads322Sheet.getRange(2, 1, leads322Sheet.getLastRow() - 1, 12).getValues();
  const leads322Map = new Map();
  leads322Data.forEach(row => {
    const cedula = row[11] ? String(row[11]).trim() : "";
    if (cedula) leads322Map.set(cedula, row);
  });
const basesSheet = ss.getSheetByName("BASES");
  const basesData = basesSheet.getRange(2, 1, basesSheet.getLastRow() - 1, 10).getValues();
  const basesMap = new Map();
  basesData.forEach(row => {
    const cedula = row[0] ? String(row[0]).trim() : "";
    const email = row[1] ? String(row[1]).trim().toLowerCase() : "";
    if (cedula) basesMap.set(cedula, row);
    if (email) basesMap.set(email, row);
  });
  function formatearFecha(fecha) {
    if (!fecha) return "";
    const d = new Date(fecha);
    if (isNaN(d)) return "";
    const dia = ("0" + d.getDate()).slice(-2);
    const mes = ("0" + (d.getMonth() + 1)).slice(-2);
    const anio = d.getFullYear();
    return `${dia}/${mes}/${anio}`;
  }
const resultados = [];
  const valoresW = [];

  cedulas.forEach((cedulaRaw, i) => {
    const cedula = cedulaRaw;
    const correo = correos[i];

    const countN = (cedula && totalLeadsMap.has(cedula)) ? 1 : 0;
    const countO = (correo && totalLeadsMap.has(correo)) ? 1 : 0;
    const countP = (cedula && leads322Map.has(cedula)) ? 1 : 0;
    const countQ = (cedula && basesMap.has(cedula)) ? 1 : 0;
    const countR = (correo && basesMap.has(correo)) ? 1 : 0;

    const suma = countN + countO + countP + countQ + countR;

    let fuente = "";
    let additionalData = Array(9).fill(""); 
    let valorW = "";

    if (suma > 0) {
      if (countN === 1 || countO === 1) {
        fuente = "TOTAL LEADS";
        const foundRow = totalLeadsMap.has(cedula) ? totalLeadsMap.get(cedula) : totalLeadsMap.get(correo);
        if (foundRow) {
          
          additionalData = foundRow.slice(9, 13).map(c => c != null ? String(c) : "").concat(Array(5).fill(""));
          valorW = formatearFecha(foundRow[13]); 
        }
      } else if (countP === 1) {
        fuente = "Leads 322";
        const foundRow = leads322Map.get(cedula);
        if (foundRow) {
          additionalData = Array(9).fill("");
          valorW = formatearFecha(foundRow[0]); 
        }
      } else if (countQ === 1 || countR === 1) {
        fuente = "BASES";
        const foundRow = basesMap.has(cedula) ? basesMap.get(cedula) : basesMap.get(correo);
        if (foundRow) {
          valorW = formatearFecha(foundRow[2]); 
          additionalData = foundRow.slice(7, 10).concat(Array(6).fill(""));
          additionalData[8] = foundRow[2] != null ? String(foundRow[2]) : ""; 
        }
      }
    }

    resultados.push([
      suma,        
      fuente,     
      ...additionalData.slice(0, 8) 
    ]);

    valoresW.push([valorW]);
  });

  const rangoResultados = hoja.getRange(2, 19, resultados.length, encabezados.length);
  rangoResultados.setValues(resultados);
  rangoResultados.setHorizontalAlignment("right");

  const rangoW = hoja.getRange(2, 23, valoresW.length, 1);
  rangoW.setValues(valoresW);
  rangoW.setHorizontalAlignment("right");
}
