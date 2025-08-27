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