function EmisionesCreditoPractica_Todas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Emisiones 11 ago");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  hoja.getRange(2, 14, ultimaFila - 1, 12).clearContent();

  const cedulas = hoja.getRange(2, 9, ultimaFila - 1, 1).getValues().map(r => String(r[0]).trim());
  const correos = hoja.getRange(2, 11, ultimaFila - 1, 1).getValues().map(r => String(r[0]).trim());

  const totalLeadsSheet = ss.getSheetByName("TOTAL LEADS");
  const totalLeadsData = totalLeadsSheet.getRange(2, 1, totalLeadsSheet.getLastRow() - 1, 14).getValues();
  const totalLeadsMap = new Map();
  totalLeadsData.forEach(row => {
    const cedula = String(row[4]).trim();
    const email = String(row[2]).trim();
    if (cedula) totalLeadsMap.set(cedula, row);
    if (email) totalLeadsMap.set(email, row);
  });

  const leads322Sheet = ss.getSheetByName("Leads 322");
  const leads322Data = leads322Sheet.getRange(2, 12, leads322Sheet.getLastRow() - 1, 1).getValues();
  const leads322Map = new Map();
  leads322Data.forEach(row => {
    const cedula = String(row[0]).trim();
    if (cedula) leads322Map.set(cedula, true);
  });

  const basesSheet = ss.getSheetByName("BASES");
  const basesData = basesSheet.getRange(2, 1, basesSheet.getLastRow() - 1, 10).getValues();
  const basesMap = new Map();
  basesData.forEach(row => {
    const cedula = String(row[0]).trim();
    const email = String(row[1]).trim();
    if (cedula) basesMap.set(cedula, row);
    if (email) basesMap.set(email, row);
  });

  const resultados = cedulas.map((cedula, i) => {
    const correo = correos[i];
    
    const countN = (cedula && totalLeadsMap.has(cedula)) ? 1 : 0;
    
    const countO = (correo && totalLeadsMap.has(correo)) ? 1 : 0;
    
    const countP = (cedula && leads322Map.has(cedula)) ? 1 : 0;
    
    const countQ = (cedula && basesMap.has(cedula)) ? 1 : 0;
    
    const countR = (correo && basesMap.has(correo)) ? 1 : 0;
    
    const suma = countN + countO + countP + countQ + countR;
    
    let fuente = "";
    let additionalData = ["", "", "", "", ""]; 
    
    if (suma > 0) {
      if (countN === 1 || countO === 1) {
        fuente = "TOTAL LEADS";
        const foundRow = totalLeadsMap.has(cedula) ? totalLeadsMap.get(cedula) : totalLeadsMap.get(correo);
        if (foundRow) {
          additionalData = foundRow.slice(9, 14);
        }
      } else if (countP === 1) {
        fuente = "Leads 322";
      } else if (countQ === 1 || countR === 1) {
        fuente = "BASES";
        const foundRow = basesMap.has(cedula) ? basesMap.get(cedula) : basesMap.get(correo);
        if (foundRow) {
          const basesDataSlice = foundRow.slice(7, 10);
          additionalData = [...basesDataSlice, "", ""]; 
        }
      }
    }

    return [countN, countO, countP, countQ, countR, suma, fuente, ...additionalData];
  });

  hoja.getRange(2, 14, resultados.length, 12).setValues(resultados);
}