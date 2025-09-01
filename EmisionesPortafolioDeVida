function EmisionesCreditoPractica_Todas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("EMISIONES 24 AGO");
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "EMISIONES 24 AGO".');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const encabezados = [
    "TotalLeads CC", "TotalLeads Email", "Leadforms CC", "Leadforms Email", "322", "Bases CC", "Bases Email",
    "venta", "fuente", "med", "campaña", "fecha lead",
    "prod", "cruce cami", "Prioridad", "Base", "Cruce Email"
  ];

  hoja.getRange(1, 15, 1, encabezados.length)
      .setValues([encabezados])
      .setHorizontalAlignment("right");
  hoja.getRange(2, 15, ultimaFila - 1, encabezados.length).clearContent();

  const cedulas = hoja.getRange(2, 10, ultimaFila - 1, 1)
                      .getValues()
                      .map(r => r[0] ? String(r[0]).trim() : "");
  const correos = hoja.getRange(2, 12, ultimaFila - 1, 1)
                      .getValues()
                      .map(r => r[0] ? String(r[0]).trim().toLowerCase() : "");

  function getSheetData(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    return sheet && sheet.getLastRow() > 1 ? sheet.getDataRange().getValues().slice(1) : [];
  }

  const totalLeadsData = getSheetData("TOTAL LEADS");
  const totalLeadsCedulaMap = new Map();
  const totalLeadsEmailMap = new Map();
  totalLeadsData.forEach(row => {
    const cedula = row[4] ? String(row[4]).trim() : "";
    const email = row[0] ? String(row[0]).trim().toLowerCase() : "";
    if (cedula) totalLeadsCedulaMap.set(cedula, row);
    if (email) totalLeadsEmailMap.set(email, row);
  });

  const leadformsData = getSheetData("Leadforms");
  const leadformsMap = new Map();
  leadformsData.forEach(row => {
    const cedula = row[1] ? String(row[1]).trim() : "";
    const email = row[0] ? String(row[0]).trim().toLowerCase() : "";
    if (cedula) leadformsMap.set(cedula, row);
    if (email) leadformsMap.set(email, row);
  });
  
  const leads322Data = getSheetData("Leads 322");
  const leads322Map = new Map();
  leads322Data.forEach(row => {
    const cedula = row[11] ? String(row[11]).trim() : "";
    if (cedula) leads322Map.set(cedula, row);
  });

  const basesData = getSheetData("BASES");
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
    const yyyy = d.getFullYear();
    const MM = ("0" + (d.getMonth() + 1)).slice(-2);
    const dd = ("0" + d.getDate()).slice(-2);
    const HH = ("0" + d.getHours()).slice(-2);
    const mm = ("0" + d.getMinutes()).slice(-2);
    const ss = ("0" + d.getSeconds()).slice(-2);
    return `${yyyy}-${MM}-${dd} ${HH}:${mm}:${ss}`;
  }

  const resultados = [];

  cedulas.forEach((cedulaRaw, i) => {
    const cedula = cedulaRaw;
    const correo = correos[i];

    const countN = (cedula && totalLeadsCedulaMap.has(cedula)) ? 1 : 0;
    const countO = (correo && totalLeadsEmailMap.has(correo)) ? 1 : 0;
    const countX = (cedula && leadformsMap.has(cedula)) ? 1 : 0;
    const countY = (correo && leadformsMap.has(correo)) ? 1 : 0;
    const countP = (cedula && leads322Map.has(cedula)) ? 1 : 0;
    const countQ = (cedula && basesMap.has(cedula)) ? 1 : 0;
    const countR = (correo && basesMap.has(correo)) ? 1 : 0;

    const suma = countN + countO + countX + countY + countP + countQ + countR;

    let fuente = "", med = "", campaña = "";
    let valorZ = "", valorAA = "";
    let valorW = "";

    if (suma > 0) {
      if (countN === 1 || countO === 1) {
        let foundRow = null;
        if (countN === 1) {
            foundRow = totalLeadsCedulaMap.get(cedula);
        } else if (countO === 1) {
            foundRow = totalLeadsEmailMap.get(correo);
        }
        
        if (foundRow) {
            if (countN === 1) { 
                fuente = foundRow[5] ? String(foundRow[5]) : "";
                med = foundRow[6] ? String(foundRow[6]) : "";
                campaña = foundRow[7] ? String(foundRow[7]) : "";
                valorW = formatearFecha(foundRow[8]);
            } else if (countO === 1) { 
                fuente = foundRow[5] ? String(foundRow[5]) : "";
                med = foundRow[6] ? String(foundRow[6]) : "";
                campaña = foundRow[7] ? String(foundRow[7]) : "";
                valorW = formatearFecha(foundRow[8]);
            }
        }
      } else if (countX === 1 || countY === 1) {
        fuente = "Facebook";
        med = "CPL";
        const foundRow = leadformsMap.get(cedula) || leadformsMap.get(correo);
        if (foundRow) {
          campaña = foundRow[3] ? String(foundRow[3]) : "";
          valorW = formatearFecha(foundRow[2]);
        }
      } else if (countP === 1) {
        fuente = "322";
        const foundRow = leads322Map.get(cedula);
        if (foundRow) {
          valorW = formatearFecha(foundRow[0]);
          med = "";
          campaña = foundRow[25] ? String(foundRow[25]) : "";
        }
      } else if (countQ === 1 || countR === 1) {
        const foundRow = basesMap.get(cedula) || basesMap.get(correo);
        if (foundRow) {
          fuente = foundRow[7] ? String(foundRow[7]) : "BASES";
          med = foundRow[3] ? String(foundRow[3]) : "";
          campaña = foundRow[6] ? String(foundRow[6]) : "";
          valorW = formatearFecha(foundRow[2]);
          if (fuente === "ESTRATEGO") {
            valorZ = foundRow[3] ? String(foundRow[3]) : "";
            valorAA = foundRow[4] ? String(foundRow[4]) : "";
          }
        }
      }
    }

    resultados.push([
      countN, countO, countX, countY, countP, countQ, countR,
      suma,
      fuente, med, campaña,
      valorW,
      "", "", 
      valorZ,
      valorAA,
      "" 
    ]);
  });

  hoja.getRange(2, 15, resultados.length, encabezados.length)
       .setValues(resultados)
       .setHorizontalAlignment("right");
}