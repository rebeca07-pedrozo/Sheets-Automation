function EmisionesPDeVidaNuevo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Emisiones 14 sep");
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "EMISIONES 24 AGO".');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const encabezados = [
    "TotalLeads CC", "TotalLeads Email", "322", "Bases CC", "Bases Email",
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

    const countTotalLeadsCC = (cedula && totalLeadsCedulaMap.has(cedula)) ? 1 : 0;
    const countTotalLeadsEmail = (correo && totalLeadsEmailMap.has(correo)) ? 1 : 0;
    const count322 = (cedula && leads322Map.has(cedula)) ? 1 : 0;
    const countBasesCC = (cedula && basesMap.has(cedula)) ? 1 : 0;
    const countBasesEmail = (correo && basesMap.has(correo)) ? 1 : 0;

    const suma = countTotalLeadsCC + countTotalLeadsEmail + count322 + countBasesCC + countBasesEmail;

    let fuente = "", med = "", campaña = "";
    let valorZ = "", valorAA = "";
    let valorW = "";

    if (suma > 0) {
      if (countTotalLeadsCC === 1 || countTotalLeadsEmail === 1) {
        let foundRow = null;
        if (countTotalLeadsCC === 1) {
            foundRow = totalLeadsCedulaMap.get(cedula);
        } else if (countTotalLeadsEmail === 1) {
            foundRow = totalLeadsEmailMap.get(correo);
        }
        
        if (foundRow) {
            fuente = foundRow[5] ? String(foundRow[5]) : ""; 
            med = foundRow[6] ? String(foundRow[6]) : ""; 
            campaña = foundRow[7] ? String(foundRow[7]) : ""; 
            valorW = formatearFecha(foundRow[8]); 
        }
      } else if (count322 === 1) {
        fuente = "322";
        const foundRow = leads322Map.get(cedula);
        if (foundRow) {
          valorW = formatearFecha(foundRow[0]);
          med = "";
          campaña = foundRow[25] ? String(foundRow[25]) : "";
        }
      } else if (countBasesCC === 1 || countBasesEmail === 1) {
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
      countTotalLeadsCC, countTotalLeadsEmail, count322, countBasesCC, countBasesEmail,
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