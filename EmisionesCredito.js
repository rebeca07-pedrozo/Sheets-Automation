function EmisionesCreditoPractica_Todas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Emisiones 24");
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
                      .map(r => r[0] != null ? String(r[0]).trim() : "");
  const correos = hoja.getRange(2, 12, ultimaFila - 1, 1)
                      .getValues()
                      .map(r => r[0] != null ? String(r[0]).trim().toLowerCase() : "");

  const totalLeadsSheet = ss.getSheetByName("TOTAL LEADS");
  const totalLeadsData = totalLeadsSheet.getRange(2, 1, totalLeadsSheet.getLastRow() - 1, 14).getValues();
  const totalLeadsMap = new Map();
  totalLeadsData.forEach(row => {
    const cedula = row[4] != null ? String(row[4]).trim() : "";
    const email = row[2] != null ? String(row[2]).trim().toLowerCase() : "";
    if (cedula) totalLeadsMap.set(cedula, row);
    if (email) totalLeadsMap.set(email, row);
  });

  const leads322Sheet = ss.getSheetByName("Leads 322");
  const leads322Data = leads322Sheet.getRange(2, 1, leads322Sheet.getLastRow() - 1, 26).getValues();
  const leads322Map = new Map();
  leads322Data.forEach(row => {
    const cedula = row[11] != null ? String(row[11]).trim() : "";
    if (cedula) leads322Map.set(cedula, row);
  });

  const basesSheet = ss.getSheetByName("BASES");
  const basesData = basesSheet.getRange(2, 1, basesSheet.getLastRow() - 1, 10).getValues();
  const basesMap = new Map();
  basesData.forEach(row => {
    const cedula = row[0] != null ? String(row[0]).trim() : "";
    const email = row[1] != null ? String(row[1]).trim().toLowerCase() : "";
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

    const countN = (cedula && totalLeadsMap.has(cedula)) ? 1 : 0;
    const countO = (correo && totalLeadsMap.has(correo)) ? 1 : 0;
    const countP = (cedula && leads322Map.has(cedula)) ? 1 : 0;
    const countQ = (cedula && basesMap.has(cedula)) ? 1 : 0;
    const countR = (correo && basesMap.has(correo)) ? 1 : 0;

    const suma = countN + countO + countP + countQ + countR;

    let fuente = "", med = "", campaña = "";
    let valorZ = "", valorAA = "";
    let valorW = "";

    if (suma > 0) {
      if (countN === 1 || countO === 1) {
        fuente = "TOTAL LEADS";
        const foundRow = totalLeadsMap.get(cedula) || totalLeadsMap.get(correo);
        if (foundRow) {
          valorW = formatearFecha(foundRow[13]);
          med = foundRow[9] != null ? String(foundRow[9]) : "";
          campaña = foundRow[10] != null ? String(foundRow[10]) : "";
        }
      }
      else if (countP === 1) {
        fuente = "322";
        const foundRow = leads322Map.get(cedula);
        if (foundRow) {
          valorW = formatearFecha(foundRow[0]);
          med = "";
          campaña = foundRow[25] != null ? String(foundRow[25]) : "";
        }
      }
      else if (countQ === 1 || countR === 1) {
        const foundRow = basesMap.get(cedula) || basesMap.get(correo);
        if (foundRow) {
          fuente = foundRow[7] != null ? String(foundRow[7]) : "BASES";
          med = foundRow[3] != null ? String(foundRow[3]) : "";
          campaña = foundRow[6] != null ? String(foundRow[6]) : "";
          valorW = formatearFecha(foundRow[2]);
          if (fuente === "ESTRATEGO") {
            valorZ = foundRow[3] != null ? String(foundRow[3]) : "";
            valorAA = foundRow[4] != null ? String(foundRow[4]) : "";
          }
        }
      }
    }

    resultados.push([
      countN, countO, countP, countQ, countR,
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
