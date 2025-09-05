function EmsionesHogarOriginal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Copia de Ventas 31 agosto");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Copia de Ventas 31 agosto'");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const encabezados = [
    "TotalLeads CC", "TotalLeads Email",
    "FB CC", "FB Email",
    "Bases CC", "Bases Email",
    "suma", "fuente", "med", "campaña", "fecha lead", "dif fecha"
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

  const totalLeadsSheet = ss.getSheetByName("Leads");
  const totalLeadsData = totalLeadsSheet.getRange(2, 1, totalLeadsSheet.getLastRow() - 1, 16).getValues();
  const totalLeadsMap = new Map();
  totalLeadsData.forEach(row => {
    const cedula = row[3] != null ? String(row[3]).trim() : "";
    const email = row[4] != null ? String(row[4]).trim().toLowerCase() : "";
    if (cedula) totalLeadsMap.set(cedula, row);
    if (email) totalLeadsMap.set(email, row);
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

  const fbSheet = ss.getSheetByName("Leadforms FB");
  const fbData = fbSheet.getRange(2, 1, fbSheet.getLastRow() - 1, 9).getValues();
  const fbMapCC = new Map();
  const fbMapEmail = new Map();
  fbData.forEach(row => {
    const cedula = row[5] != null ? String(row[5]).trim() : ""; 
    const email = row[3] != null ? String(row[3]).trim().toLowerCase() : ""; 
    if (cedula) fbMapCC.set(cedula, row);
    if (email) fbMapEmail.set(email, row);
  });

  function formatearFechaHora(fecha) {
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

    let countLeadsCC = 0, countLeadsEmail = 0;
    let countFBCC = 0, countFBEmail = 0;
    let countBasesCC = 0, countBasesEmail = 0;
    let fuente = "", med = "", campaña = "", fechaLead = "", difFecha = "";

    if ((cedula && totalLeadsMap.has(cedula)) || (correo && totalLeadsMap.has(correo))) {
      const foundRow = totalLeadsMap.get(cedula) || totalLeadsMap.get(correo);
      countLeadsCC = cedula && totalLeadsMap.has(cedula) ? 1 : 0;
      countLeadsEmail = correo && totalLeadsMap.has(correo) ? 1 : 0;
      fechaLead = formatearFechaHora(foundRow[15]);
      fuente = foundRow[9] != null ? String(foundRow[9]) : "";
      med = foundRow[10] != null ? String(foundRow[10]) : "";
      campaña = foundRow[11] != null ? String(foundRow[11]) : "";
    } else if ((cedula && fbMapCC.has(cedula)) || (correo && fbMapEmail.has(correo))) {
      const foundRow = fbMapCC.get(cedula) || fbMapEmail.get(correo);
      countFBCC = cedula && fbMapCC.has(cedula) ? 1 : 0;
      countFBEmail = correo && fbMapEmail.has(correo) ? 1 : 0;
      fechaLead = formatearFechaHora(foundRow[6]);
      fuente = foundRow[7] != null ? String(foundRow[7]) : "FB";
      med = "CPL";  
      campaña = foundRow[8] != null ? String(foundRow[8]) : "";
    } else if ((cedula && basesMap.has(cedula)) || (correo && basesMap.has(correo))) {
      const foundRow = basesMap.get(cedula) || basesMap.get(correo);
      countBasesCC = cedula && basesMap.has(cedula) ? 1 : 0;
      countBasesEmail = correo && basesMap.has(correo) ? 1 : 0;
      fechaLead = formatearFechaHora(foundRow[2]);
      fuente = foundRow[7] != null ? String(foundRow[7]) : "BASES";
      med = foundRow[3] != null ? String(foundRow[3]) : "";
      campaña = foundRow[6] != null ? String(foundRow[6]) : "";
    }

    const suma = countLeadsCC + countLeadsEmail + countFBCC + countFBEmail + countBasesCC + countBasesEmail;

    resultados.push([
      countLeadsCC, countLeadsEmail,
      countFBCC, countFBEmail,
      countBasesCC, countBasesEmail,
      suma,
      fuente,
      med,
      campaña,
      fechaLead,
      difFecha
    ]);
  });

  hoja.getRange(2, 15, resultados.length, encabezados.length)
       .setValues(resultados)
       .setHorizontalAlignment("right");
}
