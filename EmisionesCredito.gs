function actualizarEmisionesCompletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Emisiones 11 ago");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; 

  hoja.getRange(2, 20, ultimaFila - 1, 7).clearContent();

  const cedulaEmisiones = hoja.getRange(2, 9, ultimaFila - 1, 1).getValues(); 
  const correoEmisiones = hoja.getRange(2, 11, ultimaFila - 1, 1).getValues(); 

  const totalLeads = ss.getSheetByName("TOTAL LEADS");
  const datosLeads = totalLeads.getRange(2, 1, totalLeads.getLastRow() - 1, 14).getValues(); 

  const bases = ss.getSheetByName("BASES");
  const datosBases = bases.getRange(2, 1, bases.getLastRow() - 1, 8).getValues(); 

  const leads322 = ss.getSheetByName("Leads 322");
  const datos322 = leads322.getRange(2, 12, leads322.getLastRow() - 1, 1).getValues(); 

  for (let i = 0; i < cedulaEmisiones.length; i++) {
    const cedula = cedulaEmisiones[i][0];
    const correo = correoEmisiones[i][0];
    let resultado = []; 

    let encontradoLeads = datosLeads.find(row => row[4] == cedula || row[2] == correo);
    if (encontradoLeads) {
      resultado = [encontradoLeads[9], encontradoLeads[10], encontradoLeads[11], encontradoLeads[13]]; 
    } else {
      let encontrado322 = datos322.find(row => row[0] == cedula);
      if (encontrado322) {
        resultado = ["322"];
      } else {
        let encontradoBases = datosBases.find(row => row[0] == cedula || row[1] == correo);
        if (encontradoBases) {
          resultado = [encontradoBases[7]]; 
        }
      }
    }

    if (resultado.length > 0) {
      hoja.getRange(i + 2, 20, 1, resultado.length).setValues([resultado]);
    }
  }
}
