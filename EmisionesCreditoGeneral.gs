function EmisionesCreditoPractica_Todas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Sheet18");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;