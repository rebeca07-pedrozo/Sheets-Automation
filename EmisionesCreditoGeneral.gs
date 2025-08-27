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
