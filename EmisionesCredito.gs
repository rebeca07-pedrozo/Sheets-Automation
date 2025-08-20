function EmisionesCreditoPractica(){
  const ss=SpreadsheetApp.getActiveSheet();
  const hoja=ss.getSheetByName("Emisiones 11 ago");
  const ultimaFila=hoja.getLastRow();
  const columna9=hoja.getRange(2,9,ultimaFila,-1,1).getValues();
  const columna11=hoja.getRange(2,11,ultimaFila,-1,1).getValues();

}

 
