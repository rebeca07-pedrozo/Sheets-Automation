function cruzarNumerosConSegundoDoc() {
  const primerDocId = "1zGPjQKZ2_h0VmR-n34CZHS8mXv1V4aSx-0zaCIHTX8I";
  const segundoDocId = "1bR3uCiJdrLyPXqhkvPgrtnDqXBjsR3iRzuPAqh8csSM";

  const ss1 = SpreadsheetApp.openById(primerDocId);
  const hoja1 = ss1.getSheetByName("Hoja 5");

  const ss2 = SpreadsheetApp.openById(segundoDocId);
  const hoja2 = ss2.getSheetByName("Emisiones 28 julio");

  const headers1 = hoja1.getRange(1, 1, 1, hoja1.getLastColumn()).getValues()[0];
  const colNumero1 = headers1.indexOf("EMAIL") + 1; 

  if (colNumero1 === 0) {
    SpreadsheetApp.getUi().alert("No se encontró la columna 'EMAIL' en 'Hoja 5'");
    return;
  }
  const newCols = ["CLAVE", "numero_poliza", "codigo_producto", "fecha_emision", "prima"];

  let lastCol1 = hoja1.getLastColumn();
  let colMap = {}; 

  newCols.forEach(col => {
    let idx = headers1.indexOf(col);
    if (idx === -1) {
      lastCol1++;
      hoja1.getRange(1, lastCol1).setValue(col);
      colMap[col] = lastCol1;
    } else {
      colMap[col] = idx + 1;
    }
  });
