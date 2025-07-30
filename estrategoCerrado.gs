function cruzarNumerosConSegundoDoc() {
  const primerDocId = "1zGPjQKZ2_h0VmR-n34CZHS8mXv1V4aSx-0zaCIHTX8I";
  const segundoDocId = "1bR3uCiJdrLyPXqhkvPgrtnDqXBjsR3iRzuPAqh8csSM";

  const ss1 = SpreadsheetApp.openById(primerDocId);
  const hoja1 = ss1.getSheetByName("Hoja 5");