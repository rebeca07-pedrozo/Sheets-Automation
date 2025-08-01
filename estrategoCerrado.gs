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
  const lastRow1 = hoja1.getLastRow();
  const numeros1 = hoja1.getRange(2, colNumero1, lastRow1 - 1).getValues();

  const lastRow2 = hoja2.getLastRow();
  const data2 = hoja2.getRange(2, 1, lastRow2 - 1, hoja2.getLastColumn()).getValues();

  let mapaDocs = {};
  data2.forEach(row => {
    let key = String(row[8]).trim();
    if (key) {
      mapaDocs[key] = {
        CLAVE: row[0],             
        numero_poliza: row[1],     
        codigo_producto: row[2],   
        fecha_emision: row[3],     
        prima: row[6],             
      };
    }
  });
   let resultadoClave = [];
  let resultadoPoliza = [];
  let resultadoProducto = [];
  let resultadoFecha = [];
  let resultadoPrima = [];

  for (let i = 0; i < numeros1.length; i++) {
    let num = String(numeros1[i][0]).trim();
    if (num && mapaDocs.hasOwnProperty(num)) {
      resultadoClave.push([mapaDocs[num].CLAVE]);
      resultadoPoliza.push([mapaDocs[num].numero_poliza]);
      resultadoProducto.push([mapaDocs[num].codigo_producto]);
      resultadoFecha.push([mapaDocs[num].fecha_emision]);
      resultadoPrima.push([mapaDocs[num].prima]);
    } else {
      resultadoClave.push([""]);
      resultadoPoliza.push([""]);
      resultadoProducto.push([""]);
      resultadoFecha.push([""]);
      resultadoPrima.push([""]);
    }
  }
  hoja1.getRange(2, colMap["CLAVE"], resultadoClave.length).setValues(resultadoClave);
  hoja1.getRange(2, colMap["numero_poliza"], resultadoPoliza.length).setValues(resultadoPoliza);
  hoja1.getRange(2, colMap["codigo_producto"], resultadoProducto.length).setValues(resultadoProducto);
  hoja1.getRange(2, colMap["fecha_emision"], resultadoFecha.length).setValues(resultadoFecha);
  hoja1.getRange(2, colMap["prima"], resultadoPrima.length).setValues(resultadoPrima);

}