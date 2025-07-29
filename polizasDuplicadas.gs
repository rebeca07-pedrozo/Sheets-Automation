function encontrarPolizasDuplicadas() {
  const archivoOriginal = SpreadsheetApp.openById("1HHEvZBQW-U7Vv2I8FEh1DF45TMds-g7VdVNGxcsgANI");
  const hojaDestino = SpreadsheetApp.getActiveSpreadsheet();
  
  const nombresHojas = ["TOTAL VENTAS ASISTIDAS", "Ventas 322 referidos"];
  const polizas = [];

  nombresHojas.forEach(nombre => {
    const hojaOrigen = archivoOriginal.getSheetByName(nombre);
    const datos = hojaOrigen.getDataRange().getValues();
    const encabezados = datos[0];
    const colPoliza = encabezados.indexOf("Número poliza");
    const colTipo = encabezados.indexOf("Tipo de producto"); 

    
    if (colPoliza !== -1) {
      for (let i = 1; i < datos.length; i++) {
        let valor = datos[i][colPoliza];
        if (valor !== "" && valor != null) {
          valor = valor.toString().trim();
          if (valor !== "") {
            polizas.push(valor);
          }
        }
      }
    }
  });
const conteo = {};
  polizas.forEach(p => {
    conteo[p] = (conteo[p] || 0) + 1;
  });

  const duplicadas = Object.entries(conteo).filter(([_, count]) => count > 1);

  duplicadas.sort((a, b) => a[0].localeCompare(b[0]));
  const hojaDuplicadas = hojaDestino.getSheetByName("Duplicadas") || hojaDestino.insertSheet("Duplicadas");
  hojaDuplicadas.clear();
  hojaDuplicadas.getRange(1, 1).setValue("Número de póliza duplicado");
  hojaDuplicadas.getRange(1, 2).setValue("Cantidad de veces que aparece");

  duplicadas.forEach((fila, i) => {
    hojaDuplicadas.getRange(i + 2, 1).setValue(fila[0]);
    hojaDuplicadas.getRange(i + 2, 2).setValue(fila[1]);
  });
}