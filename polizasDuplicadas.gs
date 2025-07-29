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
