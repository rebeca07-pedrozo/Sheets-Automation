function separarDuplicadosFB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const hojaAcumulado = ss.getSheetByName("Acumulado leads fb");
  const hojaFormulario = ss.getSheetByName("formulario faltante fb");
  
  const filaInicioAcumulado = 3783;
  const datosAcumulado = hojaAcumulado.getRange(filaInicioAcumulado, 1, hojaAcumulado.getLastRow() - filaInicioAcumulado + 1, 6).getValues(); 
  
  const encabezadoFormulario = hojaFormulario.getRange(1, 1, 1, 6).getValues()[0]; 
  const datosFormulario = hojaFormulario.getRange(2, 1, hojaFormulario.getLastRow() - 1, 6).getValues(); 
  const clavesAcumulado = new Set(datosAcumulado.map(fila => {
    return [
      fila[0], 
      fila[1], 
      fila[2], 
      fila[3], 
      fila[4]  
    ].join("|").toLowerCase();
  }));

  let noDuplicados = [];
  let duplicados = [];

  datosFormulario.forEach(fila => {
    const clave = [
      fila[0], 
      fila[1], 
      fila[2], 
      fila[3], 
      fila[4]  
    ].join("|").toLowerCase();

    if (clavesAcumulado.has(clave)) {
      duplicados.push(fila);
    } else {
      noDuplicados.push(fila);
    }
  });

  hojaFormulario.clearContents();

  hojaFormulario.getRange(1, 1).setValue("Elementos no duplicados");
  hojaFormulario.getRange(2, 1, 1, 6).setValues([encabezadoFormulario]);
  if (noDuplicados.length > 0) {
    hojaFormulario.getRange(3, 1, noDuplicados.length, noDuplicados[0].length).setValues(noDuplicados);
  }

  // Escribir duplicados
  let filaInicioDuplicados = noDuplicados.length + 5;
  hojaFormulario.getRange(filaInicioDuplicados - 1, 1).setValue("Elementos duplicados");
  hojaFormulario.getRange(filaInicioDuplicados, 1, 1, 6).setValues([encabezadoFormulario]);
  if (duplicados.length > 0) {
    hojaFormulario.getRange(filaInicioDuplicados + 1, 1, duplicados.length, duplicados[0].length).setValues(duplicados);
  }

  SpreadsheetApp.getUi().alert(`Proceso completado:
No duplicados: ${noDuplicados.length}
Duplicados: ${duplicados.length}`);
}
