function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Emisiones')
    .addItem('Ejecutar Integral', 'menuEjecutarIntegral')
    .addItem('Ejecutar A su medida', 'menuEjecutarMedida')
    .addToUi();
}

function menuEjecutarIntegral() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Nombre de la hoja de "Integral" (ej: Emisiones Integral 13 oct):');
  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText().trim()) {
    const nombreHoja = response.getResponseText().trim();
    try {
      // CAMBIA 'EmisionesIntegral' por 'CruceDatosSaludIntegral'
      CruceDatosSaludIntegral(nombreHoja); 
      ui.alert('Proceso completado para ' + nombreHoja);
    } catch (e) {
      ui.alert('Error: ' + e.message);
    }
  } else {
    ui.alert('Operación cancelada.');
  }
}

function menuEjecutarMedida() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Nombre de la hoja de "A su medida" (ej: Emisiones A su medida 13 oct):');
  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText().trim()) {
    const nombreHoja = response.getResponseText().trim();
    try {
      SaludAsuMedida(nombreHoja);
      ui.alert('Proceso completado para ' + nombreHoja);
    } catch (e) {
      ui.alert('Error: ' + e.message);
    }
  } else {
    ui.alert('Operación cancelada.');
  }
}
