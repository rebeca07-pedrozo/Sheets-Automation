function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automatizaciones')
    .addItem('Emisiones funcionarios', 'menuFuncionarios')
    .addToUi();
}

function menuEjecutarAutos() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Nombre de la hoja de Emisiones (ej: Copia de Emisiones 7 oct):');

  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText().trim()) {
    const nombreHoja = response.getResponseText().trim();
    try {
      EmisionesAutosCruzados(nombreHoja);
      ui.alert('Proceso completado correctamente para: ' + nombreHoja);
    } catch (e) {
      ui.alert('Error al ejecutar: ' + e.message);
    }
  } else {
    ui.alert('Operación cancelada.');
  }
}

