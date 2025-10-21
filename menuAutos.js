function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Emisiones 🚗')
    .addItem('Ejecutar Emisiones Autos', 'menuEjecutarAutos')
    .addToUi();
}

function menuEjecutarAutos() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Nombre de la hoja de Emisiones Autos (ej: Copia de Emisiones 7 oct):');

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

