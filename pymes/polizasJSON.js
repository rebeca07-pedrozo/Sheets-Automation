function extraerNumeroPoliza() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const FILA_INICIO = 2; 
  const COL_ORIGEN = 1;  
  const COL_DESTINO = 2; 

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_INICIO) return;

  const rangoOrigen = hoja.getRange(FILA_INICIO, COL_ORIGEN, ultimaFila - FILA_INICIO + 1, 1);
  const datos = rangoOrigen.getValues();
  const salida = [];

  const reNumberPoliza = /"numberPoliza"\s*:\s*"([^"]+)"/i;

  datos.forEach(fila => {
    const texto = (fila[0] || '').toString().trim();
    let numero = '';

    if (texto) {
      try {
        if (/^[\[\{]/.test(texto)) {
          const js = JSON.parse(texto);
          if (Array.isArray(js) && js.length > 0 && js[0].numberPoliza) {
            numero = js[0].numberPoliza;
          } else if (js.numberPoliza) {
            numero = js.numberPoliza;
          }
        }
      } catch (e) {
      }

      if (!numero) {
        const match = texto.match(reNumberPoliza);
        if (match) numero = match[1];
      }
    }

    salida.push([numero]);
  });

  hoja.getRange(FILA_INICIO, COL_DESTINO, salida.length, 1).setValues(salida);
}
