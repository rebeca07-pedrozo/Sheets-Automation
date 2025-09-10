function consolidarVentasRobusto() {
  const hojaDestinoNombre = 'Copia de Consolidado_ventas_formula';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName(hojaDestinoNombre);

  hojaDestino.clearContents();

  const columnasInteres = [
    'Fecha', 'Consecutivo', 'Campaña', 'Nombre Campañas', 'Intención',
    'Audiencia', 'Segmento', 'Canal', 'Entregas', 'Aperturas',
    'Clics', 'Leads', 'Venta', 'Primas', '% Apertura', '% CTO',
    '% Conversión Lead', '% Conversión Venta'
  ];

  const hojas = ss.getSheets();
  let filasConsolidadas = [];

  hojas.forEach(hoja => {
    if (hoja.getName() === hojaDestinoNombre) return;

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;

    const encabezados = datos[0];
    const indiceIntencion = encabezados.indexOf('Intención');
    if (indiceIntencion === -1) return;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const valorIntencion = fila[indiceIntencion];
      if (!valorIntencion) continue;

      const valorMinuscula = valorIntencion.toString().toLowerCase();
      if (valorMinuscula.startsWith('venta') || valorMinuscula === 'incentivo uso asistencia bancaseguros') {
        const filaFiltrada = columnasInteres.map(col => {
          const idx = encabezados.indexOf(col);
          return idx !== -1 ? fila[idx] : '';
        });

        let fechaObj;
        const fechaStr = filaFiltrada[0];
        if (fechaStr && typeof fechaStr === 'string' && fechaStr.includes('/')) {
          const partes = fechaStr.split('/');
          fechaObj = new Date(partes[2], partes[1] - 1, partes[0]); 
        } else {
          fechaObj = new Date(fechaStr); 
        }

        if (isNaN(fechaObj.getTime())) {
          fechaObj = new Date(9999, 0, 1); 
        }

        filaFiltrada.push(fechaObj);
        filasConsolidadas.push(filaFiltrada);
      }
    }
  });

  if (filasConsolidadas.length === 0) return;

  filasConsolidadas.sort((a, b) => a[a.length - 1] - b[a.length - 1]);

  filasConsolidadas = filasConsolidadas.map(fila => fila.slice(0, columnasInteres.length));

  hojaDestino.getRange(1, 1, 1, columnasInteres.length).setValues([columnasInteres]);
  hojaDestino.getRange(2, 1, filasConsolidadas.length, columnasInteres.length).setValues(filasConsolidadas);

  Logger.log('Consolidación completada: ' + filasConsolidadas.length + ' filas copiadas.');
}
