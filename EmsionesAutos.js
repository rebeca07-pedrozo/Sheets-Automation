function EmisionesCreditoPractica_FINAL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Copia de Emisiones 31 ago");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'Copia de Emisiones 31 ago'.");
    return;
  }
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const limpiarDoc = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').toLowerCase();
  const limpiarPlaca = p => String(p || '').replace(/\s/g, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();

  // --- Leer datos de todas las hojas ---
  const datosEmisiones = hoja.getRange(2, 2, ultimaFila - 1, 14).getValues();
  const leadsWADatos = ss.getSheetByName("Leads WA").getDataRange().getValues().slice(1);
  const totalLeadsDatos = ss.getSheetByName("TOTAL LEADS").getDataRange().getValues().slice(1);
  const fbAcumuladoDatos = ss.getSheetByName("Acumulado leads fb").getDataRange().getValues().slice(1);
  const basesDatos = ss.getSheetByName("BASES").getDataRange().getValues().slice(1);

  // --- Crear Mapas para búsquedas rápidas ---
  const mapWA_CC = new Map(), mapWA_Placa = new Map(), mapWA_Fecha = new Map();
  leadsWADatos.forEach(r => {
    const doc = limpiarDoc(r[0]);
    const placa = limpiarPlaca(r[1]);
    const fecha = r[2] ? new Date(r[2]) : null; // Fecha columna C
    mapWA_CC.set(doc, r);
    mapWA_Placa.set(placa, r);
    if(fecha){ mapWA_Fecha.set(doc, fecha); mapWA_Fecha.set(placa, fecha); }
  });

  const mapTL_CC = new Map(), mapTL_Placa = new Map(), mapTL_Fecha = new Map();
  totalLeadsDatos.forEach(r => {
    const doc = limpiarDoc(r[5]);
    const placa = limpiarPlaca(r[8]);
    const fecha = r[15] ? new Date(r[15]) : null; // Fecha columna P
    mapTL_CC.set(doc, r);
    mapTL_Placa.set(placa, r);
    if(fecha){ mapTL_Fecha.set(doc, fecha); mapTL_Fecha.set(placa, fecha); }
  });

  const mapFB_CC = new Map(), mapFB_Correo = new Map(), mapFB_Fecha = new Map();
  fbAcumuladoDatos.forEach(r => {
    const doc = limpiarDoc(r[1]);
    const correo = limpiarCorreo(r[5]);
    const fecha = r[2] ? new Date(r[2]) : null; // Fecha columna C
    mapFB_CC.set(doc, r);
    mapFB_Correo.set(correo, r);
    if(fecha){ mapFB_Fecha.set(doc, fecha); mapFB_Fecha.set(correo, fecha); }
  });

  const mapBASES_CC = new Map(), mapBASES_Correo = new Map(), mapBASES_Fecha = new Map();
  basesDatos.forEach(r => {
    const doc = limpiarDoc(r[0]);
    const correo = limpiarCorreo(r[1]);
    const fecha = r[2] ? new Date(r[2]) : null; // Fecha columna C
    mapBASES_CC.set(doc, r);
    mapBASES_Correo.set(correo, r);
    if(fecha){ mapBASES_Fecha.set(doc, fecha); mapBASES_Fecha.set(correo, fecha); }
  });

  // --- Calcular todo en un solo bucle ---
  const resultadosFinales = datosEmisiones.map(r => {
    const doc = limpiarDoc(r[10]);
    const placa = limpiarPlaca(r[0]);
    const correo = limpiarCorreo(r[12]);

    // Conteos Q-X y total Y
    const conteos = [
      mapWA_CC.has(doc) ? 1 : 0,
      mapWA_Placa.has(placa) ? 1 : 0,
      mapTL_CC.has(doc) ? 1 : 0,
      mapTL_Placa.has(placa) ? 1 : 0,
      mapFB_CC.has(doc) ? 1 : 0,
      mapFB_Correo.has(correo) ? 1 : 0,
      mapBASES_CC.has(doc) ? 1 : 0,
      mapBASES_Correo.has(correo) ? 1 : 0
    ];
    const totalConteo = conteos.reduce((sum, val) => sum + val, 0);

    // Fuente, Medio, Campaña
    let fuente = '', medio = '', campaña = '';
    let fila = null;

    // 1) WA
    fila = mapWA_CC.get(doc) || mapWA_Placa.get(placa);
    if(fila){ fuente=fila[4]; medio=fila[3]; campaña=fila[5]; }

    // 2) TOTAL LEADS
    else if(mapTL_CC.has(doc) || mapTL_Placa.has(placa)){
      fila = mapTL_CC.get(doc) || mapTL_Placa.get(placa);
      fuente = fila[12]; medio = fila[16]; campaña = fila[13];
    }

    // 3) FB
    else if(mapFB_CC.has(doc) || mapFB_Correo.has(correo)){
      fila = mapFB_CC.get(doc) || mapFB_Correo.get(correo);
      fuente = "Facebook"; medio="CPL"; campaña = fila[3];
    }

    // 4) BASES
    else if(mapBASES_CC.has(doc) || mapBASES_Correo.has(correo)){
      fila = mapBASES_CC.get(doc) || mapBASES_Correo.get(correo);
      let filaTLBase = mapTL_CC.get(doc) || mapTL_Placa.get(placa);
      fuente = filaTLBase ? filaTLBase[12] : '';
      medio = filaTLBase ? filaTLBase[16] : '';
      campaña = fila[6];
    }

    // --- Fecha: tomar la más reciente de todas las hojas ---
    const fechas = [
      mapWA_Fecha.get(doc), mapWA_Fecha.get(placa),
      mapTL_Fecha.get(doc), mapTL_Fecha.get(placa),
      mapFB_Fecha.get(doc), mapFB_Fecha.get(correo),
      mapBASES_Fecha.get(doc), mapBASES_Fecha.get(correo)
    ].filter(f => f instanceof Date);

    let fechaFinal = '';
    if(fechas.length > 0){
      fechaFinal = new Date(Math.max(...fechas.map(d => d.getTime())));
    }

    return [...conteos, totalConteo, fuente, medio, campaña, fechaFinal];
  });

  // --- Escribir resultados ---
  hoja.getRange(2, 17, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);

  // --- Encabezados ---
  hoja.getRange("Z1").setValue("Fuente");
  hoja.getRange("AA1").setValue("Medio");
  hoja.getRange("AB1").setValue("Campaña");
  hoja.getRange("AC1").setValue("Fecha");

  SpreadsheetApp.getUi().alert("¡Proceso completado con éxito!");
}
