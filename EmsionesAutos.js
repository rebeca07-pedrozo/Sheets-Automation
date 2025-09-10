function EmisionesAutosOptimizado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Copia de Emisiones 8 sept");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja.");
    return;
  }
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const limpiarDoc = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').toLowerCase();
  const limpiarPlaca = p => String(p || '').replace(/\s/g, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();

  const datosEmisiones = hoja.getRange(2, 2, ultimaFila - 1, 14).getValues();
  const leadsWADatos = ss.getSheetByName("Leads WA").getDataRange().getValues().slice(1);
  const totalLeadsDatos = ss.getSheetByName("TOTAL LEADS").getDataRange().getValues().slice(1);
  const fbAcumuladoDatos = ss.getSheetByName("Acumulado leads fb").getDataRange().getValues().slice(1);
  const basesDatos = ss.getSheetByName("BASES").getDataRange().getValues().slice(1);

  const mapWA_CC = new Map(), mapWA_Placa = new Map();
  leadsWADatos.forEach(r => { mapWA_CC.set(limpiarDoc(r[0]), r); mapWA_Placa.set(limpiarPlaca(r[1]), r); });

  const mapTL_CC = new Map(), mapTL_Placa = new Map();
  totalLeadsDatos.forEach(r => { mapTL_CC.set(limpiarDoc(r[5]), r); mapTL_Placa.set(limpiarPlaca(r[8]), r); });

  const mapFB_CC = new Map(), mapFB_Correo = new Map();
  fbAcumuladoDatos.forEach(r => { 
    mapFB_CC.set(limpiarDoc(r[1]), r);       
    mapFB_Correo.set(limpiarCorreo(r[5]), r); 
  });

  const mapBASES_CC = new Map(), mapBASES_Correo = new Map();
  basesDatos.forEach(r => { mapBASES_CC.set(limpiarDoc(r[0]), r); mapBASES_Correo.set(limpiarCorreo(r[1]), r); });

  function buscarEnArrayParaCampos(arrayRows, doc, placa, correo, campIdx, medioIdx) {
    for (let i = 0; i < arrayRows.length; i++) {
      const row = arrayRows[i];
      let match = false;
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] || '');
        if (doc && cell.replace(/\./g, '').replace(/\s/g, '').toLowerCase() === doc) { match = true; break; }
        if (placa && cell.replace(/\s/g, '').toLowerCase() === placa) { match = true; break; }
        if (correo && cell.trim().toLowerCase() === correo) { match = true; break; }
      }
      if (match) {
        const camp = (campIdx !== null && row[campIdx]) ? row[campIdx] : null;
        const medio = (medioIdx !== null && row[medioIdx]) ? row[medioIdx] : null;
        if (camp || medio) return { row: row, camp: camp, medio: medio };
      }
    }
    return null;
  }

  const resultadosFinales = datosEmisiones.map(r => {
    const doc = limpiarDoc(r[10]);     
    const placa = limpiarPlaca(r[0]);  
    const correo = limpiarCorreo(r[13]); 

    const candidates = [
      { name: 'WA_CC', fila: mapWA_CC.get(doc), arr: leadsWADatos, fuenteIdx: 3, medioIdx: 4, campIdx: 5, fechaIdx: 2, defaultFuente: '', defaultMedio: '' },
      { name: 'WA_Placa', fila: mapWA_Placa.get(placa), arr: leadsWADatos, fuenteIdx: 3, medioIdx: 4, campIdx: 5, fechaIdx: 2, defaultFuente: '', defaultMedio: '' },
      { name: 'TL_CC', fila: mapTL_CC.get(doc),  arr: totalLeadsDatos, fuenteIdx: 12, medioIdx: 16, campIdx: 13, fechaIdx: 15, defaultFuente: '',       defaultMedio: '' },
      { name: 'TL_Placa', fila: mapTL_Placa.get(placa), arr: totalLeadsDatos, fuenteIdx: 12, medioIdx: 16, campIdx: 13, fechaIdx: 15, defaultFuente: '',       defaultMedio: '' },
      { name: 'FB_CC', fila: mapFB_CC.get(doc),  arr: fbAcumuladoDatos, fuenteIdx: null, medioIdx: null, campIdx: 3,  fechaIdx: 2,  defaultFuente: 'Facebook', defaultMedio: 'CPL' },
      { name: 'FB_Correo', fila: mapFB_Correo.get(correo), arr: fbAcumuladoDatos, fuenteIdx: null, medioIdx: null, campIdx: 3, fechaIdx: 2,  defaultFuente: 'Facebook', defaultMedio: 'CPL' },
      { name: 'BASES_CC', fila: mapBASES_CC.get(doc), arr: basesDatos, fuenteIdx: null, medioIdx: null, campIdx: 6, fechaIdx: 2, defaultFuente: '',       defaultMedio: '' },
      { name: 'BASES_Correo', fila: mapBASES_Correo.get(correo), arr: basesDatos, fuenteIdx: null, medioIdx: null, campIdx: 6, fechaIdx: 2, defaultFuente: '',       defaultMedio: '' }
    ];

    const conteos = candidates.map(c => c.fila ? 1 : 0);

    let primary = null;
    for (let c of candidates) { if (c.fila) { primary = c; break; } }

    let fuenteFinal = '', medioFinal = '', campañaFinal = '', fechaFinal = '';

    if (primary) {
      const fila = primary.fila;
      if (primary.fuenteIdx !== null) fuenteFinal = fila[primary.fuenteIdx] || primary.defaultFuente || '';
      else fuenteFinal = primary.defaultFuente || fuenteFinal || '';

      if (primary.medioIdx !== null) medioFinal = fila[primary.medioIdx] || primary.defaultMedio || '';
      else medioFinal = primary.defaultMedio || medioFinal || '';

      if (primary.campIdx !== null) campañaFinal = fila[primary.campIdx] || '';

      if (primary.fechaIdx !== null && fila[primary.fechaIdx]) {
        fechaFinal = (fila[primary.fechaIdx] instanceof Date) ? fila[primary.fechaIdx] : new Date(fila[primary.fechaIdx]);
      }

      if ((!medioFinal || !campañaFinal) && primary.arr) {
        const foundSameSheet = buscarEnArrayParaCampos(primary.arr, doc, placa, correo, primary.campIdx, primary.medioIdx);
        if (foundSameSheet) {
          if (!campañaFinal && foundSameSheet.camp) campañaFinal = foundSameSheet.camp;
          if (!medioFinal && foundSameSheet.medio) medioFinal = foundSameSheet.medio;
        }
      }

      if (!medioFinal || !campañaFinal) {
        for (let c of candidates) {
          if (c === primary) continue;
          const found = buscarEnArrayParaCampos(c.arr, doc, placa, correo, c.campIdx, c.medioIdx);
          if (found) {
            if (!campañaFinal && found.camp) campañaFinal = found.camp;
            if (!medioFinal && found.medio) medioFinal = found.medio;
          }
          if (medioFinal && campañaFinal) break;
        }
      }

      if (!medioFinal && primary.defaultMedio) medioFinal = primary.defaultMedio;
    }

    const totalConteo = conteos.reduce((s, v) => s + v, 0);
    return [...conteos, totalConteo, fuenteFinal, medioFinal, campañaFinal, fechaFinal || ''];
  });

  hoja.getRange(2, 17, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);

  hoja.getRange("Z1").setValue("Fuente");
  hoja.getRange("AA1").setValue("Medio");
  hoja.getRange("AB1").setValue("Campaña");
  hoja.getRange("AC1").setValue("Fecha");
  hoja.getRange(2, 29, resultadosFinales.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");

  SpreadsheetApp.getUi().alert("Datos listos :)");
}
