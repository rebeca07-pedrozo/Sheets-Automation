function EmisionesAutosCruzados(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Si no se pasa nombre, usa este por defecto (cámbialo si tu hoja se llama distinto)
  if (!nombreHoja) nombreHoja = "Emisiones 17 nov"; 

  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    ui.alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  // --- 1. FUNCIONES DE LIMPIEZA AVANZADAS (Corrección JSON) ---
  
  // Función auxiliar para arreglar el formato {"0":"A", "1":"B"...}
  const extraerValorDeJson = (valor) => {
    let texto = String(valor || '');
    // Si empieza con llave { y parece un objeto JSON
    if (texto.trim().startsWith('{') && texto.includes(':')) {
      try {
        const obj = JSON.parse(texto);
        // Toma solo los valores (las letras) y los une
        return Object.values(obj).join(''); 
      } catch (e) {
        return texto; // Si falla, devuelve el texto original
      }
    }
    return texto;
  };

  // Limpia Documento: Extrae JSON -> Quita todo lo que no sea letra/número -> Minúsculas
  const limpiarDoc = d => {
    const textoReal = extraerValorDeJson(d);
    return textoReal.replace(/[^a-z0-9]/gi, '').toLowerCase();
  };

  // Limpia Placa: Extrae JSON -> Quita todo lo que no sea letra/número -> Minúsculas
  const limpiarPlaca = p => {
    const textoReal = extraerValorDeJson(p);
    return textoReal.replace(/[^a-z0-9]/gi, '').toLowerCase();
  };

  // Limpia Correo: Quita espacios y minúsculas
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();  


  // --- 2. LECTURA DE DATOS (Empezando desde Columna A) ---
  // getRange(fila, columna, numFilas, numColumnas) -> (2, 1...) es Columna A
  const datosEmisiones = hoja.getRange(2, 1, ultimaFila - 1, 20).getValues();


  // --- 3. CARGA DE "TOTAL LEADS - CRUZAR" ---
  const hojaLeads = ss.getSheetByName("TOTAL LEADS - CRUZAR");
  if (!hojaLeads) { ui.alert("Falta hoja TOTAL LEADS - CRUZAR"); return; }
  
  const datosLeads = hojaLeads.getDataRange().getValues().slice(1); 
  
  const mapTLC_CC = new Map(), mapTLC_Placa = new Map(), mapTLC_Correo = new Map();
  
  datosLeads.forEach(r => {
    // Índices TOTAL LEADS: B=1 (Doc), C=2 (Correo), E=4 (Placa)
    const cedula = r[1] ? limpiarDoc(r[1]) : ''; 
    const correo = r[2] ? limpiarCorreo(r[2]) : ''; 
    const placa = r[4] ? limpiarPlaca(r[4]) : '';  

    // Guardamos el primer dato encontrado (el más reciente)
    if (cedula && !mapTLC_CC.has(cedula)) mapTLC_CC.set(cedula, r);
    if (correo && !mapTLC_Correo.has(correo)) mapTLC_Correo.set(correo, r);
    if (placa && !mapTLC_Placa.has(placa)) mapTLC_Placa.set(placa, r);
  });


  // --- 4. CARGA DE "BASES" ---
  const hojaBases = ss.getSheetByName("BASES");
  if (!hojaBases) { ui.alert("Falta hoja BASES"); return; }
  const basesDatos = hojaBases.getDataRange().getValues().slice(1);
  
  // Ordenamos BASES por fecha (Columna C -> índice 2) Descendente
  basesDatos.sort((a, b) => {
    const fA = new Date(a[2]); const fB = new Date(b[2]);
    return (isNaN(fB) ? -1 : (isNaN(fA) ? 1 : fB - fA));
  });

  const mapBASES_CC = new Map(), mapBASES_Correo = new Map();
  basesDatos.forEach(r => {
    // Índices BASES: A=0 (Doc), B=1 (Correo)
    const cedula = r[0] ? limpiarDoc(r[0]) : '';
    const correo = r[1] ? limpiarCorreo(r[1]) : '';
    if (cedula && !mapBASES_CC.has(cedula)) mapBASES_CC.set(cedula, r);
    if (correo && !mapBASES_Correo.has(correo)) mapBASES_Correo.set(correo, r);
  });
  
  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || 'GMT-5', "yyyy-MM-dd HH:mm:ss");
  }


  // --- 5. CRUCE DE DATOS ---
  console.log("--- INICIANDO CRUCE ---");
  
  const resultadosFinales = datosEmisiones.map((r, i) => {
    
    // ÍNDICES DE LA HOJA 'EMISIONES' (A=0)
    // Columna A (Placa) -> Índice 0
    const placa = limpiarPlaca(r[0]); 
    // Columna L (Doc) -> Índice 11
    const doc = limpiarDoc(r[11]);     
    // Columna O (Correo) -> Índice 14 (Si tu correo está en N, cambia a 13)
    const correo = limpiarCorreo(r[14]); 

    // LOG DE CONTROL: Mira esto en "Registros de ejecución" para ver si limpia bien el JSON
    if (i === 0) {
        console.log(`FILA 1 | Placa Original: '${r[0]}' -> Limpia: '${placa}'`);
        console.log(`FILA 1 | Doc Original: '${r[11]}' -> Limpio: '${doc}'`);
    }

    let foundRowTLC = null;
    let foundRowBASES = null;

    // Prioridad 1: Buscar en TOTAL LEADS
    if (doc && mapTLC_CC.has(doc)) foundRowTLC = mapTLC_CC.get(doc);
    if (placa && !foundRowTLC && mapTLC_Placa.has(placa)) foundRowTLC = mapTLC_Placa.get(placa);
    if (correo && !foundRowTLC && mapTLC_Correo.has(correo)) foundRowTLC = mapTLC_Correo.get(correo);
    
    // Prioridad 2: Buscar en BASES
    if (!foundRowTLC) {
      if (doc && mapBASES_CC.has(doc)) foundRowBASES = mapBASES_CC.get(doc);
      if (correo && !foundRowBASES && mapBASES_Correo.has(correo)) foundRowBASES = mapBASES_Correo.get(correo);
    }
    
    // Contadores
    let countTLC_CC = (doc && mapTLC_CC.has(doc)) ? 1 : 0;
    let countTLC_Placa = (placa && mapTLC_Placa.has(placa)) ? 1 : 0;
    let countTLC_Correo = (correo && mapTLC_Correo.has(correo)) ? 1 : 0;
    let countBASES_CC = (doc && mapBASES_CC.has(doc)) ? 1 : 0;
    let countBASES_Correo = (correo && mapBASES_Correo.has(correo)) ? 1 : 0;

    let fuenteFinal = '', medioFinal = '', campañaFinal = '', adnameFinal = '', fechaFinal = '';

    if (foundRowTLC) {
      fuenteFinal = foundRowTLC[8] || '';   // Col I
      medioFinal = foundRowTLC[12] || '';   // Col M
      campañaFinal = foundRowTLC[9] || '';  // Col J
      adnameFinal = foundRowTLC[13] || '';  // Col N
      fechaFinal = formatearFecha(foundRowTLC[11]); // Col L (Fecha)
    } else if (foundRowBASES) {
      fuenteFinal = foundRowBASES[7] || 'BASES';
      medioFinal = foundRowBASES[3] || '';
      campañaFinal = foundRowBASES[6] || '';
      fechaFinal = formatearFecha(foundRowBASES[2]);
    }

    const conteos = [countTLC_CC, countTLC_Placa, countTLC_Correo, countBASES_CC, countBASES_Correo];
    const totalConteo = conteos.reduce((a, b) => a + b, 0);

    return [...conteos, totalConteo, totalConteo, fuenteFinal, medioFinal, campañaFinal, adnameFinal, fechaFinal];
  });

  const nuevosEncabezados = ["TOTAL LEADS CC", "TOTAL LEADS Placa", "TOTAL LEADS Correo", "Bases CC", "Bases Correo", "Total Leads", "Ventas", "Fuente", "Medio", "Campaña", "Adname", "Fecha Lead"];

  if (resultadosFinales.length > 0) {
    // Escribir en Columna Q (17)
    hoja.getRange(1, 17, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hoja.getRange(2, 17, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }
  
  SpreadsheetApp.getUi().alert("Proceso completado correctamente :)");
}