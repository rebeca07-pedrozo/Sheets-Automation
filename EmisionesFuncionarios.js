// --- CONSTANTES ---
const HOJA_LEADS_TOTAL = "LEADS TOTAL";
const COLUMNA_CC_BUSCAR = 10; // J
const COLUMNA_CORREO1_BUSCAR = 12; // L
const COLUMNA_CORREO2_BUSCAR = 13; // M
const COLUMNA_CELULAR_BUSCAR = 14; // N
const COLUMNA_FECHA_BASE = 8; // H

// Columnas de LEADS TOTAL
const COL_LEADS_TOTAL_CC = 2; // B
const COL_LEADS_TOTAL_CORREO = 3; // C
const COL_LEADS_TOTAL_CELULAR = 4; // D
const COL_LEADS_TOTAL_FUENTE = 5; // E
const COL_LEADS_TOTAL_MEDIO = 6; // F
const COL_LEADS_TOTAL_CAMPANA = 7; // G
const COL_LEADS_TOTAL_PRODUCTO = 8; // H
const COL_LEADS_TOTAL_FECHA_FORM = 9; // I

// Columnas de RESULTADOS en la hoja del usuario
const COLUMNA_P_CC_ENCONTRADO = 16; // P
const COLUMNA_Q_CORREO1_ENCONTRADO = 17; // Q
const COLUMNA_R_CORREO2_ENCONTRADO = 18; // R
const COLUMNA_S_CELULAR_ENCONTRADO = 19; // S
const COLUMNA_T_SUMA_TOTAL = 20; // T
const COLUMNA_U_FUENTE = 21; // U
const COLUMNA_V_MEDIO = 22; // V
const COLUMNA_W_CAMPANA = 23; // W
const COLUMNA_X_PRODUCTO = 24; // X
const COLUMNA_Y_FECHA_FORM = 25; // Y
const COLUMNA_Z_DIF_FECHAS = 26; // Z


// 1) MENÚ
/**
 * Crea un menú personalizado al abrir la hoja de cálculo.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Automatizaciones')
      .addItem('🔍 Buscar y Normalizar Datos', 'iniciarNormalizacion')
      .addToUi();
}

/**
 * Solicita el nombre de la hoja al usuario e inicia el proceso.
 */
function iniciarNormalizacion() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const respuesta = ui.prompt(
        'Iniciar Normalización',
        'Por favor, ingresa el nombre de la hoja donde se ejecutará la automatización:',
        ui.ButtonSet.OK_CANCEL
    );

    if (respuesta.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert('Proceso cancelado por el usuario.');
      return;
    }

    const nombreHojaUsuario = respuesta.getResponseText().trim();
    if (!nombreHojaUsuario) {
      ui.alert('Error', 'El nombre de la hoja no puede estar vacío.', ui.ButtonSet.OK);
      return;
    }
    
    ejecutarNormalizacion(nombreHojaUsuario);

  } catch (e) {
    ui.alert('⚠️ Error General', 'Ocurrió un error inesperado: ' + e.message);
  }
}

// 4) PRIORIDAD DE BÚSQUEDA Y 5) ESCRITURA DE RESULTADOS
/**
 * Ejecuta la lógica principal de búsqueda y escritura.
 * @param {string} nombreHojaUsuario - El nombre de la hoja a procesar.
 */
function ejecutarNormalizacion(nombreHojaUsuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Validar si la hoja existe
  const hojaUsuario = ss.getSheetByName(nombreHojaUsuario);
  if (!hojaUsuario) {
    ui.alert('Error de Hoja', `La hoja "${nombreHojaUsuario}" no fue encontrada. Por favor, verifica el nombre.`);
    return;
  }
  
  const hojaLeadsTotal = ss.getSheetByName(HOJA_LEADS_TOTAL);
  if (!hojaLeadsTotal) {
    ui.alert('Error de Hoja', `La hoja de referencia "${HOJA_LEADS_TOTAL}" no fue encontrada.`);
    return;
  }

  // Validar si la hoja de usuario tiene datos (ignorando el encabezado)
  const ultimaFilaUsuario = hojaUsuario.getLastRow();
  if (ultimaFilaUsuario < 2) {
    ui.alert('Advertencia', `La hoja "${nombreHojaUsuario}" no tiene datos para procesar (solo encabezados o está vacía).`);
    return;
  }
  
  // Obtener todos los datos de ambas hojas
  // Hoja del usuario (desde la fila 2)
  const rangoDatosUsuario = hojaUsuario.getRange(2, 1, ultimaFilaUsuario - 1, COLUMNA_CELULAR_BUSCAR);
  const datosUsuario = rangoDatosUsuario.getValues();
  
  // Hoja LEADS TOTAL (desde la fila 2)
  const ultimaFilaLeads = hojaLeadsTotal.getLastRow();
  if (ultimaFilaLeads < 2) {
    ui.alert('Advertencia', `La hoja "${HOJA_LEADS_TOTAL}" no tiene datos de referencia.`);
    return;
  }
  
  const rangoDatosLeads = hojaLeadsTotal.getRange(2, 1, ultimaFilaLeads - 1, COLUMNA_LEADS_TOTAL_FECHA_FORM);
  const datosLeads = rangoDatosLeads.getValues();
  
  // Crear un mapa de datos de LEADS TOTAL normalizados para búsqueda eficiente
  const mapaLeads = crearMapaNormalizado(datosLeads);

  // Array para almacenar los resultados a escribir
  const resultados = [];

  // Procesar cada fila de la hoja del usuario
  datosUsuario.forEach((filaUsuario, indice) => {
    let resultadoFila = new Array(11).fill(''); // U-Y (5) + P-T (5) + Z (1) = 11 columnas
    
    // Extracción y normalización de datos de la fila del usuario
    const ccUsuario = filaUsuario[COLUMNA_CC_BUSCAR - 1];
    const correo1Usuario = filaUsuario[COLUMNA_CORREO1_BUSCAR - 1];
    const correo2Usuario = filaUsuario[COLUMNA_CORREO2_BUSCAR - 1];
    const celularUsuario = filaUsuario[COLUMNA_CELULAR_BUSCAR - 1];
    const fechaBaseUsuario = filaUsuario[COLUMNA_FECHA_BASE - 1]; // Columna H (Fecha)

    const ccNormalizado = normalizarCC(ccUsuario);
    const correo1Normalizado = normalizarCorreo(correo1Usuario);
    const correo2Normalizado = normalizarCorreo(correo2Usuario);
    const celularNormalizado = normalizarCelular(celularUsuario);
    
    // Variables para resultados P, Q, R, S
    let P = 0, Q = 0, R = 0, S = 0;
    let datosCoincidencia = null;

    // 4) PRIORIDAD DE BÚSQUEDA
    
    // 1. Búsqueda por CC
    if (ccNormalizado && mapaLeads.cc.has(ccNormalizado)) {
      datosCoincidencia = mapaLeads.cc.get(ccNormalizado);
      P = 1;
    } 
    // 2. Búsqueda por Correo1
    else if (correo1Normalizado && mapaLeads.correo.has(correo1Normalizado)) {
      datosCoincidencia = mapaLeads.correo.get(correo1Normalizado);
      Q = 1;
    }
    // 3. Búsqueda por Correo2
    else if (correo2Normalizado && mapaLeads.correo.has(correo2Normalizado)) {
      datosCoincidencia = mapaLeads.correo.get(correo2Normalizado);
      R = 1;
    }
    // 4. Búsqueda por Celular
    else if (celularNormalizado && mapaLeads.celular.has(celularNormalizado)) {
      datosCoincidencia = mapaLeads.celular.get(celularNormalizado);
      S = 1;
    }

    // 5) ESCRITURA DE RESULTADOS
    
    // Escribir P, Q, R, S y T
    const T = P + Q + R + S;
    resultadoFila[0] = P;
    resultadoFila[1] = Q;
    resultadoFila[2] = R;
    resultadoFila[3] = S;
    resultadoFila[4] = T;
    
    // Escribir U, V, W, X, Y (Fuente, Medio, Campaña, Producto, Fecha formulario)
    let fechaFormularioEncontrada = null;
    
    if (datosCoincidencia) {
      const fuente = datosCoincidencia[COL_LEADS_TOTAL_FUENTE - 1];
      const medio = datosCoincidencia[COL_LEADS_TOTAL_MEDIO - 1];
      const campana = datosCoincidencia[COL_LEADS_TOTAL_CAMPANA - 1];
      const producto = datosCoincidencia[COL_LEADS_TOTAL_PRODUCTO - 1];
      fechaFormularioEncontrada = datosCoincidencia[COL_LEADS_TOTAL_FECHA_FORM - 1];

      resultadoFila[5] = fuente; // U
      resultadoFila[6] = medio; // V
      resultadoFila[7] = campana; // W
      resultadoFila[8] = producto; // X
      resultadoFila[9] = fechaFormularioEncontrada; // Y
    }
    
    // Escribir Z (Diferencia de fechas)
    if (fechaBaseUsuario instanceof Date && fechaFormularioEncontrada instanceof Date) {
      const difTiempo = fechaBaseUsuario.getTime() - fechaFormularioEncontrada.getTime();
      const difDias = Math.floor(difTiempo / (1000 * 60 * 60 * 24));
      resultadoFila[10] = difDias; // Z
    } else {
      resultadoFila[10] = '';
    }

    resultados.push(resultadoFila);
  });
  
  // Escribir todos los resultados al final
  if (resultados.length > 0) {
    hojaUsuario.getRange(2, COLUMNA_P_CC_ENCONTRADO, resultados.length, resultados[0].length).setValues(resultados);
  }
  
  ui.alert('✅ Proceso Finalizado', `Se procesaron ${resultados.length} filas en la hoja "${nombreHojaUsuario}".`);
}


/**
 * Crea un mapa de búsqueda eficiente a partir de los datos de LEADS TOTAL.
 * @param {Array<Array<any>>} datosLeads - Los datos de la hoja LEADS TOTAL.
 * @returns {{cc: Map<string, Array<any>>, correo: Map<string, Array<any>>, celular: Map<string, Array<any>>}}
 */
function crearMapaNormalizado(datosLeads) {
  const mapaCC = new Map();
  const mapaCorreo = new Map();
  const mapaCelular = new Map();
  
  datosLeads.forEach(fila => {
    // Normalizar valores de LEADS TOTAL
    const ccNormalizado = normalizarCC(fila[COL_LEADS_TOTAL_CC - 1]);
    const correoNormalizado = normalizarCorreo(fila[COL_LEADS_TOTAL_CORREO - 1]);
    const celularNormalizado = normalizarCelular(fila[COL_LEADS_TOTAL_CELULAR - 1]);
    
    // Almacenar en los mapas
    if (ccNormalizado) {
      if (!mapaCC.has(ccNormalizado)) { // Tomar la primera coincidencia (si hay duplicados)
        mapaCC.set(ccNormalizado, fila);
      }
    }
    if (correoNormalizado) {
      if (!mapaCorreo.has(correoNormalizado)) {
        mapaCorreo.set(correoNormalizado, fila);
      }
    }
    if (celularNormalizado) {
      if (!mapaCelular.has(celularNormalizado)) {
        mapaCelular.set(celularNormalizado, fila);
      }
    }
  });
  
  return { cc: mapaCC, correo: mapaCorreo, celular: mapaCelular };
}


// 3) NORMALIZACIÓN OBLIGATORIA
// La comparación debe ser insensible a mayúsculas/minúsculas y a signos.
/**
 * Normaliza la cédula: quita puntos y otros signos.
 * @param {any} valor - El valor de la cédula.
 * @returns {string | null} El valor normalizado o null si es vacío.
 */
function normalizarCC(valor) {
  if (!valor) return null;
  let s = String(valor).trim();
  // Quitar puntos y cualquier carácter que no sea dígito
  s = s.replace(/[^0-9]/g, '');
  return s.toLowerCase(); // Por si acaso, aunque un CC suelen ser solo números
}

/**
 * Normaliza el correo: pasa a minúsculas, quita espacios y quita signos diacríticos.
 * @param {any} valor - El valor del correo.
 * @returns {string | null} El valor normalizado o null si es vacío.
 */
function normalizarCorreo(valor) {
  if (!valor) return null;
  let s = String(valor).trim().toLowerCase();
  // Eliminar acentos y otros signos (insensible a signos)
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  // Eliminar cualquier cosa que no sea un caracter de correo válido (letras, números, @, ., _, -)
  // s = s.replace(/[^a-z0-9@\.\_\-]/g, ''); // Ejemplo de una normalización más estricta
  return s;
}

/**
 * Normaliza el celular: usa solo números.
 * @param {any} valor - El valor del celular.
 * @returns {string | null} El valor normalizado o null si es vacío.
 */
function normalizarCelular(valor) {
  if (!valor) return null;
  let s = String(valor).trim();
  // Usar solo números
  s = s.replace(/[^0-9]/g, '');
  return s;
}