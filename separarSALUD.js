function extraerDatosDeSalud() {
  const SHEET_NAME = "SALUD";
  const START_ROW = 2; 
  const INPUT_COLUMN = 'B'; 
  const OUTPUT_START_COLUMN = 'U'; 
  
  const HEADERS = [
    "Nombre", 
    "Tipo de Documento", 
    "Número de Documento", 
    "Correo Electrónico", 
    "Celular",
    "Plan"
  ];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return;  
    }
    
    const outputStartColIndex = sheet.getRange(OUTPUT_START_COLUMN + '1').getColumn();
    sheet.getRange(1, outputStartColIndex, 1, HEADERS.length).setValues([HEADERS]);

    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) return;
    
    const inputRange = sheet.getRange(`${INPUT_COLUMN}${START_ROW}:${INPUT_COLUMN}${lastRow}`);
    const inputValues = inputRange.getValues();
    
    const results = [];
    
    const regexDocumentoOld = /(\d+)\s+(\d+_\w+)/; 
    const regexDocumentoSimple = /\b(CC|CE|TI|PA|RC|NIT|CD)\s+(\d{7,15})\b/i;
    const regexPlan = /(\d_Plan_\w+)/;
    const regexCorreo = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/; 
    
    inputValues.forEach((row) => {
      const dataString = String(row[0] || '').trim(); 
      let numeroDocumento = '';
      let tipoDocumento = '';
      let documentBlockMatch = null; 
      
      if (dataString === '') {
        results.push(["", "", "", "", "", ""]); 
        return; 
      }
      
      const matchDocOld = dataString.match(regexDocumentoOld);
      const matchDocSimple = dataString.match(regexDocumentoSimple);

      if (matchDocOld) {
          numeroDocumento = matchDocOld[1];
          tipoDocumento = matchDocOld[2]; 
          documentBlockMatch = matchDocOld[0];
      } else if (matchDocSimple) {
          tipoDocumento = matchDocSimple[1]; 
          numeroDocumento = matchDocSimple[2];
          documentBlockMatch = matchDocSimple[0];
      }
      
      let nombre = dataString;
      if (numeroDocumento && documentBlockMatch) {
          const docStart = dataString.indexOf(documentBlockMatch);
          if (docStart !== -1) {
              nombre = dataString.substring(0, docStart).trim();
          }
      }
      nombre = nombre.replace(/\{[^}]+\}/g, '').trim(); 
      const nombreParts = nombre.split(/\s+/);
      const nombreLimpio = [];
      for(const part of nombreParts) {
          if (part.match(/\d+_\w+/) || part.includes('Plan')) break; 
          nombreLimpio.push(part);
      }
      nombre = nombreLimpio.join(' ').trim();

      const matchPlan = dataString.match(regexPlan);
      let plan = matchPlan ? matchPlan[1] : ''; 

      const matchCorreo = dataString.match(regexCorreo);
      const correo = matchCorreo ? matchCorreo[1] : '';
      
      let celular = '';
      const parts = dataString.split(/[\s,]/).filter(p => p.length > 0);
      for (let i = parts.length - 1; i >= 0; i--) {
        const part = parts[i];
        if (part.match(/^\d{7,15}$/) && part !== numeroDocumento) {
          celular = part;
          break;
        }
      }
      
      if (tipoDocumento) {
          tipoDocumento = tipoDocumento.replace(/^\d+_/, '').toUpperCase(); 
      }
      if (plan) {
          plan = plan.replace(/^\d_/, ''); 
      }

      results.push([
        nombre,
        tipoDocumento,
        numeroDocumento,
        correo, 
        celular, 
        plan 
      ]);
    });

    const numRows = results.length;
    const numCols = HEADERS.length;
    
    if (numRows > 0) {
      const outputRange = sheet.getRange(START_ROW, outputStartColIndex, numRows, numCols);
      outputRange.setValues(results);
    }
    
  } catch (e) {
    Logger.log(`Error al procesar los datos: ${e.toString()}`);
  }
}
