var HOJAS_ORIGEN = [
    "Hogar Digital", "Venta_Asistencias", "Protección Créditos", "Soat", 
    "Salud", "AUTOS", "Tranquilidad Vida", "Uso Banca Servicios", 
    "Venta Banca Servicios", "Seminarios Web", "Conocimiento del Cliente", 
    "Acceso clientes"
];

var NOMBRE_HOJA_DESTINO = "Consolidado_ventas_formula_AUTOMATIZADO";
var NOMBRE_COLUMNA_CLAVE = "Intención"; 
var FILA_ENCABEZADO = 1;


function applyRowBorders(sheet, startRow, numRows, numCols) {
    var borderRange = sheet.getRange(startRow, 1, numRows, numCols);
    var style = SpreadsheetApp.BorderStyle.SOLID;
    var color = 'black';
    borderRange.setBorder(true, true, true, true, true, true, color, style);
}


function procesarDatosNuevos() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var hojaDestino = ss.getSheetByName(NOMBRE_HOJA_DESTINO);

    if (!hojaDestino) {
        ui.alert('Error Fatal', 'La hoja de destino "' + NOMBRE_HOJA_DESTINO + '" no fue encontrada. Por favor, créala o verifica su nombre.', ui.ButtonSet.OK);
        Logger.log('ERROR: La hoja de destino no existe.');
        return;
    }

    var filasParaConsolidar = [];
    var scriptProperties = PropertiesService.getScriptProperties();
    var totalFilasConsolidadas = 0;

    HOJAS_ORIGEN.forEach(function(nombreHoja) {
        var hoja = ss.getSheetByName(nombreHoja);

        if (!hoja) {
            Logger.log('WARNING: La hoja de origen "' + nombreHoja + '" no fue encontrada y fue saltada.');
            return;
        }

        var lastRow = hoja.getLastRow();

        if (lastRow < 2) {
            Logger.log('Procesada hoja: ' + nombreHoja + ' | Filas nuevas: 0 (Hoja vacía o solo con encabezado).');
            return;
        }

        var encabezados = hoja.getRange(FILA_ENCABEZADO, 1, 1, hoja.getLastColumn()).getValues()[0];
        var intencionColIndex = encabezados.indexOf(NOMBRE_COLUMNA_CLAVE);

        if (intencionColIndex === -1) {
            Logger.log('WARNING: Columna "' + NOMBRE_COLUMNA_CLAVE + '" no encontrada en la hoja: ' + nombreHoja + '. Saltando.');
            return;
        }

        var key = 'lastProcessedRow_' + nombreHoja.replace(/\s/g, '_');
        var startRow = (scriptProperties.getProperty(key)) ? parseInt(scriptProperties.getProperty(key)) + 1 : 2;

        if (lastRow >= startRow) {
            var numRowsToProcess = lastRow - startRow + 1;
            var lastCol = hoja.getLastColumn();

            var data = hoja.getRange(startRow, 1, numRowsToProcess, lastCol).getValues();
            var filasAgregadasEnHoja = 0;

            data.forEach(function(row) {
                if (!row[0]) return;
                var intencion = row[intencionColIndex];
                if (typeof intencion !== 'string' && typeof intencion !== 'number') return;
                var intencionNormalizada = String(intencion).trim().toUpperCase();
                var cumpleCondicion = intencionNormalizada.includes("VENTA") ||
                                      (intencionNormalizada === "INCENTIVO USO ASISTENCIA BANCASEGUROS");

                if (cumpleCondicion) {
                    filasParaConsolidar.push(row);
                    filasAgregadasEnHoja++;
                }
            });

            if (numRowsToProcess > 0) {
                scriptProperties.setProperty(key, lastRow.toString());
                Logger.log('Procesada hoja: ' + nombreHoja + ' | Filas nuevas: ' + filasAgregadasEnHoja + ' | Marcador actualizado a Fila ' + lastRow + '.');
                totalFilasConsolidadas += filasAgregadasEnHoja;
            }
        } else {
            Logger.log('Procesada hoja: ' + nombreHoja + ' | Filas nuevas: 0 (No hay data nueva).');
        }
    });

    if (filasParaConsolidar.length > 0) {
        var numFilasAgregadas = filasParaConsolidar.length;
        var startRowConsolidado = hojaDestino.getLastRow() + 1;
        var numCols = filasParaConsolidar[0].length;
        
        hojaDestino.getRange(startRowConsolidado, 1, numFilasAgregadas, numCols).setValues(filasParaConsolidar);
        
        applyRowBorders(hojaDestino, startRowConsolidado, numFilasAgregadas, numCols);

        var lastRowConsolidado = hojaDestino.getLastRow();
        if (lastRowConsolidado > FILA_ENCABEZADO) {
            hojaDestino.getRange(FILA_ENCABEZADO + 1, 1, lastRowConsolidado - FILA_ENCABEZADO, numCols)
                       .sort({column: 1, ascending: true});
            Logger.log('Data consolidada ordenada por fecha (A -> Z).');
        }
        
        Logger.log('--- CONSOLIDACIÓN FINALIZADA ---');
        Logger.log('Total de filas añadidas al consolidado: ' + totalFilasConsolidadas + '. Se aplicaron bordes.');
    } else {
        Logger.log('No se encontraron filas nuevas que cumplan las condiciones para consolidar.');
    }
}


function ejecucionInicialCompleta() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var hojaDestino = ss.getSheetByName(NOMBRE_HOJA_DESTINO);
    var scriptProperties = PropertiesService.getScriptProperties();

    if (!hojaDestino) {
        ui.alert('Error Fatal', 'La hoja de destino "' + NOMBRE_HOJA_DESTINO + '" no fue encontrada. Por favor, créala.', ui.ButtonSet.OK);
        return;
    }

    var filasABorrar = hojaDestino.getLastRow() - FILA_ENCABEZADO;
    if (filasABorrar > 0) {
        hojaDestino.deleteRows(FILA_ENCABEZADO + 1, filasABorrar);
        Logger.log('Consolidado limpiado. Borradas ' + filasABorrar + ' filas históricas.');
    }

    var primeraHoja = ss.getSheetByName(HOJAS_ORIGEN[0]);
    if (primeraHoja && hojaDestino.getLastRow() === 0) {
        var encabezadosRange = primeraHoja.getRange(FILA_ENCABEZADO, 1, 1, primeraHoja.getLastColumn());
        encabezadosRange.copyTo(hojaDestino.getRange(FILA_ENCABEZADO, 1));
        
        hojaDestino.getRange(FILA_ENCABEZADO, 1, 1, primeraHoja.getLastColumn()).setBorder(null, null, true, null, false, false);
        
        Logger.log('Encabezados copiados de ' + HOJAS_ORIGEN[0] + ' al Consolidado.');
    }

    scriptProperties.deleteAllProperties();
    Logger.log('Punteros de última fila (PropertiesService) borrados.');

    procesarDatosNuevos();

    ui.alert('Éxito', '¡La ejecución inicial de data histórica ha finalizado y los marcadores de última fila se han guardado!', ui.ButtonSet.OK);
}


function configurarDisparadorAutomatico() {
    var ui = SpreadsheetApp.getUi();

    var disparadores = ScriptApp.getProjectTriggers();
    disparadores.forEach(function(t) {
        if (t.getHandlerFunction() == 'procesarDatosNuevos') {
            ScriptApp.deleteTrigger(t);
            Logger.log('Disparador anterior eliminado.');
        }
    });

    ScriptApp.newTrigger('procesarDatosNuevos')
        .timeBased()
        .everyMinutes(5) 
        .create();

    ui.alert('Reviso de consolidado completado', 'El script ahora se ejecutará cada 5 minutos para procesar los nuevos datos de Salesforce.', ui.ButtonSet.OK);
}