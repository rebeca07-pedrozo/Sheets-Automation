//KEY_PRIMAS = 'Archivo_de_primas_'
//KEY_PERSONAS = 'Archivo_de_personas_'
const KEY_VALORES = 'Archivo_de_primas_'; // CLAVE CORREGIDA
const KEY_PERSONAS = 'Archivo_de_personas_'; // CLAVE CORREGIDA

const binaryData = $input.item.binary;

if (!binaryData || !binaryData[KEY_VALORES] || !binaryData[KEY_PERSONAS]) {
    throw new Error("No se encontraron los archivos binarios subidos.");
}
const valoresCSV = Buffer.from(binaryData[KEY_VALORES].data, 'base64').toString('utf8');
const personasCSV = Buffer.from(binaryData[KEY_PERSONAS].data, 'base64').toString('utf8');

function csvToJson(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length === 0) return [];
    
    const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
    const data = [];

    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',');
        if (values.length !== headers.length) continue; 
        
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
            obj[headers[j]] = values[j] ? values[j].trim().replace(/"/g, '') : '';
        }
        data.push(obj);
    }
    return data;
}

const valores = csvToJson(valoresCSV);
const personas = csvToJson(personasCSV);

const valoresMap = new Map();
valores.forEach(v => {
    if (v.numero_poliza) {
        valoresMap.set(String(v.numero_poliza), v.PRIMA);
    }
});

const grouped = new Map();

personas.forEach(p => {
    const poliza = String(p.numero_poliza);

    if (!grouped.has(poliza)) {
        grouped.set(poliza, {
            numero_poliza: poliza,
            clave_agente: p.clave_agente,
            codigo_producto: p.codigo_producto,
            fecha_emision: p.fecha_emision,
            nombre_producto: p.nombre_producto,
            nombre_opcion_poliza: p.nombre_opcion_poliza,
            tipo_documento: p.tipo_documento,
            NUMERO_DOCUMENTO: new Set(),
            NOMBRE: new Set(),
            CORREO_1: new Set(),
            CELULAR_1: new Set(),
            FECHA_PROCESO: p.FECHA_PROCESO
        });
    }

    const item = grouped.get(poliza);
    if (p.NUMERO_DOCUMENTO) item.NUMERO_DOCUMENTO.add(String(p.NUMERO_DOCUMENTO));
    if (p.NOMBRE) item.NOMBRE.add(String(p.NOMBRE));
    if (p.CORREO_1) item.CORREO_1.add(String(p.CORREO_1));
    if (p.CELULAR_1) item.CELULAR_1.add(String(p.CELULAR_1));
});


function expandirColumnaJS(dfItem, columnaBaseSet, prefijo) {
    const lista = Array.from(columnaBaseSet).filter(s => s.trim() !== '').sort();
    
    for (let i = 0; i < lista.length; i++) {
        dfItem[`${prefijo}_${i + 1}`] = lista[i];
    }
    return dfItem;
}

let resultadoFinal = Array.from(grouped.values()).map(agrupado => {
    
    const prima = valoresMap.get(agrupado.numero_poliza) || '';

    let nuevoItem = {
        numero_poliza: agrupado.numero_poliza,
        clave_agente: agrupado.clave_agente,
        codigo_producto: agrupado.codigo_producto,
        fecha_emision: agrupado.fecha_emision,
        nombre_producto: agrupado.nombre_producto,
        nombre_opcion_poliza: agrupado.nombre_opcion_poliza,
        tipo_documento: agrupado.tipo_documento,
        
        Prima_totalizada: prima,
        FECG: agrupado.FECHA_PROCESO,
        
        CORREO_1_temp: Array.from(agrupado.CORREO_1).filter(Boolean).sort().join(', '),

        DOC_SET: agrupado.NUMERO_DOCUMENTO,
        CORREO_SET: agrupado.CORREO_1
    };
    
    nuevoItem = expandirColumnaJS(nuevoItem, nuevoItem.DOC_SET, 'Numero_documento');
    nuevoItem = expandirColumnaJS(nuevoItem, nuevoItem.CORREO_SET, 'CORREO');
    
    return nuevoItem;
});

const columnas_finales = [
    'emisiones',
    'revision agente',
    'codigo_producto',
    'clave_agente',
    'numero_poliza',
    'fecha_emision',
    'nombre_producto',
    'nombre_opcion_poliza',
    'CORREO_1', 
    'Numero_documento_1',
    'Numero_documento_2',
    'Prima_totalizada',
    'tipo_documento',
    'FECG'
];

resultadoFinal = resultadoFinal.map(item => {
    const finalItem = {};
    
    columnas_finales.forEach(col => {
        if (col === 'CORREO_1') {
            finalItem[col] = item.CORREO_1_temp || '';
        } else if (item.hasOwnProperty(col)) {
            finalItem[col] = item[col];
        } else {
            finalItem[col] = '';
        }
    });

    return finalItem;
});

return resultadoFinal.map(item => ({json: item}));