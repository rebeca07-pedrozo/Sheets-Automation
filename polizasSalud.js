const KEY_PRIMAS = 'Archivo_de_primas_';
const KEY_PERSONAS = 'Archivo_de_personas_';

const primasBinary = items[0].binary[KEY_PRIMAS];
const personasBinary = items[0].binary[KEY_PERSONAS];

const primasCSV = Buffer.from(primasBinary.data, 'base64').toString('utf8');
const personasCSV = Buffer.from(personasBinary.data, 'base64').toString('utf8');

// =========================
// PARSER CSV A JSON PURO JS
// =========================
function csvToJson(csv) {
  const lines = csv.split(/\r?\n/);
  const headers = lines.shift().split(",").map(h => h.trim());

  return lines
    .filter(l => l.trim().length > 0)
    .map(line => {
      const cols = line.split(",").map(c => c.trim());
      const obj = {};
      headers.forEach((h, i) => obj[h] = cols[i] || "");
      return obj;
    });
}

const primas = csvToJson(primasCSV);
const personas = csvToJson(personasCSV);

// ====================================
// AGRUPAR PERSONAS POR numero_poliza
// ====================================
const grupos = {};

for (const row of personas) {
  const poliza = row.numero_poliza;

  if (!grupos[poliza]) {
    grupos[poliza] = {
      numero_poliza: poliza,
      clave_agente: row.clave_agente,
      codigo_producto: row.codigo_producto,
      fecha_emision: row.fecha_emision,
      nombre_producto: row.nombre_producto,
      nombre_opcion_poliza: row.nombre_opcion_poliza,
      tipo_documento: row.tipo_documento,
      FECHA_PROCESO: row.FECHA_PROCESO,
      NUMERO_DOCUMENTO: new Set(),
      NOMBRE: new Set(),
      CORREO_1: new Set(),
      CELULAR_1: new Set()
    };
  }

  if (row.NUMERO_DOCUMENTO) grupos[poliza].NUMERO_DOCUMENTO.add(String(row.NUMERO_DOCUMENTO));
  if (row.NOMBRE) grupos[poliza].NOMBRE.add(String(row.NOMBRE));
  if (row.CORREO_1) grupos[poliza].CORREO_1.add(String(row.CORREO_1));
  if (row.CELULAR_1) grupos[poliza].CELULAR_1.add(String(row.CELULAR_1));
}

const agrupado = Object.values(grupos).map(r => ({
  ...r,
  NUMERO_DOCUMENTO: [...r.NUMERO_DOCUMENTO].sort().join(", "),
  NOMBRE: [...r.NOMBRE].sort().join(", "),
  CORREO_1: [...r.CORREO_1].sort().join(", "),
  CELULAR_1: [...r.CELULAR_1].sort().join(", "),
}));

// =========================
// MERGE CON PRIMAS
// =========================
const mapPrimas = Object.fromEntries(primas.map(p => [p.numero_poliza, p.PRIMA]));

for (const row of agrupado) {
  row.Prima_totalizada = mapPrimas[row.numero_poliza] || "";
  row.FECG = row.FECHA_PROCESO;
}

// ====================================
// EXPANDIR COLUMNAS
// ====================================
function expandir(colBase, prefijo) {
  for (const row of agrupado) {
    const lista = row[colBase] ? row[colBase].split(",").map(i => i.trim()).filter(Boolean) : [];
    lista.forEach((val, i) => {
      row[`${prefijo}_${i + 1}`] = val;
    });
  }
}

expandir("NUMERO_DOCUMENTO", "Numero_documento");
expandir("CORREO_1", "CORREO");

// =========================
// COLUMNAS FINALES
// =========================
const columnas_finales = [
  "emisiones",
  "revision agente",
  "codigo_producto",
  "clave_agente",
  "numero_poliza",
  "fecha_emision",
  "nombre_producto",
  "nombre_opcion_poliza",
  "CORREO_1",
  "Numero_documento_1",
  "Numero_documento_2",
  "Prima_totalizada",
  "tipo_documento",
  "FECG"
];

const resultado_final = agrupado.map(row => {
  const out = {};
  for (const col of columnas_finales) {
    out[col] = row[col] || "";
  }
  return out;
});

return resultado_final.map(r => ({ json: r }));
