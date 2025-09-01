function separarDatosFlexible() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();

  hoja.getRange(1, 21, 1, 5).setValues([["Nombre", "Tipo Doc", "Número Doc", "Placa", "Correo"]]);

  const rango = hoja.getRange("B2:B" + ultimaFila);
  const valores = rango.getValues();

  const salida = valores.map(([texto]) => {
    if (!texto) return ["", "", "", "", ""];

    const correoMatch = texto.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    const correo = correoMatch ? correoMatch[0] : "";

    
    const docMatch = texto.match(/\b([A-Z]{2})?\s*(\d{5,})\b/i);
    const tipoDoc = docMatch && docMatch[1] ? docMatch[1].toUpperCase() : "";
    const numDoc = docMatch && docMatch[2] ? docMatch[2] : "";

    const placaMatch = texto.match(/\b([A-Z]{3}\d{2,4}[A-Z]?)\b/i);
    const placa = placaMatch ? placaMatch[1] : "";

    let nombre = texto;
    if (correo) nombre = nombre.replace(correo, "");
    if (placa) nombre = nombre.replace(placa, "");
    if (docMatch && docMatch[0]) nombre = nombre.replace(docMatch[0], "");

    nombre = nombre.split("//")[0];
    nombre = nombre.replace(/[-_]+.*/, "");
    nombre = nombre.replace(/\d{2,}.*/, ""); 
    nombre = nombre.replace(/\b(PRUEBA|CALI|BUCARAMANGA)\b/gi, ""); 
    nombre = nombre.replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]/g, ""); 
    nombre = nombre.replace(/\s+/g, " ").trim();

    return [nombre, tipoDoc, numDoc, placa, correo];
  });

  hoja.getRange(2, 21, salida.length, 5).setValues(salida);
}
