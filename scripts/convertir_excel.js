const xlsx = require("xlsx");
const fs = require("fs");

// Leer el archivo Excel
const workbook = xlsx.readFile("Productos.xlsx");

// Tomar la primera hoja
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convertir a JSON
const productos = xlsx.utils.sheet_to_json(sheet);

// Guardar como productos.json
fs.writeFileSync("productos.json", JSON.stringify(productos, null, 2));

console.log(`Se generaron ${productos.length} productos en productos.json`);