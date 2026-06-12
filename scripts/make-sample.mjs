// Genera samples/ejemplo.xlsx con datos ficticios para probar la herramienta.
// Uso: npm run sample

import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = resolve(root, "samples");
const outPath = resolve(outDir, "ejemplo.xlsx");

const productos = [
  "Camiseta básica",
  "Sudadera con capucha",
  "Taza cerámica",
  "Libreta A5",
  "Mochila campus",
  "Botella térmica",
];
const comerciales = ["Begoña Ruiz", "Iñigo Díaz", "María Solís", "Andrés Peña"];

const ventas = [["Fecha", "Producto", "Unidades", "Precio unitario (€)", "Total (€)", "Comercial"]];
for (let i = 0; i < 28; i++) {
  const unidades = 3 + ((i * 7) % 40);
  const precio = Math.round((5 + ((i * 13) % 30) + 0.95) * 100) / 100;
  ventas.push([
    new Date(2026, 0, 5 + i),
    productos[i % productos.length],
    unidades,
    precio,
    Math.round(unidades * precio * 100) / 100,
    comerciales[i % comerciales.length],
  ]);
}

const resumen = [
  ["Indicador", "Valor"],
  ["Ventas totales (€)", 15234.5],
  ["Unidades vendidas", 1280],
  ["Ticket medio (€)", 11.9],
  ["Mejor producto", "Sudadera con capucha"],
  ["Observaciones", "Datos ficticios generados para pruebas.\nIncluyen acentos: año, señal, métrica."],
];

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(ventas, { cellDates: true }), "Ventas");
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(resumen), "Resumen");
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), "Vacía");

mkdirSync(outDir, { recursive: true });
// La build ESM de SheetJS no enlaza `fs` por sí sola: se escribe el buffer a mano.
const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", cellDates: true });
writeFileSync(outPath, buffer);
console.log(`Archivo de ejemplo creado: ${outPath}`);
