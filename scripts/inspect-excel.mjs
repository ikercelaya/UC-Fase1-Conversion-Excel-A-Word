// Inspector de archivos Excel: imprime hojas, dimensiones y contenido.
// Uso: node scripts/inspect-excel.mjs <ruta.xlsx> [hoja] [filas]

import { readFileSync } from "node:fs";
import * as XLSX from "xlsx";

const [, , path, sheetFilter, maxRowsArg] = process.argv;
if (!path) {
  console.error("Uso: node scripts/inspect-excel.mjs <ruta.xlsx> [hoja] [filas]");
  process.exit(1);
}
const maxRows = Number(maxRowsArg ?? 40);

const workbook = XLSX.read(readFileSync(path), { type: "buffer", cellDates: true });
console.log(`Hojas (${workbook.SheetNames.length}): ${workbook.SheetNames.join(" | ")}`);

for (const name of workbook.SheetNames) {
  if (sheetFilter && name !== sheetFilter) continue;
  const worksheet = workbook.Sheets[name];
  const ref = worksheet["!ref"] ?? "(vacía)";
  console.log(`\n=== Hoja: ${name} · rango ${ref} ===`);
  if (!worksheet["!ref"]) continue;
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: true,
  });
  rows.slice(0, maxRows).forEach((row, index) => {
    const cells = row.map((cell) => String(cell).replace(/\n/g, "\\n"));
    // recorta celdas largas para que quepa en pantalla
    const compact = cells.map((c) => (c.length > 24 ? `${c.slice(0, 24)}…` : c));
    console.log(`${String(index + 1).padStart(4)} | ${compact.join(" ┊ ")}`);
  });
  if (rows.length > maxRows) console.log(`  … (${rows.length} filas en total)`);
}
