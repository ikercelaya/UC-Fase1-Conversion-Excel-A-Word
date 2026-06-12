import * as XLSX from "xlsx";

/** Máximo de filas por hoja que se vuelcan al Word (protege tiempo/memoria en Vercel). */
export const MAX_ROWS_PER_SHEET = 2000;

/** Máximo de caracteres por celda en el informe. */
const MAX_CELL_CHARS = 500;

/** Caracteres de control no imprimibles (excepto salto de línea y tabulador). */
const CONTROL_CHARS = new RegExp("[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F]", "g");

export interface SheetData {
  name: string;
  /** Filas con los valores ya formateados como texto, tal y como se muestran en Excel. */
  rows: string[][];
  /** Filas reales de la hoja antes de aplicar el recorte. */
  totalRows: number;
  columnCount: number;
  truncated: boolean;
}

export interface WorkbookData {
  sheets: SheetData[];
  totalRows: number;
}

/**
 * Lee un libro Excel (.xlsx, .xls, .xlsm o .csv) y devuelve todo su contenido
 * como texto formateado, hoja por hoja.
 */
export function parseWorkbook(buffer: Buffer): WorkbookData {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });

  const sheets: SheetData[] = workbook.SheetNames.map((name) => {
    const worksheet = workbook.Sheets[name];
    // header: 1 → matriz de filas; raw: false → valores con el formato visible en Excel
    const raw = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });
    const rows = normalize(raw);
    const totalRows = rows.length;
    const truncated = totalRows > MAX_ROWS_PER_SHEET;

    return {
      name,
      rows: truncated ? rows.slice(0, MAX_ROWS_PER_SHEET) : rows,
      totalRows,
      columnCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
      truncated,
    };
  });

  return {
    sheets,
    totalRows: sheets.reduce((sum, sheet) => sum + sheet.totalRows, 0),
  };
}

/** Convierte celdas a texto limpio y recorta columnas vacías al final, igualando el ancho de todas las filas. */
function normalize(raw: unknown[][]): string[][] {
  const rows = raw.map((row) => row.map(toCellText));

  let lastColumn = 0;
  for (const row of rows) {
    for (let i = row.length - 1; i >= 0; i--) {
      if (row[i] !== "") {
        lastColumn = Math.max(lastColumn, i + 1);
        break;
      }
    }
  }

  return rows.map((row) => {
    const out = row.slice(0, lastColumn);
    while (out.length < lastColumn) out.push("");
    return out;
  });
}

function toCellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  let text: string;
  if (value instanceof Date) {
    text = value.toLocaleDateString("es-ES");
  } else {
    text = String(value);
  }
  text = text.replace(/\r\n/g, "\n").replace(CONTROL_CHARS, "").trim();
  return text.length > MAX_CELL_CHARS ? `${text.slice(0, MAX_CELL_CHARS)}…` : text;
}
