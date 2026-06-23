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

// ---------------------------------------------------------------------------
// Extracción específica del formato LaRUC (informes de radón)
//
// Los Excel del laboratorio tienen una hoja "Resultados" con bloques
// "RESULTADOS PARA INFORME" por cada slide: ID, DÍA INICIAL, DÍA FINAL,
// EXPOSICION, u(EXP), LD EXPOSICION, CONCENTRACION, u(CONC), LD CONCENTRACION.
// Esos son exactamente los valores que aparecen en la tabla de cada detector
// del informe de ensayo.
// ---------------------------------------------------------------------------

/** Identificador de detector, p. ej. HK3287 o HH8401. Excluye huecos ("0"). */
const DETECTOR_ID = /^[A-Za-z]{1,4}\d{2,6}$/;

export interface RadonSample {
  /** Referencia UC/expediente del informe, p. ej. P-25002-TRA-1. */
  expediente: string;
  id: string;
  /** Procedencia declarada en el bloque "RESULTADOS PARA INFORME". */
  procedencia: string;
  /** Bloque del que procede (p. ej. "SLIDE 1"), informativo. */
  slide: string;
  fechaColocacion: string;
  fechaRetirada: string;
  exposicion: string;
  incertidumbreExposicion: string;
  ldExposicion: string;
  concentracion: string;
  incertidumbreConcentracion: string;
  ldConcentracion: string;
}

export interface RadonData {
  sheetName: string;
  samples: RadonSample[];
}

interface ColumnMap {
  expediente: number;
  id: number;
  procedencia: number;
  diaInicial: number;
  diaFinal: number;
  exposicion: number;
  uExp: number;
  ldExp: number;
  concentracion: number;
  uConc: number;
  ldConc: number;
}

/**
 * Busca en el libro los bloques de resultados de detectores de radón.
 * Devuelve null si el archivo no tiene ese formato.
 */
export function extractRadonSamples(buffer: Buffer): RadonData | null {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });

  // La hoja "Resultados" tiene prioridad: su bloque "RESULTADOS PARA INFORME"
  // contiene los L.D. del informe (otras hojas tienen valores intermedios).
  const names = [...workbook.SheetNames].sort(
    (a, b) => sheetPriority(b) - sheetPriority(a),
  );

  for (const name of names) {
    const worksheet = workbook.Sheets[name];
    if (!worksheet || !worksheet["!ref"]) continue;

    // Dos pasadas alineadas fila a fila: texto formateado y valores crudos
    // (las fechas llegan como Date en la pasada cruda).
    const formatted = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: true,
    });
    const raw = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    });

    const samples: RadonSample[] = [];
    let columns: ColumnMap | null = null;
    let slide = "";

    for (let r = 0; r < formatted.length; r++) {
      const row = (formatted[r] ?? []).map((cell) => String(cell ?? "").trim());

      const headerColumns = findColumns(row);
      if (headerColumns) {
        columns = headerColumns;
        slide = row.find((cell) => /^SLIDE/i.test(cell)) ?? slide;
        continue;
      }
      if (!columns) continue;

      const id = row[columns.id] ?? "";
      if (!DETECTOR_ID.test(id)) continue;

      const rawRow = raw[r] ?? [];
      const text = (index: number) => (index >= 0 ? (row[index] ?? "") : "");
      const fecha = (index: number) =>
        index >= 0 ? formatFecha(rawRow[index], row[index] ?? "") : "";

      samples.push({
        expediente: normalizeOptionalText(text(columns.expediente)),
        id,
        procedencia: normalizeOptionalText(text(columns.procedencia)),
        slide,
        fechaColocacion: fecha(columns.diaInicial),
        fechaRetirada: fecha(columns.diaFinal),
        exposicion: text(columns.exposicion),
        incertidumbreExposicion: text(columns.uExp),
        ldExposicion: text(columns.ldExp),
        concentracion: text(columns.concentracion),
        incertidumbreConcentracion: text(columns.uConc),
        ldConcentracion: text(columns.ldConc),
      });
    }

    if (samples.length > 0) {
      return { sheetName: name, samples };
    }
  }

  return null;
}

function sheetPriority(name: string): number {
  return /resultado/i.test(name) ? 1 : 0;
}

/** Quita acentos, colapsa espacios y pasa a mayúsculas para comparar encabezados. */
function normalizeHeader(text: string): string {
  return text
    .normalize("NFD")
    .replace(new RegExp("[\\u0300-\\u036f]", "g"), "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

/**
 * Detecta una fila de encabezado del bloque de resultados. Las columnas se
 * buscan a partir de la columna ID para ignorar bloques auxiliares que
 * comparten fila (p. ej. "RESULTADOS HOJA DE CÁLCULO" a la izquierda).
 */
function findColumns(row: string[]): ColumnMap | null {
  const headers = row.map(normalizeHeader);
  const idIndex = headers.findIndex((header) => header === "ID");
  if (idIndex < 0) return null;

  const anywhere = (predicate: (header: string) => boolean) => headers.findIndex(predicate);
  const after = (predicate: (header: string) => boolean) => {
    for (let i = idIndex + 1; i < headers.length; i++) {
      if (predicate(headers[i])) return i;
    }
    return -1;
  };

  const map: ColumnMap = {
    // Algunos libros historicos tienen el encabezado escrito como "EXPEDENTE".
    expediente: anywhere((h) => h === "EXPEDIENTE" || h === "EXPEDENTE"),
    id: idIndex,
    procedencia: after((h) => h === "PROCEDENCIA"),
    diaInicial: after((h) => h.startsWith("DIA INICIAL")),
    diaFinal: after((h) => h.startsWith("DIA FINAL")),
    exposicion: after((h) => h.startsWith("EXPOSICION")),
    uExp: after((h) => h.startsWith("U(EXP)")),
    ldExp: after((h) => h.startsWith("LD EXPOSICION")),
    concentracion: after((h) => h.startsWith("CONCENTRACION")),
    uConc: after((h) => h.startsWith("U(CONC)")),
    ldConc: after((h) => h.startsWith("LD CONCENTRACION")),
  };

  // Imprescindibles para considerar la fila un encabezado de resultados.
  if (map.diaInicial < 0 || map.exposicion < 0 || map.concentracion < 0) {
    return null;
  }
  return map;
}

function normalizeOptionalText(value: string): string {
  const text = value.trim();
  return text === "0" ? "" : text;
}

/** Devuelve la fecha en formato dd/mm/aaaa a partir del valor crudo o del texto. */
function formatFecha(rawValue: unknown, formattedValue: string): string {
  if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
    const day = String(rawValue.getDate()).padStart(2, "0");
    const month = String(rawValue.getMonth() + 1).padStart(2, "0");
    return `${day}/${month}/${rawValue.getFullYear()}`;
  }
  const text = formattedValue.trim();
  const match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (match) {
    const [, day, month, year] = match;
    const fullYear = year.length === 2 ? `20${year}` : year;
    return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${fullYear}`;
  }
  return text;
}
