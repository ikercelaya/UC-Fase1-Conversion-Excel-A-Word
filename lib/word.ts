import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  PageNumber,
  PageOrientation,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import { MAX_ROWS_PER_SHEET, RadonData, RadonSample, SheetData, WorkbookData } from "./excel";
import { UC_LOGO_BASE64, UC_LOGO_NATURAL } from "./ucLogo";

// Paleta corporativa (coherente con la interfaz web)
const UC_TEAL = "0D9AA9";
const INK = "1D2B32";
const MUTED = "6B7A82";
const LINE = "D9E2E5";
const ZEBRA = "F2F7F8";

// Tamaño del logo en la cabecera del documento (px)
const LOGO_HEIGHT = 34;
const LOGO_WIDTH = Math.round((UC_LOGO_NATURAL.width / UC_LOGO_NATURAL.height) * LOGO_HEIGHT);

// A4 en twips
const A4_PORTRAIT = { width: 11906, height: 16838 };

/** A partir de este número de columnas el documento se genera en horizontal. */
const LANDSCAPE_COLUMN_THRESHOLD = 7;

/** Celdas que parecen numéricas (importes, porcentajes…) se alinean a la derecha. */
const NUMERIC_PATTERN = /^-?(?:\d{1,3}(?:[.,\s]\d{3})*|\d+)(?:[.,]\d+)?\s*(?:%|€)?$/;

/**
 * Un archivo Excel ya procesado. O bien tiene el formato del laboratorio
 * (radón) o se vuelca su contenido completo. Cada archivo conserva su nombre
 * de origen y su nº de informe, que se usa para la REFERENCIA UC.
 */
export interface CombinedReportFile {
  sourceFileName: string;
  /** Nº de informe derivado del nombre del archivo ("" si no se pudo). */
  reportNumber: string;
  /** Datos de radón si el archivo tiene el formato del laboratorio. */
  radon: RadonData | null;
  /** Volcado completo si el archivo no tiene el formato del laboratorio. */
  workbook: WorkbookData | null;
}

export interface CombinedReportInput {
  files: CombinedReportFile[];
  generatedAt: Date;
}

/** Estilos comunes a todos los informes. */
function documentStyles() {
  return {
    default: {
      document: {
        run: { font: "Calibri", size: 21, color: INK }, // 10,5 pt
      },
      heading1: {
        run: { font: "Calibri", size: 26, bold: true, color: UC_TEAL },
        paragraph: { spacing: { before: 240, after: 120 } },
      },
    },
  };
}

/** Configuración de página A4 con la cabecera y márgenes corporativos. */
function pageProperties(landscape: boolean) {
  return {
    page: {
      size: landscape
        ? { width: A4_PORTRAIT.height, height: A4_PORTRAIT.width, orientation: PageOrientation.LANDSCAPE }
        : { width: A4_PORTRAIT.width, height: A4_PORTRAIT.height, orientation: PageOrientation.PORTRAIT },
      margin: { top: 1700, bottom: 1100, left: 1000, right: 1000, header: 500, footer: 400 },
    },
  };
}

/**
 * Construye un único informe Word a partir de uno o varios archivos Excel.
 * Cada archivo aporta sus detectores (formato radón LaRUC) o el volcado de sus
 * hojas, y todo se concatena en un mismo documento. El resultado se devuelve
 * como buffer .docx.
 */
export async function buildCombinedReport({
  files,
  generatedAt,
}: CombinedReportInput): Promise<Buffer> {
  const multiple = files.length > 1;
  const anyRadon = files.some((file) => file.radon);
  const anyVolcado = files.some((file) => !file.radon && file.workbook);
  const allRadon = anyRadon && !anyVolcado;
  const allVolcado = anyVolcado && !anyRadon;

  // El documento es horizontal si algún volcado tiene muchas columnas.
  const landscape = files.some((file) =>
    file.workbook?.sheets.some((sheet) => sheet.columnCount >= LANDSCAPE_COLUMN_THRESHOLD),
  );

  const children: (Paragraph | Table)[] = [
    ...buildCombinedCover(files, generatedAt, { allRadon, allVolcado }),
  ];

  // La leyenda (notas (1) y (2)) se muestra una sola vez, antes del primer
  // archivo de radón.
  let legendShown = false;
  files.forEach((file, index) => {
    const startNewPage = index > 0;
    if (file.radon) {
      children.push(buildRadonHeading(file, multiple, startNewPage));
      if (!legendShown) {
        children.push(...buildRadonLegend());
        legendShown = true;
      }
      children.push(...buildRadonTables(file));
    } else if (file.workbook) {
      children.push(...buildWorkbookBlocks(file, multiple, startNewPage));
    }
  });

  const document = new Document({
    creator: "Universidad de Cantabria",
    title: buildDocumentTitle(files),
    description: "Informe generado automáticamente a partir de archivos Excel",
    styles: documentStyles(),
    sections: [
      {
        properties: pageProperties(landscape),
        headers: { default: buildHeader() },
        footers: { default: buildFooter() },
        children,
      },
    ],
  });

  return Packer.toBuffer(document);
}

function buildDocumentTitle(files: CombinedReportFile[]): string {
  if (files.length === 1) return `Informe — ${files[0].sourceFileName}`;
  return `Informe — ${files.length} archivos`;
}

/** Cabecera de página: logo UC con una línea corporativa debajo. */
function buildHeader(): Header {
  return new Header({
    children: [
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: UC_TEAL, space: 6 } },
        spacing: { after: 0 },
        children: [
          new ImageRun({
            type: "png",
            data: Buffer.from(UC_LOGO_BASE64, "base64"),
            transformation: { width: LOGO_WIDTH, height: LOGO_HEIGHT },
            altText: {
              title: "Universidad de Cantabria",
              description: "Logotipo de la Universidad de Cantabria",
              name: "Logo UC",
            },
          }),
        ],
      }),
    ],
  });
}

/** Pie de página: nota institucional y numeración. */
function buildFooter(): Footer {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
        children: [
          new TextRun({
            text: "Universidad de Cantabria · Informe generado automáticamente",
            size: 16,
            color: MUTED,
          }),
          new TextRun({
            children: ["   ·   Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES],
            size: 16,
            color: MUTED,
          }),
        ],
      }),
    ],
  });
}

/** Portada: título del informe y metadatos de la generación (uno o varios archivos). */
function buildCombinedCover(
  files: CombinedReportFile[],
  generatedAt: Date,
  { allRadon, allVolcado }: { allRadon: boolean; allVolcado: boolean },
): Paragraph[] {
  const formattedDate = new Intl.DateTimeFormat("es-ES", {
    dateStyle: "long",
    timeStyle: "medium",
    timeZone: "Europe/Madrid",
  }).format(generatedAt);

  const title = allRadon ? "Informe de ensayo" : allVolcado ? "Informe de datos" : "Informe";
  const subtitle = allRadon
    ? "Determinación de la exposición a radón en aire"
    : allVolcado
      ? "Extracción completa del contenido de los archivos Excel"
      : "Resultados de radón y volcado de datos";

  const totalDetectores = files.reduce((sum, file) => sum + (file.radon?.samples.length ?? 0), 0);
  const totalHojas = files.reduce((sum, file) => sum + (file.workbook?.sheets.length ?? 0), 0);
  const totalFilas = files.reduce((sum, file) => sum + (file.workbook?.totalRows ?? 0), 0);

  const metadata: Array<[string, string]> = [];
  if (files.length === 1) {
    metadata.push(["Archivo de origen", files[0].sourceFileName]);
    if (files[0].reportNumber) metadata.push(["Nº de informe", files[0].reportNumber]);
  } else {
    metadata.push(["Archivos de origen", files.length.toLocaleString("es-ES")]);
  }
  metadata.push(["Fecha de generación", formattedDate]);
  if (totalDetectores > 0) {
    metadata.push(["Detectores procesados", totalDetectores.toLocaleString("es-ES")]);
  }
  if (totalHojas > 0) metadata.push(["Hojas procesadas", totalHojas.toLocaleString("es-ES")]);
  if (totalFilas > 0) metadata.push(["Filas de datos", totalFilas.toLocaleString("es-ES")]);

  const paragraphs: Paragraph[] = [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: title, size: 52, bold: true, color: INK })],
    }),
    new Paragraph({
      spacing: { after: 280 },
      children: [new TextRun({ text: subtitle, size: 22, color: MUTED })],
    }),
    ...metadata.map(
      ([label, value]) =>
        new Paragraph({
          spacing: { after: 40 },
          children: [
            new TextRun({ text: `${label}:  `, bold: true }),
            new TextRun({ text: value }),
          ],
        }),
    ),
  ];

  // Con varios archivos, se lista cada uno con su nº de detectores u hojas.
  if (files.length > 1) {
    paragraphs.push(
      new Paragraph({
        spacing: { before: 160, after: 40 },
        children: [new TextRun({ text: "Archivos incluidos", bold: true })],
      }),
      ...files.map((file, index) => {
        const detail = file.radon
          ? `${file.radon.samples.length} detectores${file.reportNumber ? ` · nº ${file.reportNumber}` : ""}`
          : `${file.workbook?.sheets.length ?? 0} hojas`;
        return new Paragraph({
          spacing: { after: 20 },
          children: [
            new TextRun({ text: `${index + 1}. ${file.sourceFileName}  ` }),
            new TextRun({ text: `(${detail})`, size: 18, color: MUTED }),
          ],
        });
      }),
    );
  }

  return paragraphs;
}

/** Sección de una hoja: título, resumen y tabla con todos los datos. */
function buildSheetSection(
  sheet: SheetData,
  index: number,
  pageBreakBefore = true,
): (Paragraph | Table)[] {
  const blocks: (Paragraph | Table)[] = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore,
      children: [new TextRun(`Hoja ${index + 1} · ${sheet.name}`)],
    }),
  ];

  if (sheet.rows.length === 0) {
    blocks.push(
      new Paragraph({
        children: [new TextRun({ text: "Esta hoja no contiene datos.", italics: true, color: MUTED })],
      }),
    );
    return blocks;
  }

  blocks.push(
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text:
            `${sheet.totalRows.toLocaleString("es-ES")} filas × ${sheet.columnCount} columnas` +
            (sheet.truncated
              ? ` — se muestran las primeras ${MAX_ROWS_PER_SHEET.toLocaleString("es-ES")}`
              : ""),
          size: 18,
          color: MUTED,
        }),
      ],
    }),
  );

  blocks.push(buildTable(sheet.rows));

  if (sheet.truncated) {
    blocks.push(
      new Paragraph({
        spacing: { before: 120 },
        children: [
          new TextRun({
            text: `Tabla recortada: la hoja original contiene ${sheet.totalRows.toLocaleString("es-ES")} filas.`,
            italics: true,
            size: 18,
            color: MUTED,
          }),
        ],
      }),
    );
  }

  return blocks;
}

/** Tabla de datos: primera fila como encabezado corporativo y filas alternas sombreadas. */
function buildTable(rows: string[][]): Table {
  const [headerRow, ...dataRows] = rows;

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.AUTOFIT,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headerRow.map(
          (text) =>
            new TableCell({
              shading: { type: ShadingType.CLEAR, fill: UC_TEAL },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: text || " ", bold: true, color: "FFFFFF", size: 18 })],
                }),
              ],
            }),
        ),
      }),
      ...dataRows.map(
        (row, rowIndex) =>
          new TableRow({
            children: row.map(
              (text) =>
                new TableCell({
                  shading:
                    rowIndex % 2 === 1 ? { type: ShadingType.CLEAR, fill: ZEBRA } : undefined,
                  children: buildCellParagraphs(text),
                }),
            ),
          }),
      ),
    ],
  });
}

/** Contenido de una celda: respeta saltos de línea y alinea a la derecha los valores numéricos. */
function buildCellParagraphs(text: string): Paragraph[] {
  const lines = text.split("\n").slice(0, 5);
  const isNumeric = NUMERIC_PATTERN.test(text.trim());

  return lines.map(
    (line) =>
      new Paragraph({
        alignment: isNumeric ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children: [new TextRun({ text: line, size: 18 })],
      }),
  );
}

// ---------------------------------------------------------------------------
// Informe de ensayo de radón (tabla por detector, formato del laboratorio)
// ---------------------------------------------------------------------------

// Anchos de columna del informe original del laboratorio (twips).
const RADON_COLUMNS = [2345, 1985, 2058, 3137];
const RADON_TABLE_WIDTH = RADON_COLUMNS.reduce((sum, width) => sum + width, 0);
const RADON_VALUE_SPAN_WIDTH = RADON_COLUMNS[1] + RADON_COLUMNS[2] + RADON_COLUMNS[3];

/** Encabezado de un archivo de radón: nombre del archivo, o título genérico si es uno solo. */
function buildRadonHeading(
  file: CombinedReportFile,
  multiple: boolean,
  startNewPage: boolean,
): Paragraph {
  const text = multiple ? file.sourceFileName : "Resultados por detector";
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: startNewPage,
    children: [new TextRun(text)],
  });
}

/** Tablas de los detectores de un archivo (REFERENCIA UC secuencial con su nº de informe). */
function buildRadonTables(file: CombinedReportFile): (Paragraph | Table)[] {
  const samples = file.radon?.samples ?? [];
  return samples.flatMap((sample, index) => [
    buildSampleTable(sample, referenciaUC(file.reportNumber, index + 1)),
    // separador: evita que Word funda tablas consecutivas
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [] }),
  ]);
}

/** Bloques del volcado completo de un archivo: título del archivo (si hay varios) y una sección por hoja. */
function buildWorkbookBlocks(
  file: CombinedReportFile,
  multiple: boolean,
  startNewPage: boolean,
): (Paragraph | Table)[] {
  const sheets = file.workbook?.sheets ?? [];
  if (!multiple) {
    return sheets.flatMap((sheet, index) => buildSheetSection(sheet, index));
  }
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: startNewPage,
      children: [new TextRun(file.sourceFileName)],
    }),
    ...sheets.flatMap((sheet, index) => buildSheetSection(sheet, index, index > 0)),
  ];
}

/** REFERENCIA UC secuencial (P-<informe>-TRA-<n>), como en los informes del laboratorio. */
function referenciaUC(reportNumber: string, position: number): string {
  return reportNumber ? `P-${reportNumber}-TRA-${position}` : "";
}

/** Notas (1) y (2) del informe original más el aviso de campos no disponibles. */
function buildRadonLegend(): Paragraph[] {
  const note = (marker: string, text: string) =>
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: marker, superScript: true, size: 18, color: MUTED }),
        new TextRun({ text: ` ${text}`, size: 18, color: MUTED }),
      ],
    });

  return [
    note("(1)", "La información ha sido proporcionada por el cliente."),
    note(
      "(2)",
      "El resultado de la concentración se ha calculado según las fechas de exposición facilitadas por el cliente.",
    ),
    new Paragraph({
      spacing: { after: 200 },
      children: [
        new TextRun({
          text: "El campo PROCEDENCIA no está disponible en el archivo Excel y se deja en blanco.",
          italics: true,
          size: 18,
          color: MUTED,
        }),
      ],
    }),
  ];
}

/** Tabla de un detector con la misma estructura que el informe del laboratorio. */
function buildSampleTable(sample: RadonSample, refUC: string): Table {
  return new Table({
    width: { size: RADON_TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: RADON_COLUMNS,
    layout: TableLayoutType.FIXED,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    },
    rows: [
      labelValueRow("PROCEDENCIA", [], true),
      labelValueRow("REFERENCIA", [new TextRun(sample.id)], true),
      labelValueRow("REFERENCIA UC", refUC ? [new TextRun(refUC)] : [], true),
      labelValueRow("FECHA COLOCACIÓN", clientDataRuns(sample.fechaColocacion), true),
      labelValueRow("FECHA RETIRADA", clientDataRuns(sample.fechaRetirada), true),
      pairRow(
        ["EXPOSICIÓN", expUnitRuns()],
        sample.exposicion,
        ["CONCENTRACIÓN", concUnitRuns()],
        sample.concentracion,
        true,
      ),
      pairRow(
        ["INCERTIDUMBRE", expUnitRuns()],
        sample.incertidumbreExposicion,
        ["INCERTIDUMBRE", concUnitRuns()],
        sample.incertidumbreConcentracion,
        true,
      ),
      pairRow(
        ["L.D.", expUnitRuns()],
        sample.ldExposicion,
        ["L.D.", concUnitRuns()],
        sample.ldConcentracion,
        false,
      ),
    ],
  });
}

function sup(text: string): TextRun {
  return new TextRun({ text, superScript: true });
}

/** Unidades de exposición: (kBq m⁻³ h). */
function expUnitRuns(): TextRun[] {
  return [new TextRun("(kBq m"), sup("-3"), new TextRun(" h)")];
}

/** Unidades de concentración: (Bq m⁻³) ⁽²⁾. */
function concUnitRuns(): TextRun[] {
  return [new TextRun("(Bq m"), sup("-3"), new TextRun(") "), sup("(2)")];
}

/** Valor aportado por el cliente: añade la nota superíndice (1) si hay valor. */
function clientDataRuns(value: string): TextRun[] {
  return value ? [new TextRun(`${value} `), sup("(1)")] : [];
}

function radonParagraph(
  runs: TextRun[],
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType],
  keepNext: boolean,
): Paragraph {
  return new Paragraph({ alignment, keepNext, children: runs });
}

function radonCell(paragraphs: Paragraph[], width: number, span?: number): TableCell {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: span,
    verticalAlign: VerticalAlign.CENTER,
    children: paragraphs,
  });
}

/** Fila "etiqueta | valor" (el valor ocupa las tres columnas restantes). */
function labelValueRow(label: string, valueRuns: TextRun[], keepNext: boolean): TableRow {
  return new TableRow({
    cantSplit: true,
    children: [
      radonCell([radonParagraph([new TextRun(label)], AlignmentType.LEFT, keepNext)], RADON_COLUMNS[0]),
      radonCell(
        [radonParagraph(valueRuns, AlignmentType.CENTER, keepNext)],
        RADON_VALUE_SPAN_WIDTH,
        3,
      ),
    ],
  });
}

/** Fila doble "magnitud exposición | valor | magnitud concentración | valor". */
function pairRow(
  [labelLeft, unitsLeft]: [string, TextRun[]],
  valueLeft: string,
  [labelRight, unitsRight]: [string, TextRun[]],
  valueRight: string,
  keepNext: boolean,
): TableRow {
  const labelCell = (label: string, units: TextRun[], width: number) =>
    radonCell(
      [
        radonParagraph([new TextRun(label)], AlignmentType.LEFT, true),
        radonParagraph(units, AlignmentType.LEFT, keepNext),
      ],
      width,
    );
  const valueCell = (value: string, width: number) =>
    radonCell([radonParagraph(value ? [new TextRun(value)] : [], AlignmentType.CENTER, keepNext)], width);

  return new TableRow({
    cantSplit: true,
    children: [
      labelCell(labelLeft, unitsLeft, RADON_COLUMNS[0]),
      valueCell(valueLeft, RADON_COLUMNS[1]),
      labelCell(labelRight, unitsRight, RADON_COLUMNS[2]),
      valueCell(valueRight, RADON_COLUMNS[3]),
    ],
  });
}
