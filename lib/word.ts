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
  WidthType,
} from "docx";
import { MAX_ROWS_PER_SHEET, SheetData, WorkbookData } from "./excel";
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

export interface ReportInput {
  data: WorkbookData;
  sourceFileName: string;
  generatedAt: Date;
}

/** Construye el informe Word y lo devuelve como buffer .docx. */
export async function buildReport({ data, sourceFileName, generatedAt }: ReportInput): Promise<Buffer> {
  const landscape = data.sheets.some((sheet) => sheet.columnCount >= LANDSCAPE_COLUMN_THRESHOLD);

  const document = new Document({
    creator: "Universidad de Cantabria",
    title: `Informe de datos — ${sourceFileName}`,
    description: "Informe generado automáticamente a partir de un archivo Excel",
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 21, color: INK }, // 10,5 pt
        },
        heading1: {
          run: { font: "Calibri", size: 26, bold: true, color: UC_TEAL },
          paragraph: { spacing: { before: 240, after: 120 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: landscape
              ? { width: A4_PORTRAIT.height, height: A4_PORTRAIT.width, orientation: PageOrientation.LANDSCAPE }
              : { width: A4_PORTRAIT.width, height: A4_PORTRAIT.height, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1700, bottom: 1100, left: 1000, right: 1000, header: 500, footer: 400 },
          },
        },
        headers: { default: buildHeader() },
        footers: { default: buildFooter() },
        children: [
          ...buildCover(sourceFileName, generatedAt, data),
          ...data.sheets.flatMap((sheet, index) => buildSheetSection(sheet, index)),
        ],
      },
    ],
  });

  return Packer.toBuffer(document);
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

/** Portada: título del informe y metadatos de la generación. */
function buildCover(sourceFileName: string, generatedAt: Date, data: WorkbookData): Paragraph[] {
  const formattedDate = new Intl.DateTimeFormat("es-ES", {
    dateStyle: "long",
    timeStyle: "medium",
    timeZone: "Europe/Madrid",
  }).format(generatedAt);

  const metadata: Array<[string, string]> = [
    ["Archivo de origen", sourceFileName],
    ["Fecha de generación", formattedDate],
    ["Hojas procesadas", String(data.sheets.length)],
    ["Filas de datos", data.totalRows.toLocaleString("es-ES")],
  ];

  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: "Informe de datos", size: 52, bold: true, color: INK })],
    }),
    new Paragraph({
      spacing: { after: 280 },
      children: [
        new TextRun({
          text: "Extracción completa del contenido del archivo Excel",
          size: 22,
          color: MUTED,
        }),
      ],
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
}

/** Sección de una hoja: título, resumen y tabla con todos los datos. */
function buildSheetSection(sheet: SheetData, index: number): (Paragraph | Table)[] {
  const blocks: (Paragraph | Table)[] = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
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
