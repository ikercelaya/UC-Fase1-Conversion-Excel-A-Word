import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  HeightRule,
  ImageRun,
  ISectionOptions,
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
  UnderlineType,
  VerticalAlign,
  WidthType,
} from "docx";
import { MAX_ROWS_PER_SHEET, RadonData, RadonSample, SheetData, WorkbookData } from "./excel";
import { ENAC_LOGO_BASE64, ENAC_LOGO_TYPE } from "./enacLogo";
import { LARUC_LOGO_BASE64, LARUC_LOGO_TYPE } from "./larucLogo";

// Paleta corporativa
const UC_TEAL = "0D9AA9";
const INK = "000000";
const MUTED = "6B7A82";
const LINE = "C7D2D7";
const ZEBRA = "F2F7F8";

// Logos del informe (px). Equivalen a los tamaños de la plantilla oficial.
const LARUC_LOGO_WIDTH = 206;
const LARUC_LOGO_HEIGHT = 84;
const ENAC_LOGO_SIZE = 103;

// A4 en twips
const A4 = { width: 11906, height: 16838 };
const MARGIN = { top: 1417, right: 1701, bottom: 1417, left: 1701, header: 708, footer: 708 };
const CONTENT_WIDTH = A4.width - MARGIN.left - MARGIN.right; // 8504

/** A partir de este número de columnas el volcado se genera en horizontal. */
const LANDSCAPE_COLUMN_THRESHOLD = 7;

/** Celdas que parecen numéricas (importes, porcentajes…) se alinean a la derecha. */
const NUMERIC_PATTERN = /^-?(?:\d{1,3}(?:[.,\s]\d{3})*|\d+)(?:[.,]\d+)?\s*(?:%|€)?$/;

// Datos fijos del laboratorio (cabecera y textos normalizados del informe).
const LAB_DEPT_LINES = [
  "Dpto. de Ciencias Médicas y Quirúrgicas",
  "Facultad de Medicina",
  "Avd. Cardenal Herrera Oria s/n C.P. 39011",
  "Santander (Tel: 942202207)",
];
const ENSAYO_OBJETO =
  "Exposición y concentración de gas radón en aire a través de los análisis llevados a cabo en el Laboratorio de Radiactividad de la Universidad de Cantabria.";
const ACREDITACION =
  "Laboratorio de ensayo acreditado por ENAC con acreditación Nº 1204/LE2219";
const FOOTER_LINE_1 =
  "Este informe no podrá reproducirse parcialmente sin la autorización escrita del Laboratorio de Radiactividad de la Universidad de Cantabria.";
const FOOTER_LINE_2 =
  "Este certificado se considera original si es un archivo digital y está firmado electrónicamente. En cualquier otro caso se considera una copia.";

/**
 * Un archivo Excel ya procesado. O bien tiene el formato del laboratorio
 * (radón) o se vuelca su contenido completo. Cada archivo conserva su nombre
 * de origen y su nº de informe, que se usa en la cabecera y la REFERENCIA UC.
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
  normativa: string;
}

/** Estilos comunes a todos los informes. */
function documentStyles() {
  return {
    default: {
      document: {
        run: { font: "Arial", size: 24, color: INK }, // 12 pt
      },
      heading1: {
        run: { font: "Arial", size: 26, bold: true, color: UC_TEAL },
        paragraph: { spacing: { before: 240, after: 120 } },
      },
    },
  };
}

/**
 * Construye un único informe Word a partir de uno o varios archivos Excel.
 * Cada archivo es un informe de ensayo completo (con su propia cabecera y nº
 * de informe), siguiendo la estructura oficial del laboratorio (LaRUC). El
 * resultado se devuelve como buffer .docx.
 */
export async function buildCombinedReport({
  files,
  generatedAt,
  normativa,
}: CombinedReportInput): Promise<Buffer> {
  void generatedAt; // la fecha de emisión la rellena el laboratorio a mano

  const sections: ISectionOptions[] = files.map((file) =>
    file.radon ? buildRadonSection(file, normativa) : buildVolcadoSection(file),
  );

  const document = new Document({
    creator: "Universidad de Cantabria · LaRUC",
    title:
      files.length === 1
        ? `Informe de ensayo — ${files[0].sourceFileName}`
        : `Informes de ensayo — ${files.length} archivos`,
    description: "Informe de ensayo generado automáticamente a partir del Excel de medidas",
    styles: documentStyles(),
    sections,
  });

  return Packer.toBuffer(document);
}

// ---------------------------------------------------------------------------
// Sección de un informe de ensayo de radón (estructura oficial del laboratorio)
// ---------------------------------------------------------------------------

function buildRadonSection(file: CombinedReportFile, normativa: string): ISectionOptions {
  const samples = file.radon?.samples ?? [];
  return {
    properties: {
      page: {
        size: { width: A4.width, height: A4.height, orientation: PageOrientation.PORTRAIT },
        margin: MARGIN,
        pageNumbers: { start: 1 }, // cada informe numera sus páginas desde 1
      },
    },
    headers: { default: buildHeader(file.reportNumber) },
    footers: { default: buildFooter() },
    children: [
      // Página 1: solo el título (portada). El resto va en la página 2.
      ...buildTitlePage(),
      // Página 2 en adelante: datos del informe, tablas y cierre con firma.
      ...buildDataPage(normativa),
      ...buildResultsIntro(),
      ...samples.flatMap((sample, index) => [
        buildSampleTable(sample, referenciaUC(file.reportNumber, index + 1)),
        new Paragraph({ spacing: { before: 60, after: 60 }, children: [] }),
      ]),
      ...buildClosing(),
    ],
  };
}

/** Cabecera de página: marca LaRUC (izq.) y datos del laboratorio + nº de informe (der.). */
function buildHeader(reportNumber: string): Header {
  const right = (runs: TextRun[]) =>
    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0, line: 240 }, children: runs });
  const headerRun = (text: string) =>
    new TextRun({ text, font: "Arial", bold: true, size: 16, color: INK });

  return new Header({
    children: [
      new Table({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        columnWidths: [3300, CONTENT_WIDTH - 3300],
        layout: TableLayoutType.FIXED,
        borders: noBorders(),
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders(),
                verticalAlign: VerticalAlign.TOP,
                children: [
                  new Paragraph({
                    spacing: { after: 0 },
                    children: [
                      new ImageRun({
                        type: LARUC_LOGO_TYPE,
                        data: Buffer.from(LARUC_LOGO_BASE64, "base64"),
                        transformation: { width: LARUC_LOGO_WIDTH, height: LARUC_LOGO_HEIGHT },
                        altText: {
                          title: "LaRUC",
                          description: "Laboratorio de Radiactividad Ambiental de la Universidad de Cantabria",
                          name: "Logo LaRUC",
                        },
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                borders: noBorders(),
                verticalAlign: VerticalAlign.TOP,
                children: [
                  right([
                    headerRun("Página "),
                    new TextRun({ children: [PageNumber.CURRENT], font: "Arial", bold: true, size: 16, color: INK }),
                    headerRun(" de "),
                    new TextRun({
                      children: [PageNumber.TOTAL_PAGES_IN_SECTION],
                      font: "Arial",
                      bold: true,
                      size: 16,
                      color: INK,
                    }),
                  ]),
                  ...LAB_DEPT_LINES.map((line) =>
                    right([headerRun(line)]),
                  ),
                  right([
                    new TextRun({
                      text: `Nº DE INFORME: ${reportNumber || ""}`,
                      font: "Arial",
                      bold: true,
                      size: 24,
                      color: INK,
                    }),
                  ]),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

/** Pie de página: logo ENAC (izq.) y nota legal (der.). */
function buildFooter(): Footer {
  const legal = (text: string) =>
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 0, line: 240 },
      children: [new TextRun({ text, font: "Arial", bold: true, size: 16, color: INK })],
    });

  return new Footer({
    children: [
      new Table({
        width: { size: CONTENT_WIDTH + MARGIN.left, type: WidthType.DXA },
        indent: { size: -MARGIN.left, type: WidthType.DXA },
        columnWidths: [1604, CONTENT_WIDTH + MARGIN.left - 1604],
        layout: TableLayoutType.FIXED,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        borders: noBorders(),
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders(),
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: { after: 0 },
                    children: [
                      new ImageRun({
                        type: ENAC_LOGO_TYPE,
                        data: Buffer.from(ENAC_LOGO_BASE64, "base64"),
                        transformation: { width: ENAC_LOGO_SIZE, height: ENAC_LOGO_SIZE },
                        altText: {
                          title: "ENAC",
                          description: "Acreditación ENAC Nº 1204/LE2219",
                          name: "Logo ENAC",
                        },
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                borders: noBorders(),
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 0, bottom: 0, left: 120, right: 0 },
                children: [legal(FOOTER_LINE_1), legal(FOOTER_LINE_2)],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

/** Página 1: portada con solo el título (centrado), debajo de la cabecera. */
function buildTitlePage(): Paragraph[] {
  return [
    new Paragraph({ children: [] }),
    new Paragraph({ children: [] }),
    new Paragraph({ children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: "INFORME DE ENSAYO", bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 0 },
      children: [],
    }),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 720 },
      children: [
        new TextRun({ text: "DETERMINACIÓN ", bold: true }),
        new TextRun({ text: "de la", bold: true, allCaps: true }),
        new TextRun({ text: " ", bold: true, allCaps: true }),
        new TextRun({ text: "CONCENTRACIÓN DE ", bold: true, allCaps: true }),
        new TextRun({ text: "radón en aire", bold: true, allCaps: true }),
      ],
    }),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { left: 360 }, children: [] }),
  ];
}

/**
 * Página 2: datos del informe siguiendo la plantilla del laboratorio.
 * Los datos que no llegan desde el Excel se dejan como marcadores "xx".
 */
function buildDataPage(normativa: string): Paragraph[] {
  return [
    new Paragraph({
      pageBreakBefore: true, // empieza en una página nueva (la 2)
      spacing: { before: 560, after: 220, line: 240 },
      children: [
        new TextRun({
          text: "DETERMINACIÓN DE LA CONCENTRACIÓN DE RADÓN EN AIRE.",
          bold: true,
          font: "Arial",
          size: 20,
          color: INK,
        }),
      ],
    }),

    dataSectionHeading("Datos del cliente"),
    dataItem("Entidad:", "Laboratorio de Radiactividad Ambiental de la Universidad de Cantabria"),
    dataItem("Dirección:", "c/ Cardenal Herrera Oria s/n"),
    dataItem("Persona de contacto:", "Carlos Sainz Fernández"),
    dataItem("Tel:", "(+34) 942 20 22 07"),
    dataItem("Email:", [
      new TextRun({
        text: "laruc@unican.es",
        font: "Arial",
        size: 20,
        color: "0563C1",
        underline: { type: UnderlineType.SINGLE },
      }),
    ]),

    dataSectionHeading("Objeto del informe"),
    dataItem("Ensayo a realizar:", ENSAYO_OBJETO),
    dataItem("Nº de detectores:", "xx"),
    dataItem("Nº de medidas realizadas:", "xx"),

    dataSectionHeading("Datos de las muestras objeto del ensayo"),
    dataItem("Los detectores han sido colocados por", "LaRUC"),
    dataItem("Los detectores han sido recogidos por", "LaRUC"),
    dataItem("Los detectores han sido aptos para su ensayo", "Sí"),
    dataItem("Lugar de colocación del detector/es:", "xx"),
    dataItem("Fecha de colocación del detector/es:", "xx/xx/2025"),
    dataItem("Fecha de retirada del detector/es:", "xx/xx/2025"),
    dataItem("Fecha de recepción en el laboratorio:", "xx/xx/2025"),
    dataItem("Fecha inicio ensayo:", "xx/xx/2025"),
    dataItem("Fecha final ensayo:", "xx/xx/2025"),

    dataSectionHeading("Método de ensayo"),
    dataItem("Lugar de realización del ensayo:", "Instalaciones de LaRUC"),
    dataItem(
      "Método de ensayo empleado:",
      "El método empleado ha sido el que se recoge en la documentación de calidad del laboratorio referencia I-Ens01_10.",
    ),

    dataSectionHeading("Normativa que afecta a este ensayo"),
    highlightedDataItem(normativa),

    dataSectionHeading("Incidencias durante la captación, retirada, transporte y/o ensayo"),
    plainItem("No aplica."),

    dataSectionHeading(ACREDITACION),
  ];
}

function dataTextRun(text: string, bold = false): TextRun {
  return new TextRun({ text, font: "Arial", bold, size: 20, color: INK });
}

function dataSectionHeading(text: string): Paragraph {
  return new Paragraph({
    indent: { left: 360, hanging: 240 },
    spacing: { before: 160, after: 80, line: 240 },
    keepNext: true,
    children: [dataTextRun("▪  ", true), dataTextRun(text, true)],
  });
}

function dataItem(label: string, value: string | TextRun[] = ""): Paragraph {
  const runs = [dataTextRun("–  "), dataTextRun(label, true)];
  if (Array.isArray(value)) {
    if (value.length > 0) runs.push(dataTextRun(" "));
    runs.push(...value);
  } else if (value) {
    runs.push(dataTextRun(` ${value}`));
  }

  return new Paragraph({
    indent: { left: 920, hanging: 240 },
    spacing: { after: 30, line: 240 },
    children: runs,
  });
}

function highlightedDataItem(text: string): Paragraph {
  return new Paragraph({
    indent: { left: 680 },
    spacing: { after: 30, line: 240 },
    children: [
      new TextRun({
        text,
        font: "Arial",
        size: 20,
        color: INK,
        highlight: "yellow",
      }),
    ],
  });
}

/** Encabezado "Resultados obtenidos" y párrafo introductorio normalizado. */
function buildResultsIntro(): Paragraph[] {
  return [
    sectionHeading("Resultados obtenidos", true),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 160, line: 240 },
      children: [
        new TextRun({
          text:
            "Los resultados que contiene este informe solo afectan a los detectores sometidos a ensayo. " +
            "Las tablas siguientes contienen los resultados de la medida expresando la exposición en unidades kBq m",
          font: "Arial",
          size: 20,
        }),
        sup("-3"),
        new TextRun({ text: " h y la concentración en unidades Bq m", font: "Arial", size: 20 }),
        sup("-3"),
        new TextRun({
          text:
            ". Los resultados de incertidumbre de este informe de ensayo se corresponden con un factor de " +
            "cobertura k = 2. Los valores de la incertidumbre aparecen expresados con dos cifras significativas " +
            "y el resto de valores del apartado de resultados se expresan en coherencia con la incertidumbre. " +
            "Se sigue lo indicado en el documento 'Evaluation of measurement data — Guide to the expression of " +
            "uncertainty in measurement' (JCGM 100:2008 GUM 1995 with minor corrections).",
          font: "Arial",
          size: 20,
        }),
      ],
    }),
  ];
}

/** Cierre del informe: "Fin del informe", la línea de firma y el recuadro vacío. */
function buildClosing(): (Paragraph | Table)[] {
  return [
    sectionHeading("Fin del informe"),
    bulletItem("Fecha de emisión y firma (Dirección Técnica):"),
    buildSignatureBox(),
  ];
}

/** Recuadro vacío para la fecha de emisión y la firma de Dirección Técnica. */
function buildSignatureBox(): Table {
  const edge = { style: BorderStyle.SINGLE, size: 6, color: "555555" };
  return new Table({
    width: { size: 6500, type: WidthType.DXA },
    indent: { size: 560, type: WidthType.DXA },
    columnWidths: [6500],
    layout: TableLayoutType.FIXED,
    borders: {
      top: edge,
      bottom: edge,
      left: edge,
      right: edge,
      insideHorizontal: edge,
      insideVertical: edge,
    },
    rows: [
      new TableRow({
        cantSplit: true,
        height: { value: 1700, rule: HeightRule.ATLEAST },
        children: [new TableCell({ children: [new Paragraph({ children: [] })] })],
      }),
    ],
  });
}

/** Encabezado de sección con viñeta cuadrada (▪). */
function sectionHeading(text: string, pageBreakBefore = false): Paragraph {
  return new Paragraph({
    pageBreakBefore,
    indent: { left: 360, hanging: 240 },
    spacing: { before: pageBreakBefore ? 560 : 200, after: 80, line: 240 },
    keepNext: true,
    children: [
      new TextRun({ text: "▪  ", font: "Arial", bold: true, size: 20, color: INK }),
      new TextRun({ text, font: "Arial", bold: true, size: 20, color: INK }),
    ],
  });
}

/** Punto "– etiqueta valor" con sangría (valor vacío = a rellenar a mano). */
function bulletItem(label: string, value = ""): Paragraph {
  const runs = [new TextRun({ text: "–  ", font: "Arial", size: 20, color: INK })];
  if (label) runs.push(new TextRun({ text: label, font: "Arial", bold: true, size: 20, color: INK }));
  if (value) runs.push(new TextRun({ text: `${label ? " " : ""}${value}`, font: "Arial", size: 20, color: INK }));
  return new Paragraph({ indent: { left: 920, hanging: 240 }, spacing: { after: 30, line: 240 }, children: runs });
}

/** Texto con sangría sin viñeta (p. ej. "ISO 11665-4", "No aplica."). */
function plainItem(text: string): Paragraph {
  return new Paragraph({
    indent: { left: 680 },
    spacing: { after: 30, line: 240 },
    children: [new TextRun({ text, font: "Arial", size: 20, color: INK })],
  });
}

// ---------------------------------------------------------------------------
// Tabla de un detector (misma estructura que el informe del laboratorio)
// ---------------------------------------------------------------------------

// Anchos de columna del informe original del laboratorio (twips).
const RADON_COLUMNS = [2600, 1700, 2200, CONTENT_WIDTH - 6500];
const RADON_TABLE_WIDTH = RADON_COLUMNS.reduce((sum, width) => sum + width, 0);
const RADON_VALUE_SPAN_WIDTH = RADON_COLUMNS[1] + RADON_COLUMNS[2] + RADON_COLUMNS[3];

/** REFERENCIA UC secuencial (P-<informe>-TRA-<n>), como en los informes del laboratorio. */
function referenciaUC(reportNumber: string, position: number): string {
  return reportNumber ? `P-${reportNumber}-TRA-${position}` : "";
}

/** Tabla de un detector con la misma estructura que el informe del laboratorio. */
function buildSampleTable(sample: RadonSample, refUC: string): Table {
  return new Table({
    width: { size: RADON_TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: RADON_COLUMNS,
    layout: TableLayoutType.FIXED,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: INK },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: INK },
      left: { style: BorderStyle.SINGLE, size: 4, color: INK },
      right: { style: BorderStyle.SINGLE, size: 4, color: INK },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: INK },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: INK },
    },
    rows: [
      labelValueRow("PROCEDENCIA", [], true),
      labelValueRow("REFERENCIA", valueRuns(sample.id, true), true),
      labelValueRow("REFERENCIA UC", valueRuns(refUC), true),
      labelValueRow("FECHA COLOCACIÓN", valueRuns(sample.fechaColocacion), true),
      labelValueRow("FECHA RETIRADA", valueRuns(sample.fechaRetirada), true),
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
  return new TextRun({ text, font: "Arial", size: 20, superScript: true });
}

/** Unidades de exposición: (kBq m⁻³ h). */
function expUnitRuns(): TextRun[] {
  return [new TextRun({ text: "(kBq m", font: "Arial", size: 20 }), sup("-3"), new TextRun({ text: " h)", font: "Arial", size: 20 })];
}

/** Unidades de concentración: (Bq m⁻³). */
function concUnitRuns(): TextRun[] {
  return [new TextRun({ text: "(Bq m", font: "Arial", size: 20 }), sup("-3"), new TextRun({ text: ")", font: "Arial", size: 20 })];
}

function valueRuns(value: string, bold = false): TextRun[] {
  return value ? [new TextRun({ text: value, font: "Arial", bold, size: 20, color: INK })] : [];
}

function radonParagraph(
  runs: TextRun[],
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType],
  keepNext: boolean,
): Paragraph {
  return new Paragraph({ alignment, keepNext, spacing: { after: 0, line: 240 }, children: runs });
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
      radonCell([radonParagraph([new TextRun({ text: label, font: "Arial", size: 20, color: INK })], AlignmentType.LEFT, keepNext)], RADON_COLUMNS[0]),
      radonCell(
        [radonParagraph(valueRuns, AlignmentType.LEFT, keepNext)],
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
        radonParagraph([new TextRun({ text: label, font: "Arial", size: 20, color: INK })], AlignmentType.CENTER, true),
        radonParagraph(units, AlignmentType.CENTER, keepNext),
      ],
      width,
    );
  const valueCell = (value: string, width: number) =>
    radonCell([radonParagraph(valueRuns(value), AlignmentType.CENTER, keepNext)], width);

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

// ---------------------------------------------------------------------------
// Sección de volcado (archivos que no tienen el formato del laboratorio)
// ---------------------------------------------------------------------------

function buildVolcadoSection(file: CombinedReportFile): ISectionOptions {
  const sheets = file.workbook?.sheets ?? [];
  const landscape = sheets.some((sheet) => sheet.columnCount >= LANDSCAPE_COLUMN_THRESHOLD);

  return {
    properties: {
      page: {
        size: landscape
          ? { width: A4.height, height: A4.width, orientation: PageOrientation.LANDSCAPE }
          : { width: A4.width, height: A4.height, orientation: PageOrientation.PORTRAIT },
        margin: MARGIN,
        pageNumbers: { start: 1 },
      },
    },
    headers: { default: buildHeader(file.reportNumber) },
    footers: { default: buildFooter() },
    children: [
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: `Volcado de datos — ${file.sourceFileName}`, bold: true, size: 28, color: INK })],
      }),
      new Paragraph({
        spacing: { after: 160 },
        children: [
          new TextRun({
            text: "El archivo no tiene el formato del laboratorio; se vuelca todo su contenido.",
            italics: true,
            size: 18,
            color: MUTED,
          }),
        ],
      }),
      ...sheets.flatMap((sheet, index) => buildSheetSection(sheet, index, index > 0)),
    ],
  };
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

/** Bordes invisibles para tablas de maquetación (cabecera/pie). */
function noBorders() {
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return {
    top: none,
    bottom: none,
    left: none,
    right: none,
    insideHorizontal: none,
    insideVertical: none,
  };
}
