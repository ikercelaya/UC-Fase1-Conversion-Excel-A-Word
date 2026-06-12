// Genera docs/Guia-de-uso.docx: guía de uso de la herramienta para el cliente.
// Uso: node scripts/make-guia.mjs
// Para actualizar la guía: edita este script y vuelve a ejecutarlo.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  PageNumber,
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

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = resolve(root, "docs");
const outPath = resolve(outDir, "Guia-de-uso.docx");

// --- identidad corporativa (la misma que la web y los informes) -------------
const TEAL = "0D9AA9";
const INK = "1D2B32";
const MUTED = "6B7A82";
const LINE = "D9E2E5";

const APP_URL = "https://uc-fase1-conversion-excel-a-word.vercel.app";
const VERSION = "1.0";
const FECHA = "junio de 2026";
const CONTACTO = "Iker Celaya — [correo de contacto]";

const logo = readFileSync(resolve(root, "public/uc-logo.png"));
const logoWidth = logo.readUInt32BE(16);
const logoHeight = logo.readUInt32BE(20);
const LOGO_H = 34;
const LOGO_W = Math.round((logoWidth / logoHeight) * LOGO_H);

// --- utilidades --------------------------------------------------------------
const p = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 120, ...(opts.spacing ?? {}) },
    children: [new TextRun({ text, ...opts.run })],
    ...opts.para,
  });

const h1 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });

const bullet = (children) =>
  new Paragraph({
    numbering: { reference: "vinetas", level: 0 },
    spacing: { after: 60 },
    children,
  });

const paso = (children) =>
  new Paragraph({
    numbering: { reference: "pasos", level: 0 },
    spacing: { after: 100 },
    children,
  });

const t = (text, run = {}) => new TextRun({ text, ...run });
const bold = (text) => t(text, { bold: true });

// --- tabla de mensajes -------------------------------------------------------
const COLS = [3400, 2800, 3300];
const TABLE_WIDTH = COLS.reduce((a, b) => a + b, 0);
const borders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
  left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
  right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
};

const headCell = (text, width) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: TEAL },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({ children: [t(text, { bold: true, color: "FFFFFF", size: 19 })] }),
    ],
  });

const cell = (text, width) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [t(text, { size: 19 })] })],
  });

const filaMensaje = ([mensaje, causa, solucion]) =>
  new TableRow({
    cantSplit: true,
    children: [cell(mensaje, COLS[0]), cell(causa, COLS[1]), cell(solucion, COLS[2])],
  });

const mensajes = [
  [
    "«Formato no admitido (…)»",
    "El archivo no es un Excel.",
    "Guardar el libro como .xlsx y volver a subirlo.",
  ],
  [
    "«El archivo supera el límite de 4 MB»",
    "Límite del servicio en línea.",
    "Eliminar hojas o datos innecesarios, o dividir el libro en varios archivos.",
  ],
  [
    "«No se pudo leer el archivo…»",
    "El libro está dañado o protegido con contraseña.",
    "Quitar la contraseña, guardar de nuevo y reintentar.",
  ],
  [
    "Aviso «se ha volcado todo su contenido»",
    "El Excel no tiene el bloque «RESULTADOS PARA INFORME» en la hoja Resultados.",
    "Comprobar que se ha subido el libro de medidas correcto del laboratorio.",
  ],
  [
    "La REFERENCIA UC sale vacía",
    "El nombre del archivo no empieza por el nº de informe.",
    "Renombrar el archivo, p. ej. «26024 (lote).xlsx», y generar de nuevo.",
  ],
];

// --- documento ---------------------------------------------------------------
const doc = new Document({
  creator: "Universidad de Cantabria",
  title: "Guía de uso — Generador de informes Word",
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 21, color: INK } },
      heading1: {
        run: { font: "Calibri", size: 26, bold: true, color: TEAL },
        paragraph: { spacing: { before: 280, after: 120 } },
      },
    },
  },
  numbering: {
    config: [
      {
        reference: "pasos",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
      {
        reference: "vinetas",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1700, bottom: 1100, left: 1000, right: 1000, header: 500, footer: 400 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 6 } },
              children: [
                new ImageRun({
                  type: "png",
                  data: logo,
                  transformation: { width: LOGO_W, height: LOGO_H },
                  altText: {
                    title: "Universidad de Cantabria",
                    description: "Logotipo de la Universidad de Cantabria",
                    name: "Logo UC",
                  },
                }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
              children: [
                t("Guía de uso · Universidad de Cantabria", { size: 16, color: MUTED }),
                new TextRun({
                  children: ["   ·   Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES],
                  size: 16,
                  color: MUTED,
                }),
              ],
            }),
          ],
        }),
      },
      children: [
        // portada compacta
        new Paragraph({
          spacing: { before: 120, after: 60 },
          children: [t("Guía de uso", { size: 48, bold: true, color: INK })],
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            t("Generador de informes Word a partir del Excel de medidas", {
              size: 24,
              color: MUTED,
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 320 },
          children: [t(`Versión ${VERSION} · ${FECHA}`, { size: 18, color: MUTED })],
        }),

        h1("1. Qué hace esta herramienta"),
        p(
          "Convierte el libro Excel de medidas del laboratorio en un informe Word con los resultados de cada detector, en segundos y sin copiar datos a mano. " +
            "La aplicación lee el bloque «RESULTADOS PARA INFORME» de la hoja Resultados y genera una tabla por detector con las fechas, la exposición, la concentración, las incertidumbres (k=2) y los límites de detección.",
        ),
        p(
          "El archivo se procesa en memoria y no se almacena en ningún servidor: al cerrar la página no queda ninguna copia del Excel ni del informe.",
        ),

        h1("2. Acceso"),
        p("La herramienta funciona desde el navegador (Chrome, Edge, Firefox o Safari), sin instalar nada y sin usuario ni contraseña:"),
        new Paragraph({
          spacing: { after: 120 },
          children: [t(APP_URL, { bold: true, color: TEAL })],
        }),
        p("Si la dirección cambia, el administrador comunicará la nueva.", {
          run: { size: 18, color: MUTED, italics: true },
        }),

        h1("3. Generar un informe, paso a paso"),
        paso([
          t("Abre la dirección anterior en el navegador."),
        ]),
        paso([
          t("Arrastra el archivo Excel a la zona punteada o haz clic en ella y selecciónalo. Formatos admitidos: "),
          bold(".xlsx, .xls, .xlsm y .csv"),
          t(", con un máximo de "),
          bold("4 MB"),
          t("."),
        ]),
        paso([
          t("Comprueba que el nombre y el tamaño que aparecen son los correctos y pulsa "),
          bold("«Generar informe Word»"),
          t(". Durante unos segundos el botón mostrará «Generando informe…»."),
        ]),
        paso([
          t("El informe se descarga automáticamente (normalmente en la carpeta Descargas) con el nombre "),
          bold("Informe_<archivo>_<fecha>_<hora>.docx"),
          t(". Un mensaje verde confirma la descarga e indica cuántos detectores se han extraído."),
        ]),
        p(
          "Para convertir otro archivo, pulsa la × de la tarjeta del archivo y repite los pasos. Cada informe lleva la fecha y la hora en el nombre, así que nunca se sobrescriben.",
          { spacing: { before: 60 } },
        ),

        h1("4. Qué contiene el informe"),
        bullet([t("Portada con el archivo de origen, el nº de informe, la fecha y hora de generación y el número de detectores procesados.")]),
        bullet([t("Las notas (1) y (2) del laboratorio sobre la información aportada por el cliente.")]),
        bullet([
          t("Una tabla por detector, con la misma estructura que el informe de ensayo: REFERENCIA, REFERENCIA UC, fechas de colocación y retirada, exposición, concentración, incertidumbres y límites de detección, con sus unidades (kBq·m⁻³·h, Bq·m⁻³)."),
        ]),
        bullet([
          t("La "),
          bold("REFERENCIA UC"),
          t(" se genera automáticamente con el patrón P-<nº de informe>-TRA-<n>, tomando el nº de informe del principio del nombre del archivo (p. ej. «26024 (HH8401_HN1150).xlsx» → P-26024-TRA-1, P-26024-TRA-2…)."),
        ]),
        bullet([
          t("El campo "),
          bold("PROCEDENCIA"),
          t(" queda en blanco porque no existe en el Excel de medidas; puede completarse a mano en Word si se necesita."),
        ]),
        p(
          "El documento es un Word normal: se puede revisar, completar o ajustar antes de enviarlo.",
          { spacing: { before: 60 } },
        ),

        h1("5. Requisitos del archivo Excel"),
        bullet([t("Debe ser el libro de medidas del laboratorio, con la hoja «Resultados» y sus bloques «RESULTADOS PARA INFORME» (uno por slide).")]),
        bullet([t("Las posiciones sin detector (ID «0») se ignoran automáticamente.")]),
        bullet([t("El nombre del archivo debe empezar por el nº de informe para que se genere la REFERENCIA UC.")]),
        bullet([t("Si el archivo no tiene este formato, la herramienta genera igualmente un Word con el volcado completo de todas las hojas y lo avisa en pantalla.")]),

        h1("6. Mensajes y solución de problemas"),
        new Table({
          width: { size: TABLE_WIDTH, type: WidthType.DXA },
          columnWidths: COLS,
          layout: TableLayoutType.FIXED,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          borders,
          rows: [
            new TableRow({
              tableHeader: true,
              cantSplit: true,
              children: [
                headCell("Mensaje", COLS[0]),
                headCell("Causa", COLS[1]),
                headCell("Qué hacer", COLS[2]),
              ],
            }),
            ...mensajes.map(filaMensaje),
          ],
        }),

        h1("7. Soporte"),
        p(
          "Para incidencias, dudas o mejoras (por ejemplo, incorporar la PROCEDENCIA o el texto introductorio del informe completo), contactar con:",
        ),
        p(CONTACTO, { run: { bold: true } }),
      ],
    },
  ],
});

mkdirSync(outDir, { recursive: true });
const buffer = await Packer.toBuffer(doc);
writeFileSync(outPath, buffer);
console.log(`Guía generada: ${outPath} (${(buffer.length / 1024).toFixed(1)} KB)`);
