import { NextRequest, NextResponse } from "next/server";
import { extractRadonSamples, parseWorkbook } from "@/lib/excel";
import { buildCombinedReport, CombinedReportFile } from "@/lib/word";

export const runtime = "nodejs";
export const maxDuration = 60;

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"];
const CLIENTES = ["LaRUC", "Externo"] as const;
const DEFAULT_CLIENTE = CLIENTES[0];
const NORMATIVAS = ["ISO11665-4", "CTE", "IS-47"] as const;
const DEFAULT_NORMATIVA = NORMATIVAS[0];

// Vercel limita el cuerpo de la petición a 4,5 MB en funciones serverless.
// El límite se aplica al conjunto de archivos (todo va en la misma petición).
const MAX_SIZE_BYTES = 4 * 1024 * 1024;

export async function POST(request: NextRequest) {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return NextResponse.json(
      { error: "La petición no contiene un formulario válido." },
      { status: 400 },
    );
  }

  // Se aceptan varios archivos en el campo "file"; se combinan en un único Word.
  const uploads = form.getAll("file").filter((entry): entry is File => entry instanceof File);
  const clienteValue = form.get("cliente");
  const cliente =
    typeof clienteValue === "string" && CLIENTES.includes(clienteValue as (typeof CLIENTES)[number])
      ? (clienteValue as (typeof CLIENTES)[number])
      : DEFAULT_CLIENTE;
  const normativaValue = form.get("normativa");
  const normativa =
    typeof normativaValue === "string" && NORMATIVAS.includes(normativaValue as (typeof NORMATIVAS)[number])
      ? normativaValue
      : DEFAULT_NORMATIVA;

  if (uploads.length === 0) {
    return NextResponse.json(
      { error: 'No se ha recibido ningún archivo. Adjunta uno o varios Excel en el campo "file".' },
      { status: 400 },
    );
  }

  for (const upload of uploads) {
    const extension = `.${(upload.name.split(".").pop() ?? "").toLowerCase()}`;
    if (!ACCEPTED_EXTENSIONS.includes(extension)) {
      return NextResponse.json(
        {
          error: `Formato no admitido en "${upload.name}" (${extension}). Formatos válidos: ${ACCEPTED_EXTENSIONS.join(", ")}.`,
        },
        { status: 400 },
      );
    }
  }

  const totalSize = uploads.reduce((sum, upload) => sum + upload.size, 0);
  if (totalSize > MAX_SIZE_BYTES) {
    return NextResponse.json(
      { error: "El conjunto de archivos supera el límite de 4 MB." },
      { status: 413 },
    );
  }

  // Extracción específica del formato del laboratorio (radón) por archivo; si un
  // archivo no tiene ese formato, se vuelca su contenido completo.
  const files: CombinedReportFile[] = [];
  try {
    for (const upload of uploads) {
      const buffer = Buffer.from(await upload.arrayBuffer());
      const radon = extractRadonSamples(buffer);
      const workbook = radon ? null : parseWorkbook(buffer);

      if (!radon && (!workbook || workbook.sheets.length === 0)) {
        return NextResponse.json(
          { error: `El archivo "${upload.name}" no contiene ninguna hoja.` },
          { status: 422 },
        );
      }

      files.push({
        sourceFileName: upload.name,
        reportNumber: extractReportNumber(upload.name),
        radon,
        workbook,
      });
    }
  } catch {
    return NextResponse.json(
      {
        error:
          "No se pudo leer alguno de los archivos. Comprueba que son Excel válidos y sin contraseña.",
      },
      { status: 422 },
    );
  }

  const generatedAt = new Date();
  const docxBuffer = await buildCombinedReport({ files, generatedAt, cliente, normativa });
  const fileName = buildFileName(
    uploads.map((upload) => upload.name),
    generatedAt,
  );

  const sampleCount = files.reduce((sum, file) => sum + (file.radon?.samples.length ?? 0), 0);
  const anyRadon = files.some((file) => file.radon);
  const anyVolcado = files.some((file) => !file.radon);
  const mode = anyRadon && anyVolcado ? "mixto" : anyRadon ? "radon" : "volcado";

  return new NextResponse(new Uint8Array(docxBuffer), {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${fileName}"`,
      "Cache-Control": "no-store",
      "X-Report-Mode": mode,
      "X-Sample-Count": String(sampleCount),
      "X-File-Count": String(files.length),
    },
  });
}

/** Nº de informe del laboratorio: dígitos iniciales del nombre del archivo (p. ej. "26024 (...)" → 26024). */
function extractReportNumber(name: string): string {
  const match = name.trim().match(/^(\d{3,6})/);
  return match ? match[1] : "";
}

/**
 * Nombre del informe: Informe_<archivo|N_archivos>_<fecha>_<hora>.docx
 * (hora peninsular). Con un solo archivo usa su nombre; con varios, el número
 * de archivos.
 */
function buildFileName(sourceNames: string[], date: Date): string {
  const parts = new Intl.DateTimeFormat("es-ES", {
    timeZone: "Europe/Madrid",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  })
    .formatToParts(date)
    .reduce<Record<string, string>>((acc, part) => {
      acc[part.type] = part.value;
      return acc;
    }, {});

  const stamp = `${parts.year}-${parts.month}-${parts.day}_${parts.hour}-${parts.minute}-${parts.second}`;
  const base =
    sourceNames.length === 1 ? sanitizeBaseName(sourceNames[0]) : `${sourceNames.length}_archivos`;
  return `Informe_${base}_${stamp}.docx`;
}

/** Limpia el nombre del archivo origen para usarlo en el nombre del informe. */
function sanitizeBaseName(name: string): string {
  const base = name.replace(/\.[^.]+$/, "");
  const clean = base
    .normalize("NFD")
    .replace(new RegExp("[\\u0300-\\u036f]", "g"), "")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/_{2,}/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 40);
  return clean || "Datos";
}
