import { NextRequest, NextResponse } from "next/server";
import { parseWorkbook } from "@/lib/excel";
import { buildReport } from "@/lib/word";

export const runtime = "nodejs";
export const maxDuration = 60;

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"];

// Vercel limita el cuerpo de la petición a 4,5 MB en funciones serverless.
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

  const file = form.get("file");
  if (!(file instanceof File)) {
    return NextResponse.json(
      { error: 'No se ha recibido ningún archivo. Adjunta un Excel en el campo "file".' },
      { status: 400 },
    );
  }

  const extension = `.${(file.name.split(".").pop() ?? "").toLowerCase()}`;
  if (!ACCEPTED_EXTENSIONS.includes(extension)) {
    return NextResponse.json(
      { error: `Formato no admitido (${extension}). Formatos válidos: ${ACCEPTED_EXTENSIONS.join(", ")}.` },
      { status: 400 },
    );
  }

  if (file.size > MAX_SIZE_BYTES) {
    return NextResponse.json(
      { error: "El archivo supera el límite de 4 MB." },
      { status: 413 },
    );
  }

  let workbook;
  try {
    workbook = parseWorkbook(Buffer.from(await file.arrayBuffer()));
  } catch {
    return NextResponse.json(
      { error: "No se pudo leer el archivo. Comprueba que es un Excel válido y sin contraseña." },
      { status: 422 },
    );
  }

  if (workbook.sheets.length === 0) {
    return NextResponse.json(
      { error: "El archivo no contiene ninguna hoja." },
      { status: 422 },
    );
  }

  const generatedAt = new Date();
  const docxBuffer = await buildReport({
    data: workbook,
    sourceFileName: file.name,
    generatedAt,
  });
  const fileName = buildFileName(file.name, generatedAt);

  return new NextResponse(new Uint8Array(docxBuffer), {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${fileName}"`,
      "Cache-Control": "no-store",
    },
  });
}

/** Nombre del informe: Informe_<archivo>_<fecha>_<hora>.docx (hora peninsular). */
function buildFileName(sourceName: string, date: Date): string {
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
  return `Informe_${sanitizeBaseName(sourceName)}_${stamp}.docx`;
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
