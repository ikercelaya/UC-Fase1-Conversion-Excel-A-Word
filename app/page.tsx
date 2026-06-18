"use client";

import { useRef, useState } from "react";

const MAX_SIZE_MB = 4; // límite del cuerpo de la petición en Vercel (4,5 MB), aplicado al conjunto
const MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024;
const ACCEPTED = [".xlsx", ".xls", ".xlsm", ".csv"];
const CLIENT_OPTIONS = ["LaRUC", "Externo"];
const NORMATIVA_OPTIONS = ["ISO11665-4", "CTE", "IS-47"];

type Status = "idle" | "working" | "done" | "error";

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function extensionOf(name: string): string {
  return `.${(name.split(".").pop() ?? "").toLowerCase()}`;
}

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [client, setClient] = useState(CLIENT_OPTIONS[0]);
  const [normativa, setNormativa] = useState(NORMATIVA_OPTIONS[0]);
  const inputRef = useRef<HTMLInputElement>(null);

  function addFiles(incoming: FileList | File[] | null | undefined) {
    const candidates = Array.from(incoming ?? []);
    if (candidates.length === 0) return;

    const next = [...files];
    const errors: string[] = [];

    for (const candidate of candidates) {
      const extension = extensionOf(candidate.name);
      if (!ACCEPTED.includes(extension)) {
        errors.push(`${candidate.name}: formato no admitido (${extension}).`);
        continue;
      }
      // Evita duplicados (mismo nombre y tamaño).
      if (next.some((file) => file.name === candidate.name && file.size === candidate.size)) {
        continue;
      }
      const projectedSize = next.reduce((sum, file) => sum + file.size, 0) + candidate.size;
      if (projectedSize > MAX_SIZE_BYTES) {
        errors.push(`Se supera el límite de ${MAX_SIZE_MB} MB en total (${candidate.name} no añadido).`);
        continue;
      }
      next.push(candidate);
    }

    setFiles(next);
    if (errors.length > 0) {
      setStatus("error");
      setMessage(errors.join(" "));
    } else {
      setStatus("idle");
      setMessage("");
    }
  }

  function removeFile(index: number) {
    setFiles(files.filter((_, i) => i !== index));
    setStatus("idle");
    setMessage("");
  }

  function clearFiles() {
    setFiles([]);
    setStatus("idle");
    setMessage("");
  }

  async function generate() {
    if (files.length === 0 || status === "working") return;
    setStatus("working");
    setMessage("");
    try {
      const data = new FormData();
      data.append("normativa", normativa);
      for (const file of files) data.append("file", file);
      const res = await fetch("/api/convert", { method: "POST", body: data });

      if (!res.ok) {
        let detail = `Error ${res.status} al generar el informe.`;
        try {
          const body = await res.json();
          if (body?.error) detail = body.error;
        } catch {
          // respuesta sin cuerpo JSON
        }
        throw new Error(detail);
      }

      const blob = await res.blob();
      const disposition = res.headers.get("Content-Disposition") ?? "";
      const match = disposition.match(/filename="?([^";]+)"?/i);
      const name = match?.[1] ?? "Informe.docx";
      const mode = res.headers.get("X-Report-Mode");
      const sampleCount = res.headers.get("X-Sample-Count");
      const fileCount = res.headers.get("X-File-Count") ?? String(files.length);

      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = name;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);

      setStatus("done");
      const origen = `${fileCount} ${Number(fileCount) === 1 ? "archivo" : "archivos"}`;
      const detail =
        mode === "radon"
          ? ` — ${sampleCount} detectores extraídos de ${origen}`
          : mode === "mixto"
            ? ` — ${sampleCount} detectores y volcado de datos (${origen})`
            : mode === "volcado"
              ? ` — los archivos no tienen el formato del laboratorio; se ha volcado todo su contenido (${origen})`
              : "";
      setMessage(`Informe descargado: ${name}${detail}`);
    } catch (error) {
      setStatus("error");
      setMessage(
        error instanceof Error ? error.message : "Error inesperado al generar el informe.",
      );
    }
  }

  const totalSize = files.reduce((sum, file) => sum + file.size, 0);

  return (
    <>
      <header className="topbar">
        <img src="/uc-logo.png" alt="Universidad de Cantabria" className="topbar-logo" />
        <span className="topbar-caption">Conversión Excel → Word · Fase 1</span>
      </header>

      <main className="wrap">
        <h1>Genera un informe Word a partir de uno o varios Excel</h1>
        <p className="lead">
          Sube uno o varios Excel de medidas y se creará un único documento Word
          con la tabla de resultados de cada detector. Si algún archivo no tiene
          el formato del laboratorio, se volcará todo su contenido en tablas.
        </p>

        <div className="report-options">
          <label className="field">
            <span>Selecciona el cliente</span>
            <select value={client} onChange={(event) => setClient(event.target.value)}>
              {CLIENT_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </label>

          <label className="field">
            <span>Selecciona la normativa</span>
            <select value={normativa} onChange={(event) => setNormativa(event.target.value)}>
              {NORMATIVA_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </label>
        </div>

        <section
          className={`dropzone${dragOver ? " over" : ""}`}
          role="button"
          aria-label="Seleccionar archivos Excel"
          tabIndex={0}
          onClick={() => inputRef.current?.click()}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") inputRef.current?.click();
          }}
          onDragOver={(event) => {
            event.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(event) => {
            event.preventDefault();
            setDragOver(false);
            addFiles(event.dataTransfer.files);
          }}
        >
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPTED.join(",")}
            multiple
            hidden
            onChange={(event) => {
              addFiles(event.target.files);
              event.target.value = "";
            }}
          />
          <svg
            width="34"
            height="34"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
          >
            <path d="M12 3v12" />
            <path d="m7 8 5-5 5 5" />
            <path d="M5 15v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4" />
          </svg>
          <p>
            <strong>Arrastra tus archivos aquí</strong> o haz clic para seleccionarlos
          </p>
          <p className="hint">
            Formatos: {ACCEPTED.join(" · ")} — máx. {MAX_SIZE_MB} MB en total · puedes añadir varios
          </p>
        </section>

        {files.length > 0 && (
          <div className="filelist">
            <div className="filelist-head">
              <span>
                {files.length} {files.length === 1 ? "archivo" : "archivos"} · {formatSize(totalSize)}
              </span>
              <button type="button" className="clear" onClick={clearFiles}>
                Quitar todos
              </button>
            </div>
            {files.map((file, index) => (
              <div className="filecard" key={`${file.name}-${file.size}-${index}`}>
                <span className="filename">{file.name}</span>
                <span className="filesize">{formatSize(file.size)}</span>
                <button
                  type="button"
                  className="remove"
                  aria-label={`Quitar ${file.name}`}
                  onClick={() => removeFile(index)}
                >
                  ×
                </button>
              </div>
            ))}
          </div>
        )}

        <button
          type="button"
          className="primary"
          disabled={files.length === 0 || status === "working"}
          onClick={generate}
        >
          {status === "working"
            ? "Generando informe…"
            : files.length > 1
              ? `Generar informe Word (${files.length} archivos)`
              : "Generar informe Word"}
        </button>

        {message && <div className={`notice ${status}`}>{message}</div>}

        <p className="privacy">
          Los archivos se procesan en memoria y no se almacenan en el servidor.
        </p>
      </main>

      <footer className="foot">
        Universidad de Cantabria · Herramienta interna de generación de informes
      </footer>
    </>
  );
}
