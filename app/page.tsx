"use client";

import { useCallback, useRef, useState } from "react";

const MAX_SIZE_MB = 4; // límite de cuerpo de petición en Vercel (4,5 MB)
const ACCEPTED = [".xlsx", ".xls", ".xlsm", ".csv"];

type Status = "idle" | "working" | "done" | "error";

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const pick = useCallback((candidate: File | null | undefined) => {
    if (!candidate) return;
    const extension = `.${(candidate.name.split(".").pop() ?? "").toLowerCase()}`;
    if (!ACCEPTED.includes(extension)) {
      setFile(null);
      setStatus("error");
      setMessage(`Formato no admitido (${extension}). Usa ${ACCEPTED.join(", ")}.`);
      return;
    }
    if (candidate.size > MAX_SIZE_MB * 1024 * 1024) {
      setFile(null);
      setStatus("error");
      setMessage(`El archivo supera el límite de ${MAX_SIZE_MB} MB.`);
      return;
    }
    setFile(candidate);
    setStatus("idle");
    setMessage("");
  }, []);

  async function generate() {
    if (!file || status === "working") return;
    setStatus("working");
    setMessage("");
    try {
      const data = new FormData();
      data.append("file", file);
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

      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = name;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);

      setStatus("done");
      const detail =
        mode === "radon"
          ? ` — ${sampleCount} detectores extraídos`
          : mode === "volcado"
            ? " — el archivo no tiene el formato del laboratorio; se ha volcado todo su contenido"
            : "";
      setMessage(`Informe descargado: ${name}${detail}`);
    } catch (error) {
      setStatus("error");
      setMessage(
        error instanceof Error ? error.message : "Error inesperado al generar el informe.",
      );
    }
  }

  return (
    <>
      <header className="topbar">
        <img src="/uc-logo.png" alt="Universidad de Cantabria" className="topbar-logo" />
        <span className="topbar-caption">Conversión Excel → Word · Fase 1</span>
      </header>

      <main className="wrap">
        <h1>Genera un informe Word a partir de un Excel</h1>
        <p className="lead">
          Sube el Excel de medidas y se creará un documento Word con la tabla de
          resultados de cada detector. Si el archivo no tiene el formato del
          laboratorio, se volcará todo su contenido en tablas.
        </p>

        <section
          className={`dropzone${dragOver ? " over" : ""}`}
          role="button"
          aria-label="Seleccionar archivo Excel"
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
            pick(event.dataTransfer.files?.[0]);
          }}
        >
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPTED.join(",")}
            hidden
            onChange={(event) => {
              pick(event.target.files?.[0]);
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
            <strong>Arrastra tu archivo aquí</strong> o haz clic para seleccionarlo
          </p>
          <p className="hint">
            Formatos: {ACCEPTED.join(" · ")} — máx. {MAX_SIZE_MB} MB
          </p>
        </section>

        {file && (
          <div className="filecard">
            <span className="filename">{file.name}</span>
            <span className="filesize">{formatSize(file.size)}</span>
            <button
              type="button"
              className="remove"
              aria-label="Quitar archivo"
              onClick={() => {
                setFile(null);
                setStatus("idle");
                setMessage("");
              }}
            >
              ×
            </button>
          </div>
        )}

        <button
          type="button"
          className="primary"
          disabled={!file || status === "working"}
          onClick={generate}
        >
          {status === "working" ? "Generando informe…" : "Generar informe Word"}
        </button>

        {message && <div className={`notice ${status}`}>{message}</div>}

        <p className="privacy">
          El archivo se procesa en memoria y no se almacena en el servidor.
        </p>
      </main>

      <footer className="foot">
        Universidad de Cantabria · Herramienta interna de generación de informes
      </footer>
    </>
  );
}
