// Incrusta public/laruc-logo.jpg como base64 en lib/larucLogo.ts.
//
// El logotipo LaRUC aparece en la cabecera de todos los informes. Va incrustado
// en codigo porque el documento se genera en una funcion serverless. Tras
// sustituir public/laruc-logo.jpg, ejecuta: npm run embed-laruc

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const jpgPath = resolve(root, "public/laruc-logo.jpg");
const outPath = resolve(root, "lib/larucLogo.ts");

const jpg = readFileSync(jpgPath);
if (jpg[0] !== 0xff || jpg[1] !== 0xd8) {
  console.error("public/laruc-logo.jpg no parece un JPG valido.");
  process.exit(1);
}

const { width, height } = readJpegSize(jpg);

const ts = `// Generado por scripts/embed-laruc.mjs - no editar a mano.
// Para cambiar el logo: sustituye public/laruc-logo.jpg y ejecuta \`npm run embed-laruc\`.

/** Dimensiones reales del JPG (px). */
export const LARUC_LOGO_NATURAL = { width: ${width}, height: ${height} };

export const LARUC_LOGO_TYPE = "jpg" as const;

export const LARUC_LOGO_BASE64 =
  "${jpg.toString("base64")}";
`;

writeFileSync(outPath, ts);
console.log(
  `lib/larucLogo.ts actualizado (${width}x${height} px, ${(jpg.length / 1024).toFixed(1)} KB).`,
);

function readJpegSize(buffer) {
  let offset = 2;
  while (offset < buffer.length) {
    if (buffer[offset] !== 0xff) break;
    const marker = buffer[offset + 1];
    const length = buffer.readUInt16BE(offset + 2);
    if (
      marker === 0xc0 ||
      marker === 0xc1 ||
      marker === 0xc2 ||
      marker === 0xc3 ||
      marker === 0xc5 ||
      marker === 0xc6 ||
      marker === 0xc7 ||
      marker === 0xc9 ||
      marker === 0xca ||
      marker === 0xcb ||
      marker === 0xcd ||
      marker === 0xce ||
      marker === 0xcf
    ) {
      return {
        height: buffer.readUInt16BE(offset + 5),
        width: buffer.readUInt16BE(offset + 7),
      };
    }
    offset += 2 + length;
  }
  throw new Error("No se pudieron leer las dimensiones del JPG.");
}
