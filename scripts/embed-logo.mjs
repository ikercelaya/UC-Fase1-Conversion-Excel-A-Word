// Incrusta public/uc-logo.png como base64 en lib/ucLogo.ts.
//
// El Word se genera en una función serverless de Vercel, donde no es fiable
// leer archivos de /public desde el sistema de ficheros; por eso el logo va
// incrustado en el código. Tras sustituir public/uc-logo.png por el logo
// definitivo, ejecuta: npm run embed-logo

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const pngPath = resolve(root, "public/uc-logo.png");
const outPath = resolve(root, "lib/ucLogo.ts");

const png = readFileSync(pngPath);
if (png.readUInt32BE(0) !== 0x89504e47) {
  console.error("public/uc-logo.png no parece un PNG válido.");
  process.exit(1);
}
const width = png.readUInt32BE(16);
const height = png.readUInt32BE(20);

const ts = `// Generado por scripts/embed-logo.mjs — no editar a mano.
// Para cambiar el logo: sustituye public/uc-logo.png y ejecuta \`npm run embed-logo\`.

/** Dimensiones reales del PNG (px). */
export const UC_LOGO_NATURAL = { width: ${width}, height: ${height} };

export const UC_LOGO_BASE64 =
  "${png.toString("base64")}";
`;

writeFileSync(outPath, ts);
console.log(
  `lib/ucLogo.ts actualizado (${width}x${height} px, ${(png.length / 1024).toFixed(1)} KB).`,
);
