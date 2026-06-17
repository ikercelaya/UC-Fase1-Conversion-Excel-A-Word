// Incrusta public/enac-logo.png como base64 en lib/enacLogo.ts.
//
// El logo de acreditación ENAC aparece en el pie del informe de ensayo. Igual
// que el resto de imágenes del Word, va incrustado en el código porque el
// documento se genera en una función serverless. Tras sustituir
// public/enac-logo.png, ejecuta: npm run embed-enac

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const pngPath = resolve(root, "public/enac-logo.png");
const outPath = resolve(root, "lib/enacLogo.ts");

const png = readFileSync(pngPath);
if (png.readUInt32BE(0) !== 0x89504e47) {
  console.error("public/enac-logo.png no parece un PNG válido.");
  process.exit(1);
}
const width = png.readUInt32BE(16);
const height = png.readUInt32BE(20);

const ts = `// Generado por scripts/embed-enac.mjs — no editar a mano.
// Para cambiar el logo: sustituye public/enac-logo.png y ejecuta \`npm run embed-enac\`.

/** Dimensiones reales del PNG (px). */
export const ENAC_LOGO_NATURAL = { width: ${width}, height: ${height} };

export const ENAC_LOGO_BASE64 =
  "${png.toString("base64")}";
`;

writeFileSync(outPath, ts);
console.log(
  `lib/enacLogo.ts actualizado (${width}x${height} px, ${(png.length / 1024).toFixed(1)} KB).`,
);
