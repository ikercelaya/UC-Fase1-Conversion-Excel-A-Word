# UC — Conversión de Excel a Word (informes de radón LaRUC)

Herramienta web para la Universidad de Cantabria que convierte los Excel de
medidas del Laboratorio de Radiactividad Ambiental (LaRUC) en informes Word.
Sustituye el proceso manual de copiar y pegar datos entre documentos: se suben
uno o varios Excel, se extraen los resultados de cada detector y se descarga un
único documento Word con formato corporativo.

## Qué hace

1. El usuario sube **uno o varios** archivos **`.xlsx`, `.xls`, `.xlsm` o
   `.csv`** desde la web (límite de 4 MB para el conjunto).
2. El servidor procesa **cada archivo** y busca los bloques
   **"RESULTADOS PARA INFORME"** (hoja `Resultados`): por cada detector extrae
   ID, fechas de colocación y retirada, exposición, concentración,
   incertidumbres (k=2) y límites de detección. Los huecos sin detector
   (ID `0`) se descartan.
3. Genera un **único documento Word (.docx)** combinando todos los archivos con:
   - Logo de la UC en la cabecera de todas las páginas y pie con numeración.
   - Portada con metadatos: archivos de origen, fecha y hora de generación y
     nº total de detectores, más las notas (1) y (2) del laboratorio. Con
     varios archivos se incluye además la lista de archivos procesados.
   - **Una tabla por detector** con la misma estructura que el informe de
     ensayo del laboratorio: PROCEDENCIA, REFERENCIA, REFERENCIA UC, fechas,
     exposición/concentración con incertidumbre y L.D., con superíndices
     (kBq m⁻³ h, Bq m⁻³). Cada archivo conserva su propio nº de informe en la
     REFERENCIA UC.
4. El nombre del archivo incluye fecha y hora de generación (hora peninsular):
   `Informe_<archivo>_AAAA-MM-DD_HH-MM-SS.docx`.

Campos derivados o no disponibles:

- **REFERENCIA UC** se genera con el patrón del laboratorio
  `P-<nº informe>-TRA-<n>` (secuencial). El nº de informe se toma de los
  dígitos iniciales del nombre del archivo (`26024 (…).xlsx` → `26024`);
  si el nombre no empieza por dígitos, el campo queda en blanco.
- **PROCEDENCIA** no existe en el Excel de medidas y se deja en blanco.

Si el archivo no tiene el formato del laboratorio, la herramienta hace un
**volcado completo**: una sección por hoja con todos sus datos en tablas
(valores tal y como se muestran en Excel).

Los archivos se procesan **en memoria** y no se almacenan en el servidor.

## Stack

- [Next.js 15](https://nextjs.org/) (App Router) — desplegable en Vercel sin configuración.
- [SheetJS (xlsx)](https://sheetjs.com/) — lectura de los libros Excel.
- [docx](https://docx.js.org/) — generación del documento Word.

## Estructura

```
app/
  page.tsx              Interfaz de subida (arrastrar y soltar)
  layout.tsx            Metadatos y estructura HTML
  globals.css           Estilos (minimalista, color corporativo UC)
  api/convert/route.ts  Endpoint POST: Excel → Word
lib/
  excel.ts              Extracción: bloques de radón (LaRUC) y volcado completo
  word.ts               Construcción de los informes Word (radón y volcado)
  ucLogo.ts             Logo incrustado en base64 (generado, no editar a mano)
public/
  uc-logo.png           Logo mostrado en la web (provisional)
scripts/
  make-logo.py          Genera el logo provisional (Pillow)
  embed-logo.mjs        Incrusta public/uc-logo.png en lib/ucLogo.ts
  make-sample.mjs       Crea samples/ejemplo.xlsx para pruebas
  inspect-excel.mjs     Inspector: imprime hojas y contenido de un Excel
  make-guia.mjs         Genera docs/Guia-de-uso.docx (guía para el cliente)
docs/
  Guia-de-uso.docx      Guía de uso para entregar al cliente (generada)
samples/
  ejemplo.xlsx          Excel de ejemplo con varias hojas
```

## Desarrollo local

```bash
npm install
npm run dev        # http://localhost:3000
```

Para probar sin interfaz:

```bash
npm run sample     # crea samples/ejemplo.xlsx
curl -F "file=@samples/ejemplo.xlsx" -o informe.docx http://localhost:3000/api/convert

# varios archivos en una sola petición (campo "file" repetido):
curl -F "file=@archivo1.xlsx" -F "file=@archivo2.xlsx" \
  -o informe.docx http://localhost:3000/api/convert
```

## Guía de uso para el cliente

En `docs/Guia-de-uso.docx` hay una guía lista para entregar (acceso, pasos,
contenido del informe, requisitos del Excel y solución de problemas). Se genera
con `node scripts/make-guia.mjs`; antes de entregarla, revisa en ese script la
URL de la aplicación y el contacto de soporte, y vuelve a ejecutarlo.

## Despliegue en Vercel

1. Importa el repositorio en [vercel.com/new](https://vercel.com/new).
2. Vercel detecta Next.js automáticamente; no hace falta configuración adicional.
3. Cada push a `main` despliega una nueva versión.

También puede desplegarse desde terminal con `npx vercel`.

## Sustituir el logo provisional

El logo actual es un marcador de posición generado con `scripts/make-logo.py`.
Cuando tengas el logotipo oficial (PNG, idealmente con fondo transparente):

1. Sustituye `public/uc-logo.png` por el archivo oficial.
2. Ejecuta `npm run embed-logo` (regenera `lib/ucLogo.ts`, que es el logo que
   se incrusta en el Word).
3. Haz commit de ambos archivos.

## Límites actuales

| Límite | Valor | Motivo |
|---|---|---|
| Tamaño del conjunto de Excel | 4 MB | Vercel limita el cuerpo de la petición a 4,5 MB |
| Filas por hoja en el volcado | 2 000 | Tiempo y memoria de la función serverless (se indica en el informe si se recorta) |
| Caracteres por celda (volcado) | 500 | Legibilidad del documento |

Los límites son constantes configurables en `lib/excel.ts` y
`app/api/convert/route.ts`.

## Pendiente

- Campo PROCEDENCIA (no está en el Excel de medidas; habría que aportarlo de
  otra fuente o introducirlo en la web).
- Texto introductorio del informe de ensayo (datos del cliente, método,
  acreditación…), si se quiere replicar el informe completo del laboratorio.
- Logo oficial de la UC (ver «Sustituir el logo provisional»).
