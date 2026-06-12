# UC · Fase 1 — Conversión de Excel a Word

Herramienta web para la Universidad de Cantabria que convierte archivos Excel en
informes Word. Sustituye el proceso manual de copiar y pegar datos entre
documentos: se sube el Excel, se extrae todo su contenido y se descarga un
documento Word con formato corporativo.

## Qué hace

1. El usuario sube un archivo **`.xlsx`, `.xls`, `.xlsm` o `.csv`** desde la web.
2. El servidor lee **todas las hojas** del libro y extrae los valores tal y como
   se muestran en Excel (respetando el formato visible de fechas e importes).
3. Genera un documento **Word (.docx)** con:
   - Logo de la UC en la cabecera de todas las páginas y pie con numeración.
   - Portada con metadatos: archivo de origen, fecha y hora de generación,
     hojas procesadas y filas totales.
   - Una sección por hoja con su tabla de datos completa (encabezado corporativo,
     filas alternas sombreadas, valores numéricos alineados a la derecha).
   - Orientación horizontal automática si alguna hoja tiene 7 columnas o más.
4. El nombre del archivo incluye fecha y hora de generación (hora peninsular):
   `Informe_<archivo>_AAAA-MM-DD_HH-MM-SS.docx`.

El archivo se procesa **en memoria** y no se almacena en el servidor.

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
  excel.ts              Extracción y normalización de los datos del Excel
  word.ts               Construcción del informe Word
  ucLogo.ts             Logo incrustado en base64 (generado, no editar a mano)
public/
  uc-logo.png           Logo mostrado en la web (provisional)
scripts/
  make-logo.py          Genera el logo provisional (Pillow)
  embed-logo.mjs        Incrusta public/uc-logo.png en lib/ucLogo.ts
  make-sample.mjs       Crea samples/ejemplo.xlsx para pruebas
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
```

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

## Límites actuales (Fase 1)

| Límite | Valor | Motivo |
|---|---|---|
| Tamaño del Excel | 4 MB | Vercel limita el cuerpo de la petición a 4,5 MB |
| Filas por hoja en el Word | 2 000 | Tiempo y memoria de la función serverless (se indica en el informe si se recorta) |
| Caracteres por celda | 500 | Legibilidad del documento |

Los límites son constantes configurables en `lib/excel.ts` y
`app/api/convert/route.ts`.

## Próximos pasos (Fase 2)

- Extracción de los datos específicos que se definan (métricas, análisis…)
  en lugar del volcado completo.
- Plantilla del informe final con las secciones definitivas.
