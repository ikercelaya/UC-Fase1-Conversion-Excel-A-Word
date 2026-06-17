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
3. Genera un **documento Word (.docx)** que sigue la **estructura oficial del
   informe de ensayo del laboratorio (LaRUC)**. Cada Excel produce un informe
   completo (si se suben varios, se concatenan, cada uno con su propia cabecera
   y numeración de páginas):
   - **Cabecera** en todas las páginas: marca LaRUC, datos del departamento,
     numeración (`Página X de Y`) y `Nº DE INFORME` (tomado del nombre del
     archivo). **Pie** con la nota legal y el logo de acreditación **ENAC**.
   - **Página 1**: portada, solo con el título (*INFORME DE ENSAYO* /
     *DETERMINACIÓN DE LA CONCENTRACIÓN DE RADÓN EN AIRE*), cabecera y pie.
   - **Página 2**: la estructura del informe (*Datos del cliente*, *Objeto del
     informe*, *Datos de las muestras*, *Método de ensayo*, *Normativa*,
     *Incidencias* y la línea de acreditación ENAC). Siguiendo la plantilla del
     laboratorio, **los datos del cliente, de las muestras y del método se dejan
     en blanco** para rellenarlos a mano; solo se incluyen los textos fijos.
   - Apartado *Resultados obtenidos* con el párrafo normativo y **una tabla por
     detector** (PROCEDENCIA, REFERENCIA, REFERENCIA UC, fechas,
     exposición/concentración con incertidumbre y L.D.), con las unidades en
     superíndice (kBq m⁻³ h, Bq m⁻³).
   - Cierre con *Fin del informe*, la línea *Fecha de emisión y firma (Dirección
     Técnica)* y un **recuadro vacío** para la firma.
4. El nombre del archivo incluye fecha y hora de generación (hora peninsular):
   `Informe_<archivo>_AAAA-MM-DD_HH-MM-SS.docx`.

Campos derivados o no disponibles:

- **REFERENCIA UC** se genera con el patrón del laboratorio
  `P-<nº informe>-TRA-<n>` (secuencial). El nº de informe se toma de los
  dígitos iniciales del nombre del archivo (`26024 (…).xlsx` → `26024`);
  si el nombre no empieza por dígitos, el campo queda en blanco.
- **PROCEDENCIA** no existe en el Excel de medidas y se deja en blanco.
- El **logotipo LaRUC** de la cabecera se compone como texto (el logo oficial
  es un EMF que no se puede incrustar directamente); puede sustituirse más
  adelante. El **logo ENAC** del pie sí va incrustado (ver más abajo).

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
  word.ts               Construcción del informe Word (estructura oficial LaRUC)
  ucLogo.ts             Logo UC incrustado en base64 (generado, no editar a mano)
  enacLogo.ts           Logo ENAC del pie en base64 (generado, no editar a mano)
public/
  uc-logo.png           Logo mostrado en la web (provisional)
  enac-logo.png         Logo de acreditación ENAC del pie del informe
scripts/
  make-logo.py          Genera el logo provisional (Pillow)
  embed-logo.mjs        Incrusta public/uc-logo.png en lib/ucLogo.ts
  embed-enac.mjs        Incrusta public/enac-logo.png en lib/enacLogo.ts
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

## Logotipos del informe

- **Cabecera (LaRUC):** se compone como texto en `lib/word.ts` porque el logo
  oficial es un EMF que `docx` no puede incrustar. Si dispones de un PNG del
  logotipo LaRUC, puede añadirse de forma análoga al de ENAC (ver abajo).
- **Pie (ENAC):** el logo de acreditación va incrustado en `lib/enacLogo.ts`.
  Para cambiarlo, sustituye `public/enac-logo.png` y ejecuta
  `npm run embed-enac`.
- **Logo UC de la web:** sustituye `public/uc-logo.png` y ejecuta
  `npm run embed-logo` (regenera `lib/ucLogo.ts`).

Tras cambiar cualquier logo, haz commit del PNG y del `.ts` regenerado.

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
