"""Genera un logo provisional "UC" en public/uc-logo.png.

Es un marcador de posición hasta disponer del logotipo oficial de la
Universidad de Cantabria. Para usar el oficial: sustituye public/uc-logo.png
por el archivo definitivo (PNG) y ejecuta `npm run embed-logo`.

Requiere: pip install pillow
Uso: python scripts/make-logo.py  (desde la raíz del proyecto)
"""

from PIL import Image, ImageDraw, ImageFont

TEAL = (13, 154, 169, 255)  # #0D9AA9, tono corporativo aproximado
FONT_PATH = "C:/Windows/Fonts/arialbd.ttf"  # ajustar en macOS/Linux
FONT_SIZE = 170
PADDING = 22

font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
probe = ImageDraw.Draw(Image.new("RGBA", (1, 1)))
bbox = probe.textbbox((0, 0), "UC", font=font)
width, height = bbox[2] - bbox[0], bbox[3] - bbox[1]

image = Image.new("RGBA", (width + PADDING * 2, height + PADDING * 2), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.text((PADDING - bbox[0], PADDING - bbox[1]), "UC", font=font, fill=TEAL)
image.save("public/uc-logo.png")
print(f"public/uc-logo.png generado: {image.size[0]}x{image.size[1]} px")
