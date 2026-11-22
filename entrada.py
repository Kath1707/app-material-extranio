"""
entrada.py
==========
Validación de lo que escribe y sube el usuario (la app es pública: el link lo
tienen personas ajenas a Calidad).

- Textos: se quitan caracteres de control, se recortan espacios y se limita el
  largo. Los textos que empiezan con "=", "+", "-" o "@" los guarda
  registro_excel como TEXTO, nunca como fórmula.
- Fotos: solo JPG/PNG reales (se mira el contenido, no la extensión), máximo
  5 MB de entrada, orientación EXIF corregida, transparencia sobre fondo blanco
  y miniatura JPEG recomprimida. La foto original NO se guarda en ningún lado.
"""

from __future__ import annotations

import io
import unicodedata

from PIL import Image, ImageOps, UnidentifiedImageError

# ---------------------------------------------------------------- textos
LARGO_MAXIMO = {
    "linea": 50,
    "producto": 100,
    "lote": 30,
    "especificacion": 100,
    "persona": 80,
    "observaciones": 500,
}


def limpiar_texto(valor, maximo: int, *, multilinea: bool = False) -> str:
    """Texto seguro para guardar: sin caracteres de control, sin espacios
    sobrantes y con un largo máximo."""
    if valor is None:
        return ""
    texto = unicodedata.normalize("NFC", str(valor))
    permitidos = {"\n"} if multilinea else set()
    texto = "".join(
        c if (c in permitidos or unicodedata.category(c)[0] != "C") else " "
        for c in texto.replace("\r\n", "\n")
    )
    if multilinea:
        texto = "\n".join(" ".join(l.split()) for l in texto.split("\n")).strip()
    else:
        texto = " ".join(texto.split())
    return texto[:maximo].strip()


# ---------------------------------------------------------------- fotos
MAX_BYTES_FOTO = 5 * 1024 * 1024
FORMATOS_FOTO = {"JPEG", "PNG"}
MAX_PIXELES_FOTO = 40_000_000          # ~ foto de 40 MP; más que eso se rechaza
CAJA_MINIATURA_EXCEL = (150, 120)      # ancho x alto máximos (px) de la miniatura en la celda K
CALIDAD_MINIATURA_EXCEL = 75


class FotoInvalida(ValueError):
    """Mensaje apto para mostrarse al usuario."""


def _abrir_foto(datos: bytes) -> Image.Image:
    if not datos:
        raise FotoInvalida("La foto está vacía. Vuelve a adjuntarla.")
    if len(datos) > MAX_BYTES_FOTO:
        raise FotoInvalida("La foto pesa más de 5 MB. Adjunta una más liviana.")
    try:
        img = Image.open(io.BytesIO(datos))
        if img.format not in FORMATOS_FOTO:
            raise FotoInvalida("La foto debe ser JPG o PNG.")
        ancho, alto = img.size
        if ancho * alto > MAX_PIXELES_FOTO:
            raise FotoInvalida("La foto tiene una resolución demasiado grande.")
        img.load()
    except FotoInvalida:
        raise
    except (UnidentifiedImageError, Image.DecompressionBombError, OSError, ValueError):
        raise FotoInvalida("No se pudo leer la foto. Adjunta un archivo JPG o PNG válido.") from None
    return img


def _a_rgb(img: Image.Image) -> Image.Image:
    img = ImageOps.exif_transpose(img)  # fotos de celular "acostadas"
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        fondo = Image.new("RGB", img.size, (255, 255, 255))
        fondo.paste(img, mask=img.getchannel("A"))
        return fondo
    return img.convert("RGB")


def miniatura_jpeg(datos: bytes, caja: tuple[int, int], calidad: int) -> bytes:
    """Valida la foto y devuelve una miniatura JPEG que cabe en `caja` (sin EXIF)."""
    img = _a_rgb(_abrir_foto(datos))
    img.thumbnail(caja, Image.LANCZOS)
    salida = io.BytesIO()
    img.save(salida, format="JPEG", quality=calidad, optimize=True)
    return salida.getvalue()


def miniatura_para_excel(datos: bytes) -> bytes:
    return miniatura_jpeg(datos, CAJA_MINIATURA_EXCEL, CALIDAD_MINIATURA_EXCEL)
