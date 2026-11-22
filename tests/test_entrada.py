"""Validación de textos y fotos que llegan desde la app pública."""
import io

import pytest
from PIL import Image

import entrada
from conftest import foto_prueba


def abrir(datos):
    return Image.open(io.BytesIO(datos))


def test_miniatura_es_jpeg_pequena_y_cabe_en_la_celda():
    mini = entrada.miniatura_para_excel(foto_prueba(4000, 3000))
    img = abrir(mini)
    assert img.format == "JPEG"
    assert img.width <= 150 and img.height <= 120
    assert len(mini) < 30_000
    assert "exif" not in img.info                       # no se guardan metadatos (GPS, celular...)


def test_corrige_la_orientacion_exif():
    # orientación 6 = la cámara la guardó "acostada": hay que girarla 90° → queda vertical
    mini = entrada.miniatura_para_excel(foto_prueba(800, 600, orientacion=6))
    img = abrir(mini).convert("RGB")
    assert img.height > img.width
    # rojo (izquierda original) termina ARRIBA tras girar 90° en sentido horario
    r, g, b = img.getpixel((img.width // 2, 5))
    assert r > 150 and b < 100


def test_png_con_transparencia_queda_sobre_blanco():
    buf = io.BytesIO()
    Image.new("RGBA", (300, 200), (0, 0, 0, 0)).save(buf, format="PNG")
    img = abrir(entrada.miniatura_para_excel(buf.getvalue())).convert("RGB")
    assert min(img.getpixel((10, 10))) > 240


@pytest.mark.parametrize("datos,mensaje", [
    (b"", "vacía"),
    (b"x" * (5 * 1024 * 1024 + 1), "5 MB"),
    (b"no soy una imagen", "JPG o PNG"),
])
def test_rechaza_fotos_vacias_grandes_o_ilegibles(datos, mensaje):
    with pytest.raises(entrada.FotoInvalida, match=mensaje):
        entrada.miniatura_para_excel(datos)


def test_rechaza_otros_formatos_aunque_digan_jpg():
    buf = io.BytesIO()
    Image.new("RGB", (50, 50)).save(buf, format="GIF")
    with pytest.raises(entrada.FotoInvalida, match="JPG o PNG"):
        entrada.miniatura_para_excel(buf.getvalue())


def test_rechaza_resolucion_absurda(monkeypatch):
    monkeypatch.setattr(entrada, "MAX_PIXELES_FOTO", 1000)
    with pytest.raises(entrada.FotoInvalida, match="resolución"):
        entrada.miniatura_para_excel(foto_prueba(100, 100))


def test_limpiar_texto():
    assert entrada.limpiar_texto("  Ana \t Pérez\x00 ", 80) == "Ana Pérez"
    assert entrada.limpiar_texto("x" * 500, 100) == "x" * 100
    assert entrada.limpiar_texto(None, 10) == ""
    assert entrada.limpiar_texto("línea 1\r\n  línea 2 \x07", 50, multilinea=True) == "línea 1\nlínea 2"
    assert entrada.limpiar_texto("a\nb", 50) == "a b"
    assert entrada.limpiar_texto("=1+1", 10) == "=1+1"   # se conserva; registro_excel lo guarda como texto
