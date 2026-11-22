"""Ejecuta la página real de Streamlit (AppTest) con OneDrive simulado."""
import io
from datetime import date

import openpyxl
import pytest
from streamlit.testing.v1 import AppTest

from conftest import CARPETA, RAIZ, foto_prueba

APP = str(RAIZ / "streamlit_app.py")
FECHA = date(2026, 11, 22)
LIBRO = "Noviembre 2026 - Material extranio.xlsx"


@pytest.fixture
def onedrive_activo(nube, monkeypatch):
    monkeypatch.setenv("BACKEND_MATERIAL", "onedrive")
    monkeypatch.chdir(RAIZ)
    return nube


def llenar_formulario(at, *, foto=None, producto="Muffin de vainilla"):
    at.date_input(key="fecha_0").set_value(FECHA)
    at.selectbox(key="linea_0").set_value("Muffin")
    at.text_input(key="producto_0").set_value(producto)
    at.text_input(key="lote_0").set_value("326")
    at.number_input(key="batch_0").set_value(4)
    at.checkbox(key="check_PRECINTO_0").check()
    at.checkbox(key="check_BOLSA_0").check()
    at.run()
    at.number_input(key="cantidad_PRECINTO_0").set_value(4)
    at.radio(key="confirmacion_PRECINTO_0").set_value("Sí")
    at.selectbox(key="tipo_BOLSA_0").set_value("Leche")
    at.number_input(key="cantidad_BOLSA_0").set_value(3)
    at.radio(key="confirmacion_BOLSA_0").set_value("No")
    at.text_input(key="ejecutor_0").set_value("  Ana Pérez ")
    at.text_input(key="responsable_0").set_value("Luis Soto")
    at.text_area(key="obs_0").set_value("=1+1 revisar")
    if foto is not None:
        at.file_uploader(key="foto_0").set_value(("evidencia.jpg", foto, "image/jpeg"))
    at.run()


def registrar(at):
    at.button(key="registrar").click().run()
    assert not at.exception, at.exception


def libro(nube):
    return openpyxl.load_workbook(io.BytesIO(nube.archivos[f"Calidad/{CARPETA}/{LIBRO}"]["datos"]))


def test_registro_con_foto_se_guarda_en_onedrive(onedrive_activo):
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    llenar_formulario(at, foto=foto_prueba(3000, 2000, orientacion=6))
    registrar(at)
    assert not at.error, [e.value for e in at.error]
    assert at.success and "2 elemento(s)" in at.success[0].value
    # el mensaje de éxito no revela carpetas ni nombres de archivo internos
    assert "xlsx" not in at.success[0].value and "Calidad" not in at.success[0].value

    ws = libro(onedrive_activo)["2026-11-22"]
    enc = [ws.cell(4, c).value for c in range(1, 14)]
    esperado = {
        "FECHA": "22/11/2026", "LÍNEA": "Muffin", "PRODUCTO": "Muffin de vainilla",
        "LOTE JULIANO": 326, "CANT. BATCH": 4, "ELEMENTO CONTROLADO": "PRECINTO",
        "ESPECIFICACIÓN / DETALLE": None, "CANT. ELEMENTOS\nCONTROLADOS": 4,
        "CONFORMIDAD": "Conforme", "OBSERVACIONES": "=1+1 revisar", "FOTO (miniatura)": None,
        "EJECUTOR": "Ana Pérez", "RESPONSABLE DE VALIDACIÓN": "Luis Soto",
    }
    fila5 = {h: ws.cell(5, c).value for c, h in enumerate(enc, start=1)}
    fila5["FECHA"] = fila5["FECHA"].strftime("%d/%m/%Y")
    assert fila5 == esperado
    assert ws["J5"].data_type == "s"                                       # no es fórmula
    assert ws["F6"].value == "BOLSA" and ws["G6"].value == "Leche"
    assert ws["H6"].value == 3 and ws["I6"].value == "No conforme"
    fotos = sorted((i.anchor._from.row + 1, i.anchor._from.col + 1) for i in ws._images if i.anchor._from.row)
    assert fotos == [(5, 11), (6, 11)]                                      # miniatura en K de cada elemento
    assert sum(1 for m, _, _ in onedrive_activo.llamadas if m == "PUT") == 1   # una sola subida


def test_registro_sin_foto(onedrive_activo):
    at = AppTest.from_file(APP, default_timeout=30).run()
    llenar_formulario(at)
    registrar(at)
    assert not at.error, [e.value for e in at.error]
    ws = libro(onedrive_activo)["2026-11-22"]
    assert ws["C5"].value == "Muffin de vainilla" and len(ws._images) == 1   # solo el logo


def test_foto_invalida_se_rechaza_sin_tocar_onedrive(onedrive_activo):
    at = AppTest.from_file(APP, default_timeout=30).run()
    llenar_formulario(at, foto=b"esto no es una foto")
    n = len(onedrive_activo.llamadas)
    registrar(at)
    assert [e.value for e in at.error] == ["No se pudo leer la foto. Adjunta un archivo JPG o PNG válido."]
    assert len(onedrive_activo.llamadas) == n


def test_textos_largos_se_recortan(onedrive_activo):
    at = AppTest.from_file(APP, default_timeout=30).run()
    llenar_formulario(at, producto="P" * 400)
    registrar(at)
    assert not at.error
    assert libro(onedrive_activo)["2026-11-22"]["C5"].value == "P" * 100


@pytest.mark.parametrize("codigo", [403, 500])
def test_error_de_microsoft_muestra_mensaje_generico(onedrive_activo, codigo):
    onedrive_activo.fallar_con = codigo
    at = AppTest.from_file(APP, default_timeout=30).run()
    llenar_formulario(at)
    registrar(at)
    assert len(at.error) == 1
    texto = at.error[0].value
    assert texto.startswith("No se pudo guardar el registro.")
    for prohibido in ("HTTP", "Calidad/", "Control material", "OneDrive", "graph", "drive", "Azure",
                      "Access denied", "interno", "xlsx", "Traceback", str(codigo)):
        assert prohibido.lower() not in texto.lower(), prohibido


def test_con_google_por_defecto_no_toca_onedrive(nube, monkeypatch):
    """Sin BACKEND/BACKEND_MATERIAL la app sigue usando Google (aquí falla por no
    tener credenciales, con mensaje genérico) y no llama a OneDrive."""
    monkeypatch.delenv("BACKEND", raising=False)
    monkeypatch.delenv("BACKEND_MATERIAL", raising=False)
    monkeypatch.chdir(RAIZ)
    at = AppTest.from_file(APP, default_timeout=30).run()
    llenar_formulario(at, foto=foto_prueba())
    registrar(at)
    assert nube.llamadas == []
    assert [e.value for e in at.error] == [
        "No se pudo guardar el registro. Inténtalo de nuevo en unos minutos; "
        "si el problema continúa, avisa al área de Calidad."
    ]
