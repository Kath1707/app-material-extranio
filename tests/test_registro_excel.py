"""Pruebas con la plantilla REAL de material extraño y un Graph simulado."""
import io
from datetime import date

import openpyxl
import pytest
from PIL import Image

import entrada
import onedrive_client as od
import registro_excel as rx
from conftest import CARPETA, PLANTILLA, foto_prueba

HOY = date(2026, 11, 22)
TITULO = "2026-11-22"
LIBRO = "Noviembre 2026 - Material extranio.xlsx"
ENCABEZADO = [
    "FECHA", "LÍNEA", "PRODUCTO", "LOTE JULIANO", "CANT. BATCH", "ELEMENTO CONTROLADO",
    "ESPECIFICACIÓN / DETALLE", "CANT. ELEMENTOS\nCONTROLADOS", "CONFORMIDAD", "OBSERVACIONES",
    "FOTO (miniatura)", "EJECUTOR", "RESPONSABLE DE VALIDACIÓN",
]


def abrir(nube, nombre=LIBRO):
    return openpyxl.load_workbook(io.BytesIO(nube.archivos[f"Calidad/{CARPETA}/{nombre}"]["datos"]))


def fila_de(ws, predicado, desde=1):
    for f in ws.iter_rows(min_row=desde):
        for c in f:
            if c.value is not None and predicado(str(c.value)):
                return c.row


def es_leyenda(t):
    return "COORDINADORA DE CALIDAD" in t.upper()


def es_legal(t):
    return t.startswith("Se prohíbe")


def fila_ejemplo(i, foto=""):
    return ["22/11/2026", "Muffin", f"Producto {i}", "326", "4", "PRECINTO", "",
            "4", "Conforme", f"obs {i}", foto, "Ana Pérez", "Luis Soto"]


def fotos_de(ws):
    """[(fila 1-based, columna 1-based)] de las imágenes que NO son el logo."""
    return sorted((i.anchor._from.row + 1, i.anchor._from.col + 1) for i in ws._images
                  if (i.anchor._from.row, i.anchor._from.col) != (0, 0))


def comprobar_estructura(ws):
    """Logo en A1, combinadas sin solaparse."""
    logos = [i for i in ws._images if (i.anchor._from.row, i.anchor._from.col) == (0, 0)]
    assert len(logos) == 1
    rangos = list(ws.merged_cells.ranges)
    for i, a in enumerate(rangos):
        for b in rangos[i + 1:]:
            assert a.isdisjoint(b), f"combinadas solapadas {a} {b}"


# ------------------------------------------------------------ plantilla
def test_plantilla_tiene_la_estructura_esperada():
    from conftest import PLANTILLAS

    ws = openpyxl.load_workbook(PLANTILLAS / PLANTILLA)["Hoja 1"]
    assert [ws.cell(4, c).value for c in range(1, 14)] == ENCABEZADO
    assert ws.cell(4, 14).value is None
    assert all(ws.cell(r, c).value is None for r in range(5, 17) for c in range(1, 14))
    assert fila_de(ws, es_leyenda) == 17 and fila_de(ws, es_legal) == 18
    merges = {str(r) for r in ws.merged_cells.ranges}
    assert {"A17:M17", "A18:M18"} <= merges


# ------------------------------------------------ filas libres e inserción
def test_crea_libro_llena_5_a_16_y_luego_inserta_sobre_la_leyenda(nube):
    usadas = []
    for i in range(14):  # 12 libres (5..16) + 2 insertadas
        nombre, usada = rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(i))
        usadas.append(usada)
    assert nombre == LIBRO
    assert usadas == list(range(5, 19))

    wb = abrir(nube)
    assert wb.sheetnames == [TITULO, "Hoja 1"]                 # el día primero, la plantilla al final
    ws = wb[TITULO]
    comprobar_estructura(ws)
    for i in range(14):
        r = 5 + i
        assert ws.cell(r, 3).value == f"Producto {i}"
        assert ws.cell(r, 1).value.strftime("%d/%m/%Y") == "22/11/2026"   # fecha real de Excel
        assert ws.cell(r, 4).value == 326 and ws.cell(r, 5).value == 4
        assert ws.cell(r, 13).value == "Luis Soto"
        assert ws.cell(r, 1).border.left.style == "thin"

    # leyenda y nota legal bajaron 2 filas, con sus combinadas
    assert fila_de(ws, es_leyenda, desde=5) == 19
    assert fila_de(ws, es_legal) == 20
    merges = {str(r) for r in ws.merged_cells.ranges}
    assert {"A19:M19", "A20:M20"} <= merges
    assert not {"A17:M17", "A18:M18"} & merges

    # la plantilla quedó intacta
    pl = wb["Hoja 1"]
    assert fila_de(pl, es_leyenda) == 17 and not any(c.value for c in pl[5])


def test_varios_elementos_van_en_una_sola_subida(nube):
    filas = [fila_ejemplo(i) for i in range(3)]
    nombre, usadas = rx.guardar_filas_con_firma(rx.MATERIAL, HOY, TITULO, filas)
    assert usadas == [5, 6, 7]
    assert sum(1 for m, _, s in nube.llamadas if m == "PUT") == 1


def test_segundo_dia_en_el_mismo_libro(nube):
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(1))
    rx.guardar_fila_con_firma(rx.MATERIAL, date(2026, 11, 23), "2026-11-23", fila_ejemplo(2))
    wb = abrir(nube)
    assert wb.sheetnames == ["2026-11-23", TITULO, "Hoja 1"]
    assert sum(1 for w in wb.worksheets if w.sheet_view.tabSelected) == 1
    assert wb["2026-11-23"]["C5"].value == "Producto 2" and wb[TITULO]["C5"].value == "Producto 1"


# ------------------------------------------------------------------ fotos
def test_foto_queda_en_K_y_se_conserva_al_seguir_guardando(nube):
    mini = entrada.miniatura_para_excel(foto_prueba(1600, 1200))
    k = rx.COLUMNA_FOTO_MATERIAL

    # foto en la primera fila libre
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(0), imagenes={k: mini})
    ws = abrir(nube)[TITULO]
    assert fotos_de(ws) == [(5, 11)]
    alto_px = Image.open(io.BytesIO(mini)).height
    assert 100 <= alto_px <= 120                              # 1600x1200 → ~150x113 (cabe en 150x120)
    assert ws.row_dimensions[5].height >= alto_px * 0.75      # la fila creció para que se vea entera
    assert ws["K5"].value is None

    # se llenan el resto de libres (5..16) y se insertan 2 filas más, la última con foto
    for i in range(1, 13):
        rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(i))
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(13), imagenes={k: mini})
    ws = abrir(nube)[TITULO]
    assert fotos_de(ws) == [(5, 11), (18, 11)]               # la de arriba NO se movió
    assert fila_de(ws, es_leyenda, desde=5) == 19

    # un guardado más (inserta otra fila encima de la leyenda): ninguna foto se mueve ni se pierde
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(14))
    wb = abrir(nube)
    ws = wb[TITULO]
    comprobar_estructura(ws)
    assert fotos_de(ws) == [(5, 11), (18, 11)]
    assert fila_de(ws, es_leyenda, desde=5) == 20
    assert ws.row_dimensions[5].height >= alto_px * 0.75 and ws.row_dimensions[18].height >= alto_px * 0.75
    assert ws.column_dimensions["K"].width * 7 >= 150          # y la columna K también
    for img in ws._images:
        if img.anchor._from.row != 0:
            assert Image.open(io.BytesIO(img._data())).size == Image.open(io.BytesIO(mini)).size


def test_foto_en_fila_insertada_en_la_misma_operacion(nube):
    """Si en un mismo guardado hay que insertar filas, las fotos de ese mismo guardado
    quedan en su fila (anclaje explícito, se desplaza con insertar_filas)."""
    mini = entrada.miniatura_para_excel(foto_prueba())
    k = rx.COLUMNA_FOTO_MATERIAL
    for i in range(11):
        rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(i))      # 5..15
    filas = [fila_ejemplo(20 + i) for i in range(3)]                             # 16, 17*, 18*
    _, usadas = rx.guardar_filas_con_firma(rx.MATERIAL, HOY, TITULO, filas, imagenes=[{k: mini}] * 3)
    assert usadas == [16, 17, 18]
    ws = abrir(nube)[TITULO]
    assert fotos_de(ws) == [(16, 11), (17, 11), (18, 11)]
    assert [ws.cell(r, 3).value for r in (16, 17, 18)] == ["Producto 20", "Producto 21", "Producto 22"]


def test_la_foto_del_segundo_dia_no_toca_el_primero(nube):
    mini = entrada.miniatura_para_excel(foto_prueba())
    k = rx.COLUMNA_FOTO_MATERIAL
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(0), imagenes={k: mini})
    rx.guardar_fila_con_firma(rx.MATERIAL, date(2026, 11, 23), "2026-11-23", fila_ejemplo(1), imagenes={k: mini})
    wb = abrir(nube)
    assert fotos_de(wb[TITULO]) == [(5, 11)] and fotos_de(wb["2026-11-23"]) == [(5, 11)]
    assert fotos_de(wb["Hoja 1"]) == []                     # la plantilla sigue sin fotos


# ----------------------------------------------------------- concurrencia
def test_conflicto_412_se_reintenta_y_conserva_ambos_registros(nube):
    mini = entrada.miniatura_para_excel(foto_prueba())
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(1))

    def otro_usuario_guarda(ruta):
        # alguien guarda JUSTO antes de que llegue nuestra subida
        wb = openpyxl.load_workbook(io.BytesIO(nube.archivos[ruta]["datos"]))
        wb[TITULO]["C6"] = "registro-del-otro"
        buf = io.BytesIO(); wb.save(buf)
        nube.archivos[ruta] = {"datos": buf.getvalue(), "v": nube.archivos[ruta]["v"] + 1}

    nube.interferir = otro_usuario_guarda
    _, usada = rx.guardar_fila_con_firma(
        rx.MATERIAL, HOY, TITULO, fila_ejemplo(2), imagenes={rx.COLUMNA_FOTO_MATERIAL: mini}
    )
    ws = abrir(nube)[TITULO]
    assert ws["C5"].value == "Producto 1"
    assert ws["C6"].value == "registro-del-otro"          # lo del otro usuario NO se perdió
    assert usada == 7 and ws["C7"].value == "Producto 2"   # el nuestro fue a la siguiente libre
    assert fotos_de(ws) == [(7, 11)]                       # y su foto con él, una sola vez
    puts = [x for x in nube.llamadas if x[0] == "PUT"]
    assert len(puts) == 3                                  # 1º guardado + intento rechazado (412) + reintento


def test_si_el_archivo_sigue_ocupado_avisa_con_un_error_propio(nube, monkeypatch):
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(1))
    ruta = f"Calidad/{CARPETA}/{LIBRO}"

    original = nube.http

    def siempre_cambia(metodo, url, **kw):
        if metodo == "PUT":
            nube.archivos[ruta]["v"] += 1   # alguien lo edita en cada intento
        return original(metodo, url, **kw)

    monkeypatch.setattr(od, "_http", siempre_cambia)
    with pytest.raises(od.ArchivoOcupado):
        rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila_ejemplo(2))
    assert abrir(nube)[TITULO]["C5"].value == "Producto 1"


# ---------------------------------------------------------------- textos
def test_texto_libre_nunca_se_vuelve_formula(nube):
    fila = fila_ejemplo(0)
    fila[2], fila[6], fila[9], fila[11] = "=HYPERLINK(\"http://x\")", "+cuidado", "-obs", "@persona"
    fila[3] = "007"
    rx.guardar_fila_con_firma(rx.MATERIAL, HOY, TITULO, fila)
    ws = abrir(nube)[TITULO]
    assert ws["C5"].value == '=HYPERLINK("http://x")' and ws["C5"].data_type == "s"
    assert ws["G5"].value == "+cuidado" and ws["G5"].data_type == "s"
    assert ws["J5"].value == "-obs" and ws["L5"].value == "@persona"
    assert ws["D5"].value == "007"                         # el lote con ceros adelante se queda como texto
