"""
registro_excel.py
=================
Lógica de "libro mensual + una pestaña por día" sobre Excel (openpyxl).
Copiado de Kath1707/formatos-digitales-calidad (mantener ambos en sintonía) y
ampliado para esta app con:
  - `Formato.filas_firma`: cuántas filas reservadas para la firma hay justo
    encima de la leyenda (1 en Huevo/Naranja/Tamizado, 0 en Material extraño,
    cuyas filas libres llegan hasta la fila anterior a la leyenda);
  - miniaturas de fotos ancladas en una celda (`imagenes=`), dentro de la MISMA
    operación de editar_libro, es decir, en una sola subida.

Dos modos, según cómo está armada cada plantilla:

* modo "footer"  (PT, PI, Cocina): la hoja del día se crea copiando "Hoja 1",
  se borran las filas de ejemplo entre el encabezado y el pie, y cada guardado
  INSERTA las filas nuevas justo encima del pie (CONSIDERACIONES / V°B°).

* modo "firma"   (Huevo, Naranja, Tamizado, Material extraño): primero se llenan las filas libres
  que trae la plantilla entre el encabezado y la fila de la firma; cuando ya no
  hay, se INSERTA una fila nueva justo encima de la firma.

Lo que openpyxl NO hace por sí solo y aquí sí se resuelve:
  - copiar una hoja con su logo (copy_worksheet pierde las imágenes),
  - insertar / borrar filas moviendo también celdas combinadas, alturas de fila
    e imágenes (si no, el pie de firma queda desalineado),
  - dar formato de tabla (bordes) a las filas nuevas,
  - convertir texto numérico y fechas a valores reales de Excel (lo que Google
    hacía con USER_ENTERED).
"""

from __future__ import annotations

import io
import re
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime

from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU

import onedrive_client as od

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
_LADO = Side(style="thin")
BORDE_FINO = Border(left=_LADO, right=_LADO, top=_LADO, bottom=_LADO)


# ----------------------------------------------------------------------------
# CONFIGURACIÓN POR APP
# ----------------------------------------------------------------------------
@dataclass(frozen=True)
class Formato:
    app: str                       # clave para BACKEND_<APP>
    carpeta: str                   # subcarpeta dentro de ONEDRIVE_ROOT
    sufijo: str = ""               # " - PT" → "Noviembre 2026 - PT.xlsx"
    modo: str = "footer"           # "footer" | "firma"
    marcador: str = "vb_calidad"   # "vb_calidad" | "consideraciones" | "coordinadora"
    hoja_plantilla: str = "Hoja 1"
    fila_encabezado: int = 4
    filas_firma: int = 1           # modo firma: filas reservadas entre los datos y la leyenda

    def nombre_libro(self, fecha: date) -> str:
        return f"{MESES_ES[fecha.month]} {fecha.year}{self.sufijo}.xlsx"


# ----------------------------------------------------------------------------
# DETECCIÓN DEL PIE / FIRMA
# ----------------------------------------------------------------------------
def _compacto(valor) -> str:
    t = str(valor if valor is not None else "").upper()
    # ° º ª se quitan ANTES de normalizar: NFKD convertiría 'º' en 'o' ("VºB" → "VOB")
    for ch in ("°", "º", "ª", " ", "\n", "\t", "."):
        t = t.replace(ch, "")
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def es_vb_calidad(valor) -> bool:
    """'V°B° Jefe de Calidad', 'VºB COORDINADORA DE CALIDAD', 'VB° Supervisor de Calidad'..."""
    t = _compacto(valor)
    return (
        t.startswith("VB") and "CALIDAD" in t and len(t) <= 45
        and any(k in t for k in ("COORDINADORA", "JEFE", "SUPERVISOR"))
    )


def es_consideraciones(valor) -> bool:
    return _compacto(valor).startswith("CONSIDERACIONES")


def es_coordinadora(valor) -> bool:
    t = _compacto(valor)
    return "COORDINADORADECALIDAD" in t and len(t) <= 45


_MARCADORES = {
    "vb_calidad": es_vb_calidad,
    "consideraciones": es_consideraciones,
    "coordinadora": es_coordinadora,
}


def encontrar_marcador(ws, fmt: Formato) -> int | None:
    """Primera fila, por debajo del encabezado, que contiene el marcador del pie."""
    es_marcador = _MARCADORES[fmt.marcador]
    for fila in ws.iter_rows(min_row=fmt.fila_encabezado + 1):
        for c in fila:
            if c.value is not None and es_marcador(c.value):
                return c.row
    return None


# ----------------------------------------------------------------------------
# OPERACIONES SOBRE FILAS (con merges, alturas e imágenes)
# ----------------------------------------------------------------------------
def _mover_alturas(ws, desde: int, delta: int) -> None:
    previas = {
        r: (d.height, d.hidden)
        for r, d in list(ws.row_dimensions.items())
        if r >= desde and (d.height is not None or d.hidden)
    }
    for r in previas:
        ws.row_dimensions[r].height = None
        ws.row_dimensions[r].hidden = False
    for r, (alto, oculta) in previas.items():
        nueva = r + delta
        if nueva >= 1:
            ws.row_dimensions[nueva].height = alto
            ws.row_dimensions[nueva].hidden = oculta


def _mover_imagenes(ws, desde_fila: int, delta: int) -> None:
    """Desplaza las imágenes ancladas en `desde_fila` (1-based) o más abajo."""
    for img in ws._images:
        a = img.anchor
        origen = getattr(a, "_from", None)
        if origen is not None and origen.row >= desde_fila - 1:
            origen.row += delta
            fin = getattr(a, "to", None)
            if fin is not None:
                fin.row += delta


def insertar_filas(ws, fila: int, n: int = 1) -> None:
    """Inserta n filas vacías en `fila`, empujando hacia abajo todo lo demás
    (celdas, combinadas, alturas, imágenes)."""
    rangos = list(ws.merged_cells.ranges)
    for r in rangos:
        ws.merged_cells.remove(r)
    ws.insert_rows(fila, n)
    for r in rangos:
        if r.min_row >= fila:
            r.shift(row_shift=n)
        elif r.max_row >= fila:  # combinada que cruza el punto de inserción: se estira
            r.expand(down=n)
        ws.merged_cells.add(r)
    _mover_alturas(ws, fila, n)
    _mover_imagenes(ws, fila, n)


def eliminar_filas(ws, desde: int, hasta: int) -> None:
    """Borra las filas desde..hasta (incluidas) y sube todo lo de abajo."""
    n = hasta - desde + 1
    if n <= 0:
        return
    rangos = list(ws.merged_cells.ranges)
    for r in rangos:
        ws.merged_cells.remove(r)
    # alturas de las filas borradas
    for r in range(desde, hasta + 1):
        ws.row_dimensions[r].height = None
    ws.delete_rows(desde, n)
    for r in rangos:
        if r.max_row < desde:
            ws.merged_cells.add(r)
        elif r.min_row > hasta:
            r.shift(row_shift=-n)
            ws.merged_cells.add(r)
        # las que tocan las filas borradas se descartan (sus celdas ya no existen)
    _mover_alturas(ws, hasta + 1, -n)
    _mover_imagenes(ws, hasta + 1, -n)


# ----------------------------------------------------------------------------
# HOJAS
# ----------------------------------------------------------------------------
def _bytes_imagen(img) -> bytes:
    """Bytes de una imagen SIN usar Image._data(), que cierra el archivo original y
    haría fallar el guardado posterior del libro."""
    ref = img.ref
    if hasattr(ref, "read"):
        ref.seek(0)
        datos = ref.read()
        ref.seek(0)
        return datos
    with open(ref, "rb") as f:
        return f.read()


def copiar_hoja(wb, origen, titulo: str):
    """Copia `origen` como una hoja nueva `titulo` (la primera del libro, para
    que lo más reciente quede a la vista) incluyendo el logo."""
    nueva = wb.copy_worksheet(origen)
    nueva.title = titulo[:31]
    for img in origen._images:
        copia = XLImage(io.BytesIO(_bytes_imagen(img)))
        copia.anchor = copy(img.anchor)
        copia.width, copia.height = img.width, img.height
        nueva.add_image(copia)
    nueva.sheet_view.zoomScale = origen.sheet_view.zoomScale
    nueva.sheet_view.showGridLines = origen.sheet_view.showGridLines
    if origen.print_area:
        nueva.print_area = origen.print_area
    # lo más nuevo primero (como en los historiales de Google); la plantilla queda al final
    wb.move_sheet(nueva, offset=-(len(wb.sheetnames) - 1))
    wb.active = 0
    for hoja in wb.worksheets:  # evita que Excel abra con varias hojas "agrupadas"
        hoja.sheet_view.tabSelected = hoja is wb.active
    return nueva


def asegurar_hoja_del_dia(wb, fmt: Formato, titulo: str):
    """(hoja, es_nueva). Copia la plantilla si la hoja del día todavía no existe."""
    titulo = titulo[:31]
    if titulo in wb.sheetnames:
        return wb[titulo], False
    if fmt.hoja_plantilla not in wb.sheetnames:
        raise od.OneDriveError(
            f"El libro no tiene la hoja plantilla '{fmt.hoja_plantilla}'. "
            "Revisa que la plantilla de la carpeta no haya cambiado de nombre."
        )
    return copiar_hoja(wb, wb[fmt.hoja_plantilla], titulo), True


# ----------------------------------------------------------------------------
# VALORES Y ESTILO
# ----------------------------------------------------------------------------
_RE_ENTERO_O_DECIMAL = re.compile(r"^-?(0|[1-9]\d*)([.,]\d+)?$")
_RE_FECHA = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")


def _convertir(valor):
    """Texto → valor tipado, como hacía Google con USER_ENTERED."""
    if not isinstance(valor, str):
        return valor
    t = valor.strip()
    if _RE_ENTERO_O_DECIMAL.match(t):
        n = float(t.replace(",", "."))
        return int(n) if n.is_integer() and "." not in t and "," not in t else n
    m = _RE_FECHA.match(t)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return valor
    return valor


def _poner_valor(celda, valor) -> None:
    v = _convertir(valor)
    celda.value = v
    if isinstance(v, str) and v.startswith(("=", "+", "-", "@")):
        celda.data_type = "s"  # texto libre del usuario: nunca interpretarlo como fórmula
    elif isinstance(v, datetime):
        celda.number_format = "DD/MM/YYYY"


def _ancho_tabla(ws, fila_encabezado: int) -> int:
    ult = 0
    for c in ws[fila_encabezado]:
        if c.value not in (None, ""):
            ult = max(ult, c.column)
    return ult


def _estilos_fila_modelo(wb, fmt: Formato, ncols: int) -> list:
    """Estilos de la primera fila de datos de la hoja PLANTILLA (siempre intacta)."""
    modelo = wb[fmt.hoja_plantilla]
    fila = fmt.fila_encabezado + 1
    return [copy(modelo.cell(row=fila, column=c)._style) for c in range(1, ncols + 1)]


def escribir_fila(ws, fila: int, valores: list, estilos: list) -> None:
    for i, valor in enumerate(valores, start=1):
        celda = ws.cell(row=fila, column=i)
        if i <= len(estilos):
            celda._style = copy(estilos[i - 1])
        celda.border = BORDE_FINO
        if valor is not None and valor != "":
            _poner_valor(celda, valor)
        else:
            celda.value = None


def _fila_libre(ws, fila: int, ncols: int) -> bool:
    return all(
        ws.cell(row=fila, column=c).value is None or not str(ws.cell(row=fila, column=c).value).strip()
        for c in range(1, ncols + 1)
    )


# ----------------------------------------------------------------------------
# GUARDADO
# ----------------------------------------------------------------------------
def _ajustar_columnas_extra(ws, fmt: Formato, filas: list, extras=("BRIX",)) -> list:
    """Si la hoja trae una columna de encabezado que la app ya no genera
    (p. ej. 'Grados Brix'), deja un hueco ahí para que el resto no se corra."""
    encabezado = [str(c.value or "") for c in ws[fmt.fila_encabezado]]
    huecos = [i for i, h in enumerate(encabezado) if any(x in h.upper() for x in extras)]
    if not huecos:
        return filas
    salida = []
    for fila in filas:
        nueva = list(fila)
        for pos in sorted(huecos):
            nueva.insert(pos, "")
        salida.append(nueva)
    return salida


def guardar_filas(fmt: Formato, fecha: date, titulo_hoja: str, filas: list, *, ajustar_extras: bool = False) -> str:
    """Modo 'footer': inserta `filas` encima del pie en la hoja `titulo_hoja`
    del libro del mes de `fecha`. Devuelve el nombre del libro."""
    nombre = fmt.nombre_libro(fecha)

    def operacion(wb):
        ws, nueva = asegurar_hoja_del_dia(wb, fmt, titulo_hoja)
        ncols = _ancho_tabla(wb[fmt.hoja_plantilla], fmt.fila_encabezado)
        estilos = _estilos_fila_modelo(wb, fmt, ncols)

        pie = encontrar_marcador(ws, fmt)
        if nueva and pie and pie > fmt.fila_encabezado + 1:
            eliminar_filas(ws, fmt.fila_encabezado + 1, pie - 1)  # filas de ejemplo de la plantilla
            pie = encontrar_marcador(ws, fmt)

        datos = _ajustar_columnas_extra(ws, fmt, filas) if ajustar_extras else filas
        if pie is None:  # sin pie: se agrega al final de la tabla
            destino = max(fmt.fila_encabezado + 1, ws.max_row + 1)
        else:
            destino = pie
            insertar_filas(ws, destino, len(datos))
        for i, fila in enumerate(datos):
            escribir_fila(ws, destino + i, fila, estilos)

    od.editar_libro(fmt.carpeta, nombre, operacion)
    _invalidar_cache()
    return nombre


def _ubicar_fila_firma(ws, fmt: Formato, ncols: int, fila: list, estilos: list) -> int:
    """Escribe `fila` en la primera fila libre sobre la firma o, si no hay, en una
    fila nueva insertada encima de ella. Devuelve la fila usada."""
    leyenda = encontrar_marcador(ws, fmt)
    if leyenda is None:
        raise od.OneDriveError(
            "No se encontró la leyenda 'COORDINADORA DE CALIDAD' en la hoja. "
            "Verifica que la plantilla no haya cambiado de estructura."
        )
    fila_firma = max(leyenda - fmt.filas_firma, fmt.fila_encabezado + 1)

    for r in range(fmt.fila_encabezado + 1, fila_firma):
        if _fila_libre(ws, r, ncols):
            escribir_fila(ws, r, fila, estilos)
            return r
    insertar_filas(ws, fila_firma, 1)
    escribir_fila(ws, fila_firma, fila, estilos)
    return fila_firma


def guardar_filas_con_firma(
    fmt: Formato, fecha: date, titulo_hoja: str, filas: list, *, imagenes: list | None = None
) -> tuple[str, list[int]]:
    """Modo 'firma' para VARIAS filas en una sola descarga/subida.

    `imagenes` (opcional) va alineada con `filas`: para cada fila, None o un
    dict {columna (1-based): bytes de la imagen YA reducida}. Devuelve
    (nombre del libro, filas usadas)."""
    nombre = fmt.nombre_libro(fecha)
    imagenes = list(imagenes or [])

    def operacion(wb):
        ws, _ = asegurar_hoja_del_dia(wb, fmt, titulo_hoja)
        ncols = _ancho_tabla(wb[fmt.hoja_plantilla], fmt.fila_encabezado)
        estilos = _estilos_fila_modelo(wb, fmt, ncols)
        usadas = []
        for i, fila in enumerate(filas):
            r = _ubicar_fila_firma(ws, fmt, ncols, fila, estilos)
            usadas.append(r)
            for col, datos in ((imagenes[i] if i < len(imagenes) else None) or {}).items():
                anclar_imagen(ws, r, col, datos)
        return usadas

    usadas = od.editar_libro(fmt.carpeta, nombre, operacion)
    _invalidar_cache()
    return nombre, usadas


def guardar_fila_con_firma(
    fmt: Formato, fecha: date, titulo_hoja: str, fila: list, *, imagenes: dict | None = None
) -> tuple[str, int]:
    """Modo 'firma': llena la primera fila libre sobre la firma; si no hay, inserta
    una encima de ella. `imagenes` = {columna (1-based): bytes}. Devuelve
    (nombre del libro, fila usada)."""
    nombre, usadas = guardar_filas_con_firma(fmt, fecha, titulo_hoja, [fila], imagenes=[imagenes])
    return nombre, usadas[0]


# ----------------------------------------------------------------------------
# IMÁGENES EN CELDAS (miniaturas de evidencia)
# ----------------------------------------------------------------------------
MARGEN_IMAGEN_PX = 3


def anclar_imagen(ws, fila: int, columna: int, datos: bytes) -> XLImage:
    """Ancla una imagen (ya reducida) en la celda (fila, columna) y agranda la
    fila y la columna lo justo para que se vea completa.

    Se usa un OneCellAnchor explícito (y no la referencia "K5" en texto) para
    que `insertar_filas` pueda desplazarla si después se insertan filas por
    encima en la misma operación."""
    img = XLImage(io.BytesIO(datos))
    ancho, alto = img.width, img.height
    m = pixels_to_EMU(MARGEN_IMAGEN_PX)
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=columna - 1, colOff=m, row=fila - 1, rowOff=m),
        ext=XDRPositiveSize2D(pixels_to_EMU(ancho), pixels_to_EMU(alto)),
    )
    ws.add_image(img)

    # alto de fila en puntos (1 px = 0.75 pt) y ancho de columna en "caracteres" (~7 px)
    alto_pt = (alto + 2 * MARGEN_IMAGEN_PX) * 0.75
    dim = ws.row_dimensions[fila]
    dim.height = max(dim.height or 15, alto_pt)
    letra = get_column_letter(columna)
    ancho_car = (ancho + 2 * MARGEN_IMAGEN_PX) / 7 + 1
    col = ws.column_dimensions[letra]
    col.width = max(col.width or 8.43, ancho_car)
    return img


# ----------------------------------------------------------------------------
# LECTURA
# ----------------------------------------------------------------------------
def _invalidar_cache() -> None:
    try:
        od.leer_hojas_cacheado.clear()
    except Exception:
        pass


def leer_hojas_del_dia(
    fmt: Formato, fecha: date, titulos: list[str], fresco: bool = False
) -> dict[str, list[list[str]]]:
    """Valores (texto) de las hojas `titulos` del libro del mes. {} si el libro no existe aún.
    fresco=True salta la caché de 20 s (úsalo justo antes de guardar, para detectar duplicados)."""
    if fresco:
        return od.leer_hojas(fmt.carpeta, fmt.nombre_libro(fecha), tuple(titulos))
    return od.leer_hojas_cacheado(fmt.carpeta, fmt.nombre_libro(fecha), tuple(titulos))


def filas_de_datos(valores: list[list[str]], fmt: Formato) -> tuple[list[str], list[list[str]]]:
    """(encabezado, filas de datos hasta el pie) a partir de los valores de una hoja."""
    if len(valores) < fmt.fila_encabezado:
        return [], []
    encabezado = valores[fmt.fila_encabezado - 1]
    es_marcador = _MARCADORES[fmt.marcador]
    filas = []
    for fila in valores[fmt.fila_encabezado:]:
        if any(es_marcador(c) for c in fila if c):
            break
        filas.append(fila)
    return encabezado, filas


# ----------------------------------------------------------------------------
# FORMATOS DE CADA APP (nombres EXACTOS de las carpetas en OneDrive)
# ----------------------------------------------------------------------------
# Plantilla: "Control materiales extranios - Plantilla base.xlsx", hoja "Hoja 1",
# encabezado en la fila 4 (A:M), filas libres 5 a 16, leyenda "VºBº COORDINADORA
# DE CALIDAD" en A17:M17 y nota legal en A18:M18 (sin fila de firma aparte).
MATERIAL = Formato(
    "material", "Control material extranio", " - Material extranio",
    modo="firma", marcador="coordinadora", filas_firma=0,
)
COLUMNA_FOTO_MATERIAL = 11  # K: "FOTO (miniatura)"
