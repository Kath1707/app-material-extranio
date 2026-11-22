"""
Control de Materiales Extraños — María Almenara
=================================================

App Streamlit para registrar el control de empaques y materiales en contacto
con el producto (precinto, bolsa, cintillo, envoltura, tapitas, frasquitos,
otros), siguiendo el mismo patrón que las apps PT / PI / Cocina Dulce /
Desinfección de huevo:
- Carpeta en Drive: "Control de materiales extraños" (misma raíz ROOT_FOLDER_ID
  que las otras carpetas)
- Spreadsheet mensual: "<Mes> <Año> - Control materiales" (creado manualmente
  por Kath cada mes, p. ej. "Noviembre 2026 - Control materiales")
- Cada FECHA crea/usa una hoja (tab) distinta dentro del spreadsheet mensual,
  duplicando la plantilla "Hoja 1" la primera vez que se usa esa fecha.
- Cada registro nuevo se INSERTA (nunca se sobrescribe) justo encima de la
  fila con el marcador "VºBº COORDINADORA DE CALIDAD", para no pisar el pie
  de página.

OneDrive (Excel) — se activa SOLO desde los Secrets con BACKEND_MATERIAL o
BACKEND = "onedrive" (ver MIGRACION_ONEDRIVE.md):
- Carpeta "Calidad/Control material extranio"; libro del mes
  "<Mes> <Año> - Material extranio.xlsx", creado solo desde la plantilla.
- Primero llena las filas libres 5 a 16 de la hoja del día y luego inserta
  encima de la leyenda. La foto va como miniatura en la columna K.

Seguridad (la app es pública): la app solo ESCRIBE. Nunca muestra, lista ni
descarga para el usuario nada de OneDrive/Drive, y ante un error muestra un
mensaje genérico; el detalle técnico va solo al log.
"""

import base64
import logging
from datetime import date

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

import entrada
import onedrive_client as od
import registro_excel as rx

log = logging.getLogger("material_extranio")

st.set_page_config(page_title="Control de Materiales Extraños", page_icon="🧵", layout="centered")

# ──────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────────
CARPETA_NOMBRE = "Control de materiales extraños"  # subcarpeta dentro de ROOT_FOLDER_ID
HOJA_PLANTILLA = "Hoja 1"
MARCADOR_PIE = "VºBº COORDINADORA DE CALIDAD"

LINEAS = ["Panadería", "Dosimetría", "Amasados", "Batidos", "Galletón", "Muffin", "Cheesecake", "OTROS"]
ELEMENTOS = ["PRECINTO", "BOLSA", "CINTILLO", "ENVOLTURA", "TAPITAS", "FRASQUITOS", "OTROS"]
TIPOS_BOLSA = ["Polietileno", "Leche", "Poligrass", "Otro"]

# Solo Google: límite de 50,000 caracteres por celda. Una miniatura comprimida
# a este tamaño/calidad entra holgadamente dentro de ese límite.
MAX_THUMBNAIL_SIDE = 300
JPEG_QUALITY = 50

ETIQUETA_RESPONSABLE_VALIDACION = "Responsable que valida el registro"  # cambiar aquí si prefieres otra palabra

# Mismo orden que los encabezados (fila 4, A:M) de la plantilla:
# FECHA | LÍNEA | PRODUCTO | LOTE JULIANO | CANT. BATCH | ELEMENTO CONTROLADO |
# ESPECIFICACIÓN / DETALLE | CANT. ELEMENTOS CONTROLADOS | CONFORMIDAD |
# OBSERVACIONES | FOTO (miniatura) | EJECUTOR | RESPONSABLE DE VALIDACIÓN
COLUMNS = [
    "FECHA", "LINEA", "PRODUCTO", "LOTE_JULIANO", "CANTIDAD_BATCH",
    "ELEMENTO", "ESPECIFICACION", "CANTIDAD_CONTROLADA", "CONFORMIDAD",
    "OBSERVACIONES", "FOTO_BASE64", "EJECUTOR", "RESPONSABLE_VALIDACION",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


# ──────────────────────────────────────────────────────────────────────────
# CONEXIÓN A GOOGLE DRIVE / SHEETS (mismo patrón que tus otras apps)
# ──────────────────────────────────────────────────────────────────────────

@st.cache_resource
def conectar_gspread():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


def obtener_carpeta_id(gc, nombre_carpeta, root_folder_id):
    """Busca una subcarpeta por nombre dentro de ROOT_FOLDER_ID."""
    resp = gc.http_client.request(
        "get",
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": (
                f"'{root_folder_id}' in parents and "
                f"name = '{nombre_carpeta}' and "
                "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            ),
            "fields": "files(id, name)",
        },
    )
    archivos = resp.json().get("files", [])
    if not archivos:
        raise FileNotFoundError(
            f"No se encontró la carpeta '{nombre_carpeta}' dentro de la carpeta raíz. "
            "Verifica el nombre exacto y que la cuenta de servicio tenga acceso."
        )
    return archivos[0]["id"]


def obtener_spreadsheet_mensual(gc, carpeta_id, nombre_spreadsheet):
    """Busca (no crea) el spreadsheet mensual dentro de la carpeta de materiales extraños."""
    resp = gc.http_client.request(
        "get",
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": (
                f"'{carpeta_id}' in parents and "
                f"name = '{nombre_spreadsheet}' and "
                "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
            ),
            "fields": "files(id, name)",
        },
    )
    archivos = resp.json().get("files", [])
    if not archivos:
        raise FileNotFoundError(
            f"No se encontró el spreadsheet '{nombre_spreadsheet}' dentro de la carpeta "
            f"'{CARPETA_NOMBRE}'. Confirma que el nombre coincide exactamente, o créalo si ya es un mes nuevo."
        )
    return gc.open_by_key(archivos[0]["id"])


def nombre_spreadsheet_mensual(fecha: date) -> str:
    return f"{MESES_ES[fecha.month]} {fecha.year} - Control materiales"


def obtener_o_crear_hoja_del_dia(spreadsheet, fecha: date):
    """Devuelve la hoja (tab) del día, duplicando la plantilla si todavía no existe."""
    nombre_tab = fecha.strftime("%Y-%m-%d")
    try:
        return spreadsheet.worksheet(nombre_tab)
    except gspread.WorksheetNotFound:
        plantilla = spreadsheet.worksheet(HOJA_PLANTILLA)
        return plantilla.duplicate(new_sheet_name=nombre_tab)


def encontrar_fila_marcador(worksheet):
    """Ubica la fila del marcador 'VºBº COORDINADORA DE CALIDAD' en la columna A."""
    valores_col_a = worksheet.col_values(1)
    for idx, valor in enumerate(valores_col_a, start=1):
        if MARCADOR_PIE in (valor or ""):
            return idx
    raise ValueError(
        f"No se encontró el marcador '{MARCADOR_PIE}' en la hoja. "
        "Verifica que la plantilla no haya cambiado de texto."
    )


def comprimir_imagen_a_base64(datos_foto: bytes) -> str:
    miniatura = entrada.miniatura_jpeg(
        datos_foto, (MAX_THUMBNAIL_SIDE, MAX_THUMBNAIL_SIDE), JPEG_QUALITY
    )
    return base64.b64encode(miniatura).decode("utf-8")


def guardar_filas(worksheet, filas_valores, fila_destino):
    """Inserta cada fila justo encima del marcador, empujándolo hacia abajo,
    manteniendo el orden de los elementos marcados."""
    for i, valores in enumerate(filas_valores):
        worksheet.insert_row(valores, index=fila_destino + i, value_input_option="USER_ENTERED")


def guardar_en_google(fecha: date, filas_valores: list) -> None:
    gc = conectar_gspread()
    root_folder_id = st.secrets["ROOT_FOLDER_ID"]
    carpeta_id = obtener_carpeta_id(gc, CARPETA_NOMBRE, root_folder_id)
    spreadsheet = obtener_spreadsheet_mensual(gc, carpeta_id, nombre_spreadsheet_mensual(fecha))
    worksheet = obtener_o_crear_hoja_del_dia(spreadsheet, fecha)
    fila_destino = encontrar_fila_marcador(worksheet)
    guardar_filas(worksheet, filas_valores, fila_destino)


def guardar_en_onedrive(fecha: date, filas_valores: list, miniatura: bytes | None) -> None:
    """Una sola descarga/subida del libro del mes para todos los elementos marcados.
    La miniatura (si hay) se ancla en la columna K de cada fila del registro."""
    imagenes = [{rx.COLUMNA_FOTO_MATERIAL: miniatura} if miniatura else None for _ in filas_valores]
    rx.guardar_filas_con_firma(
        rx.MATERIAL, fecha, fecha.strftime("%Y-%m-%d"), filas_valores, imagenes=imagenes
    )


MENSAJE_ERROR_GENERICO = (
    "No se pudo guardar el registro. Inténtalo de nuevo en unos minutos; "
    "si el problema continúa, avisa al área de Calidad."
)
MENSAJE_ARCHIVO_OCUPADO = (
    "No se pudo guardar porque el registro del mes está siendo usado en este momento. "
    "Espera un minuto e inténtalo de nuevo (no se perdió nada de lo ya guardado)."
)


# ──────────────────────────────────────────────────────────────────────────
# INTERFAZ STREAMLIT
# ──────────────────────────────────────────────────────────────────────────

st.title("CONTROL DE MATERIALES EXTRAÑOS")
st.caption(
    "Registro de verificación de empaques y materiales en contacto con el producto, "
    "para evitar la presencia de materiales extraños en el producto terminado."
)

LARGO = entrada.LARGO_MAXIMO
MAX_CANTIDAD = 100_000

if "form_key" not in st.session_state:
    st.session_state.form_key = 0

k = st.session_state.form_key  # sufijo de las keys; cambia al pulsar "Nuevo registro"

st.subheader("Datos generales")
col1, col2 = st.columns(2)
with col1:
    fecha = st.date_input("Fecha", value=None, format="DD/MM/YYYY", key=f"fecha_{k}")
    linea = st.selectbox("Línea", LINEAS, index=None, placeholder="Seleccione la línea", key=f"linea_{k}")
    if linea == "OTROS":
        linea = st.text_input(
            "Especifique la línea", max_chars=LARGO["linea"], key=f"linea_otros_{k}"
        )
with col2:
    producto = st.text_input("Producto a elaborar", max_chars=LARGO["producto"], key=f"producto_{k}")
    lote_juliano = st.text_input("Lote juliano", max_chars=LARGO["lote"], key=f"lote_{k}")

cantidad_batch = st.number_input(
    "Cantidad de batch elaborados", min_value=0, max_value=MAX_CANTIDAD, step=1, key=f"batch_{k}"
)

st.divider()
st.subheader("Elementos a controlar")
st.caption("Marque los elementos que se verificaron en este control.")

datos_elementos = {}
for elemento in ELEMENTOS:
    marcado = st.checkbox(elemento, key=f"check_{elemento}_{k}")
    if marcado:
        with st.container(border=True):
            especificacion = ""
            if elemento == "BOLSA":
                tipo_bolsa = st.selectbox(
                    "Especifique el tipo de bolsa", TIPOS_BOLSA, key=f"tipo_{elemento}_{k}"
                )
                if tipo_bolsa == "Otro":
                    especificacion = st.text_input(
                        "Detalle el tipo de bolsa", max_chars=LARGO["especificacion"],
                        key=f"detalle_{elemento}_{k}",
                    )
                else:
                    especificacion = tipo_bolsa
            elif elemento == "OTROS":
                especificacion = st.text_input(
                    "Especifique el elemento", max_chars=LARGO["especificacion"],
                    key=f"detalle_{elemento}_{k}",
                )

            cantidad_controlada = st.number_input(
                f"Cantidad de {elemento.lower()} controlados",
                min_value=0, max_value=MAX_CANTIDAD, step=1, key=f"cantidad_{elemento}_{k}",
            )

            confirmacion = st.radio(
                "¿Confirma que la cantidad controlada corresponde a la cantidad de batch elaborados?",
                ["Sí", "No"], index=None, key=f"confirmacion_{elemento}_{k}", horizontal=True,
            )

            datos_elementos[elemento] = {
                "especificacion": entrada.limpiar_texto(especificacion, LARGO["especificacion"]),
                "cantidad_controlada": cantidad_controlada,
                "confirmacion": confirmacion,
            }

st.divider()
st.subheader("Responsables")
col3, col4 = st.columns(2)
with col3:
    ejecutor = st.text_input(
        "Ejecutor (operario que guarda/aparta los elementos para su conteo)",
        max_chars=LARGO["persona"], key=f"ejecutor_{k}",
    )
with col4:
    responsable_validacion = st.text_input(
        ETIQUETA_RESPONSABLE_VALIDACION, max_chars=LARGO["persona"], key=f"responsable_{k}"
    )

st.divider()
observaciones = st.text_area(
    "Observaciones (opcional)", max_chars=LARGO["observaciones"], key=f"obs_{k}"
)
foto = st.file_uploader(
    "Foto de evidencia (opcional, JPG o PNG de hasta 5 MB)",
    type=["jpg", "jpeg", "png"], key=f"foto_{k}",
)

# Todo lo escrito pasa por aquí antes de validarse y guardarse (largo máximo,
# sin caracteres de control). El widget ya limita el largo; esto es la defensa
# del lado del servidor.
linea = entrada.limpiar_texto(linea, LARGO["linea"])
producto = entrada.limpiar_texto(producto, LARGO["producto"])
lote_juliano = entrada.limpiar_texto(lote_juliano, LARGO["lote"])
ejecutor = entrada.limpiar_texto(ejecutor, LARGO["persona"])
responsable_validacion = entrada.limpiar_texto(responsable_validacion, LARGO["persona"])
observaciones = entrada.limpiar_texto(observaciones, LARGO["observaciones"], multilinea=True)

st.divider()
col_guardar, col_nuevo = st.columns(2)

with col_guardar:
    if st.button("💾 Registrar control", type="primary", use_container_width=True, key="registrar"):
        errores = []
        if fecha is None:
            errores.append("Seleccione la fecha.")
        if not linea:
            errores.append("Seleccione la línea.")
        if not producto:
            errores.append("Ingrese el producto.")
        if not lote_juliano:
            errores.append("Ingrese el lote juliano.")
        if not datos_elementos:
            errores.append("Marque al menos un elemento a controlar.")
        if not ejecutor:
            errores.append("Ingrese el ejecutor.")
        if not responsable_validacion:
            errores.append(f"Ingrese el {ETIQUETA_RESPONSABLE_VALIDACION.lower()}.")
        for elemento, datos in datos_elementos.items():
            if datos["confirmacion"] is None:
                errores.append(f"Confirme la conformidad de {elemento}.")
            if elemento in ("BOLSA", "OTROS") and not datos["especificacion"]:
                errores.append(f"Especifique el detalle de {elemento}.")

        usar_onedrive = od.activo("material")
        foto_b64, miniatura = "", None
        if not errores and foto is not None:
            try:
                datos_foto = foto.getvalue()
                if usar_onedrive:
                    miniatura = entrada.miniatura_para_excel(datos_foto)
                else:
                    foto_b64 = comprimir_imagen_a_base64(datos_foto)
            except entrada.FotoInvalida as e:
                errores.append(str(e))

        if errores:
            for e in errores:
                st.error(e)
        else:

            filas = []
            for elemento, datos in datos_elementos.items():
                filas.append({
                    "FECHA": fecha.strftime("%d/%m/%Y"),
                    "LINEA": linea,
                    "PRODUCTO": producto,
                    "LOTE_JULIANO": lote_juliano,
                    "CANTIDAD_BATCH": cantidad_batch,
                    "ELEMENTO": elemento,
                    "ESPECIFICACION": datos["especificacion"],
                    "CANTIDAD_CONTROLADA": datos["cantidad_controlada"],
                    "CONFORMIDAD": "Conforme" if datos["confirmacion"] == "Sí" else "No conforme",
                    "OBSERVACIONES": observaciones,
                    "FOTO_BASE64": foto_b64,
                    "EJECUTOR": ejecutor,
                    "RESPONSABLE_VALIDACION": responsable_validacion,
                })
            filas_valores = [[fila[col] for col in COLUMNS] for fila in filas]

            try:
                with st.spinner("Guardando el registro..."):
                    if usar_onedrive:
                        guardar_en_onedrive(fecha, filas_valores, miniatura)
                    else:
                        guardar_en_google(fecha, filas_valores)
            except od.ArchivoOcupado:
                log.warning("Libro del mes ocupado al guardar material extraño", exc_info=True)
                st.error(MENSAJE_ARCHIVO_OCUPADO)
            except Exception:
                # El detalle (rutas, nombres de archivo, respuesta de Microsoft/Google)
                # va SOLO al log de Streamlit Cloud; nunca a la pantalla.
                log.exception("Error al guardar el control de materiales extraños")
                st.error(MENSAJE_ERROR_GENERICO)
            else:
                st.success(
                    f"Se registraron {len(filas)} elemento(s) del {fecha.strftime('%d/%m/%Y')}."
                )
                if foto is not None:
                    st.image(foto, caption="Vista previa de la foto adjuntada", width=200)

with col_nuevo:
    if st.button("➕ Nuevo registro", use_container_width=True):
        st.session_state.form_key += 1
        st.rerun()
