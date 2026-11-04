import streamlit as st
import pandas as pd
from PIL import Image
import io
import base64
import os
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

st.set_page_config(page_title="Control de Materiales Extraños", page_icon="🧵", layout="centered")

# ---------- Configuración ----------
LINEAS = ["Panadería", "Dosimetría", "Amasados", "Batidos", "Galletón", "Muffin", "Cheesecake"]
ELEMENTOS = ["PRECINTO", "BOLSA", "CINTILLO", "ENVOLTURA", "TAPITAS", "FRASQUITOS", "OTROS"]
TIPOS_BOLSA = ["Polietileno", "Leche", "Poligrass", "Otro"]

LOCAL_FILE = "registros_materiales_extranos.csv"  # solo como respaldo si falla la conexión a Sheets

HOJA_PLANTILLA = "Hoja 1"
MARCADOR_PIE = "VºBº COORDINADORA DE CALIDAD"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Límite de Google Sheets: 50,000 caracteres por celda.
# Una miniatura comprimida a este tamaño/calidad entra holgadamente dentro de ese límite.
MAX_THUMBNAIL_SIDE = 300  # px
JPEG_QUALITY = 50

ETIQUETA_RESPONSABLE_VALIDACION = "Responsable que valida el registro"  # cambiar aquí si prefieres otra palabra

COLUMNS = [
    "FECHA", "LINEA", "PRODUCTO", "LOTE_JULIANO", "CANTIDAD_BATCH",
    "ELEMENTO", "ESPECIFICACION", "CANTIDAD_CONTROLADA", "CONFORMIDAD",
    "OBSERVACIONES", "FOTO_BASE64", "EJECUTOR", "RESPONSABLE_VALIDACION",
]
# Este mismo orden es el que debe tener la fila de encabezados en la plantilla de Google Sheets.


def comprimir_imagen_a_base64(uploaded_file) -> str:
    """Comprime la imagen a una miniatura JPEG y la devuelve como texto base64,
    pensada para caber en el límite de 50,000 caracteres por celda de Google Sheets."""
    img = Image.open(uploaded_file)
    img = img.convert("RGB")
    img.thumbnail((MAX_THUMBNAIL_SIDE, MAX_THUMBNAIL_SIDE))
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=JPEG_QUALITY)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def guardar_filas_local(filas: list[dict]):
    """Respaldo local: solo se usa si falla la escritura en Google Sheets, para no perder el registro."""
    df_nuevo = pd.DataFrame(filas, columns=COLUMNS)
    if os.path.exists(LOCAL_FILE):
        df_existente = pd.read_csv(LOCAL_FILE)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
    df_final.to_csv(LOCAL_FILE, index=False)


@st.cache_resource(show_spinner=False)
def conectar_cliente_sheets():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds)


def obtener_spreadsheet():
    gc = conectar_cliente_sheets()
    return gc.open_by_key(st.secrets["SHEET_ID"])


def obtener_hoja_del_dia(sh, fecha) -> gspread.Worksheet:
    """Devuelve la pestaña del día (nombrada como la fecha elegida en el formulario);
    si no existe todavía, la crea duplicando la plantilla 'Hoja 1'."""
    nombre_hoja = fecha.strftime("%Y-%m-%d")
    try:
        return sh.worksheet(nombre_hoja)
    except WorksheetNotFound:
        plantilla = sh.worksheet(HOJA_PLANTILLA)
        return sh.duplicate_sheet(source_sheet_id=plantilla.id, new_sheet_name=nombre_hoja)


def insertar_filas_en_sheet(ws: gspread.Worksheet, filas: list[dict]):
    """Inserta las filas justo arriba del marcador de pie de página, empujándolo hacia abajo."""
    marcador = ws.find(MARCADOR_PIE)
    if marcador is None:
        raise RuntimeError(
            f"No se encontró el marcador '{MARCADOR_PIE}' en la hoja '{ws.title}'. "
            "Revisa que el pie de la plantilla no haya cambiado de texto."
        )
    valores = [[fila[col] for col in COLUMNS] for fila in filas]
    ws.insert_rows(valores, row=marcador.row, value_input_option="USER_ENTERED")


def guardar_filas_en_sheets(filas: list[dict], fecha):
    sh = obtener_spreadsheet()
    ws = obtener_hoja_del_dia(sh, fecha)
    insertar_filas_en_sheet(ws, filas)


# ---------- Interfaz ----------
st.title("CONTROL DE MATERIALES EXTRAÑOS")
st.caption(
    "Registro de verificación de empaques y materiales en contacto con el producto, "
    "para evitar la presencia de materiales extraños en el producto terminado."
)

with st.form("form_materiales_extranos", clear_on_submit=False):

    st.subheader("Datos generales")
    col1, col2 = st.columns(2)
    with col1:
        fecha = st.date_input("Fecha", value=None, format="DD/MM/YYYY")
        linea = st.selectbox("Línea", LINEAS, index=None, placeholder="Seleccione la línea")
    with col2:
        producto = st.text_input("Producto a elaborar")
        lote_juliano = st.text_input("Lote juliano")

    cantidad_batch = st.number_input("Cantidad de batch elaborados", min_value=0, step=1)

    st.divider()
    st.subheader("Elementos a controlar")
    st.caption("Marque los elementos que se verificaron en este control.")

    datos_elementos = {}
    for elemento in ELEMENTOS:
        marcado = st.checkbox(elemento)
        if marcado:
            with st.container(border=True):
                especificacion = ""
                if elemento == "BOLSA":
                    tipo_bolsa = st.selectbox(
                        "Especifique el tipo de bolsa", TIPOS_BOLSA, key=f"tipo_{elemento}"
                    )
                    if tipo_bolsa == "Otro":
                        especificacion = st.text_input(
                            "Detalle el tipo de bolsa", key=f"detalle_{elemento}"
                        )
                    else:
                        especificacion = tipo_bolsa
                elif elemento == "OTROS":
                    especificacion = st.text_input("Especifique el elemento", key=f"detalle_{elemento}")

                cantidad_controlada = st.number_input(
                    f"Cantidad de {elemento.lower()} controlados",
                    min_value=0, step=1, key=f"cantidad_{elemento}",
                )

                confirmacion = st.radio(
                    "¿Confirma que la cantidad controlada corresponde a la cantidad de batch elaborados?",
                    ["Sí", "No"], index=None, key=f"confirmacion_{elemento}", horizontal=True,
                )

                datos_elementos[elemento] = {
                    "especificacion": especificacion,
                    "cantidad_controlada": cantidad_controlada,
                    "confirmacion": confirmacion,
                }

    st.divider()
    st.subheader("Responsables")
    col3, col4 = st.columns(2)
    with col3:
        ejecutor = st.text_input("Ejecutor (operario que guarda/aparta los elementos para su conteo)")
    with col4:
        responsable_validacion = st.text_input(ETIQUETA_RESPONSABLE_VALIDACION)

    st.divider()
    observaciones = st.text_area("Observaciones (opcional)")
    foto = st.file_uploader("Foto de evidencia (opcional)", type=["jpg", "jpeg", "png"])

    enviado = st.form_submit_button("Registrar control")

if enviado:
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

    if errores:
        for e in errores:
            st.error(e)
    else:
        foto_b64 = comprimir_imagen_a_base64(foto) if foto is not None else ""

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

        try:
            guardar_filas_en_sheets(filas, fecha)
            st.success(f"Se registraron {len(filas)} elemento(s) en Google Sheets correctamente.")
        except Exception as e:
            guardar_filas_local(filas)
            st.warning(
                "No se pudo escribir en Google Sheets, así que el registro se guardó localmente "
                f"como respaldo (no se perdió). Detalle del error: {e}"
            )

        if foto is not None:
            st.image(foto, caption="Vista previa de la foto adjuntada", width=200)

st.divider()
if os.path.exists(LOCAL_FILE):
    with st.expander("Ver registros de respaldo local (solo si hubo fallas de conexión)"):
        st.dataframe(pd.read_csv(LOCAL_FILE))
