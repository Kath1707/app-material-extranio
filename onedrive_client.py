"""
onedrive_client.py
==================
Cliente mínimo de Microsoft Graph para guardar el historial de la app de
Control de materiales extraños en archivos Excel dentro de la carpeta "Calidad"
del OneDrive de la cuenta corporativa de María Almenara.

(Copiado de Kath1707/formatos-digitales-calidad; mantener ambos en sintonía.)

Cómo funciona (todo con msal + requests + openpyxl):
  1. Se pide un token a Microsoft Entra ID con credenciales de aplicación
     (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET).
  2. Para guardar: se DESCARGA el Excel del mes, se edita con openpyxl y se
     SUBE de vuelta con control de versión (ETag / If-Match). Si alguien
     guardó en medio (412) o el archivo está bloqueado por Excel (423), se
     reintenta sobre la versión nueva, así nunca se pisa el trabajo de otro.
  3. Si el archivo del mes no existe, se crea copiando la "Plantilla" de esa
     carpeta (a diferencia de Google, aquí la app SÍ puede crear archivos).

Secrets de Streamlit (nivel superior, sin sección):
    AZURE_TENANT_ID     = "..."
    AZURE_CLIENT_ID     = "..."
    AZURE_CLIENT_SECRET = "..."          # el VALOR del secreto, no su ID
    ONEDRIVE_USER       = "tu_correo@empresa.pe"
    ONEDRIVE_ROOT       = "Calidad"      # opcional, por defecto "Calidad"
    BACKEND             = "google"       # "onedrive" para activar este módulo
    BACKEND_MATERIAL    = "onedrive"     # opcional: tiene prioridad sobre BACKEND

Nunca se imprime ni se registra el secreto; los mensajes de error solo llevan
la descripción pública que devuelve Microsoft. Aun así, esos mensajes incluyen
rutas y nombres internos: la app pública los manda SOLO al log y al usuario le
muestra un texto genérico.
"""

from __future__ import annotations

import io
import os
import time
import unicodedata
from urllib.parse import quote

import requests
import streamlit as st
from openpyxl import load_workbook

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = ["https://graph.microsoft.com/.default"]
LIMITE_SUBIDA_SIMPLE = 4 * 1024 * 1024  # Graph: PUT simple hasta 4 MB
TIMEOUT = 60


# ----------------------------------------------------------------------------
# ERRORES
# ----------------------------------------------------------------------------
class OneDriveError(RuntimeError):
    """Error genérico de OneDrive (mensaje apto para mostrarse al usuario)."""


class ArchivoNoEncontrado(OneDriveError):
    pass


class Conflicto(OneDriveError):
    """412 (otro guardó antes), 409 (ya existe) o 423 (archivo bloqueado)."""


class ArchivoOcupado(OneDriveError):
    """Se agotaron los reintentos por conflicto (el Excel está abierto o en uso)."""


# ----------------------------------------------------------------------------
# CONFIGURACIÓN (Secrets de Streamlit o variables de entorno)
# ----------------------------------------------------------------------------
def _secret(nombre: str, defecto=None):
    valor = os.environ.get(nombre)
    if valor not in (None, ""):
        return valor
    try:
        valor = st.secrets.get(nombre, defecto)
    except Exception:  # sin secrets.toml (tests, ejecución local)
        return defecto
    return defecto if valor in (None, "") else valor


def configurado() -> bool:
    return all(
        _secret(k)
        for k in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "ONEDRIVE_USER")
    )


def activo(app: str | None = None) -> bool:
    """True si esta app debe guardar en OneDrive.

    Lee BACKEND_<APP> (si existe) y, si no, BACKEND. Por defecto "google", así
    nada cambia hasta que se active a propósito en los Secrets. Si se pidió
    OneDrive pero faltan credenciales, igual devuelve True: el error de
    configuración se verá al guardar en vez de volver en silencio a Google.
    """
    if app:
        valor = _secret(f"BACKEND_{app.upper()}")
        if valor:
            return str(valor).strip().lower() == "onedrive"
    return str(_secret("BACKEND", "google")).strip().lower() == "onedrive"


def _raiz() -> str:
    return str(_secret("ONEDRIVE_ROOT", "Calidad")).strip().strip("/")


def _usuario() -> str:
    return str(_secret("ONEDRIVE_USER", "")).strip()


# ----------------------------------------------------------------------------
# TOKEN
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _app_msal():
    import msal

    tenant = _secret("AZURE_TENANT_ID")
    return msal.ConfidentialClientApplication(
        _secret("AZURE_CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{tenant}",
        client_credential=_secret("AZURE_CLIENT_SECRET"),
    )


def _token() -> str:
    if not configurado():
        raise OneDriveError(
            "Faltan credenciales de Microsoft en los Secrets de Streamlit "
            "(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, ONEDRIVE_USER)."
        )
    resultado = _app_msal().acquire_token_for_client(scopes=SCOPE)  # msal usa su caché
    if "access_token" not in resultado:
        detalle = str(resultado.get("error_description") or resultado.get("error") or "error desconocido")
        raise OneDriveError(f"No se pudo autenticar con Microsoft: {detalle.splitlines()[0][:300]}")
    return resultado["access_token"]


# ----------------------------------------------------------------------------
# HTTP con reintentos
# ----------------------------------------------------------------------------
def _http(metodo: str, url: str, *, headers: dict | None = None, reintentos: int = 4, **kw):
    for intento in range(reintentos + 1):
        h = {"Authorization": f"Bearer {_token()}"}
        h.update(headers or {})
        try:
            resp = requests.request(metodo, url, headers=h, timeout=TIMEOUT, **kw)
        except requests.RequestException as e:
            if intento == reintentos:
                raise OneDriveError(f"No se pudo conectar con OneDrive: {e}") from e
            time.sleep(1.5 * (intento + 1))
            continue
        if resp.status_code in (429, 500, 502, 503, 504) and intento < reintentos:
            try:
                espera = float(resp.headers.get("Retry-After", 0))
            except ValueError:
                espera = 0
            time.sleep(min(max(espera, 1.5 * (intento + 1)), 20))
            continue
        return resp
    return resp  # pragma: no cover


def _error_graph(resp, contexto: str) -> OneDriveError:
    try:
        detalle = resp.json().get("error", {}).get("message", "")
    except Exception:
        detalle = ""
    if resp.status_code in (401, 403):
        detalle = (detalle or "sin permiso") + (
            " — revisa que la app de Azure tenga permisos Files.ReadWrite.All / "
            "Sites.ReadWrite.All con consentimiento de administrador."
        )
    return OneDriveError(f"{contexto} (HTTP {resp.status_code}): {detalle[:300]}")


# ----------------------------------------------------------------------------
# RUTAS
# ----------------------------------------------------------------------------
def _norm(texto) -> str:
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.casefold().split())


def _url_item(*partes: str, sufijo: str = "") -> str:
    ruta = "/".join(quote(p, safe="") for p in [_raiz(), *[p for p in partes if p]])
    return f"{GRAPH}/users/{quote(_usuario())}/drive/root:/{ruta}{sufijo}"


# ----------------------------------------------------------------------------
# LISTADO Y BÚSQUEDA (insensible a mayúsculas y acentos)
# ----------------------------------------------------------------------------
def _listar_hijos(*partes: str) -> list[dict]:
    """Lista el contenido de una carpeta (ruta relativa a ONEDRIVE_ROOT)."""
    url = _url_item(*partes, sufijo=":/children") + "?$select=name,size,eTag,folder,file&$top=200"
    items: list[dict] = []
    while url:
        resp = _http("GET", url)
        if resp.status_code == 404:
            raise ArchivoNoEncontrado(
                f"No encontré la carpeta '{'/'.join([_raiz(), *partes])}' en OneDrive."
            )
        if resp.status_code != 200:
            raise _error_graph(resp, "No se pudo listar la carpeta de OneDrive")
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


_cache_carpetas: dict[str, tuple[float, str]] = {}


def resolver_carpeta(nombre: str) -> str:
    """Nombre REAL (con sus acentos y mayúsculas) de una subcarpeta de la raíz."""
    clave = _norm(nombre)
    guardado = _cache_carpetas.get(clave)
    if guardado and time.time() - guardado[0] < 600:
        return guardado[1]
    for it in _listar_hijos():
        if "folder" in it and _norm(it["name"]) == clave:
            _cache_carpetas[clave] = (time.time(), it["name"])
            return it["name"]
    raise ArchivoNoEncontrado(
        f"No encontré la subcarpeta '{nombre}' dentro de '{_raiz()}' en OneDrive. "
        "Verifica el nombre exacto de la carpeta."
    )


def buscar_archivo(carpeta: str, nombre: str) -> dict | None:
    """Archivo por nombre (sin distinguir acentos/mayúsculas; la extensión .xlsx es opcional)."""
    real = resolver_carpeta(carpeta)
    objetivo = _norm(nombre)
    objetivos = {objetivo, objetivo + ".xlsx"} if not objetivo.endswith(".xlsx") else {objetivo}
    for it in _listar_hijos(real):
        if "file" in it and _norm(it["name"]) in objetivos:
            it["_carpeta"] = real
            return it
    return None


def buscar_plantilla(carpeta: str) -> dict:
    real = resolver_carpeta(carpeta)
    candidatos = sorted(
        (it for it in _listar_hijos(real)
         if "file" in it and "plantilla" in _norm(it["name"]) and _norm(it["name"]).endswith(".xlsx")),
        key=lambda it: it["name"],
    )
    if not candidatos:
        raise ArchivoNoEncontrado(
            f"No encontré la plantilla (un .xlsx con 'Plantilla' en el nombre) en la carpeta '{real}'."
        )
    candidatos[0]["_carpeta"] = real
    return candidatos[0]


# ----------------------------------------------------------------------------
# DESCARGA / SUBIDA
# ----------------------------------------------------------------------------
def descargar(carpeta_real: str, nombre_real: str) -> tuple[bytes, str]:
    """Devuelve (bytes, etag). El ETag se lee ANTES del contenido: si el archivo
    cambia entre ambas llamadas, la subida falla con 412 y se reintenta (nunca
    se pisa una versión más nueva)."""
    resp = _http("GET", _url_item(carpeta_real, nombre_real) + "?$select=eTag,size")
    if resp.status_code == 404:
        raise ArchivoNoEncontrado(f"No encontré '{nombre_real}' en la carpeta '{carpeta_real}'.")
    if resp.status_code != 200:
        raise _error_graph(resp, f"No se pudo leer '{nombre_real}'")
    etag = resp.json().get("eTag", "")
    cont = _http("GET", _url_item(carpeta_real, nombre_real, sufijo=":/content"))
    if cont.status_code != 200:
        raise _error_graph(cont, f"No se pudo descargar '{nombre_real}'")
    return cont.content, etag


def subir(carpeta_real: str, nombre: str, datos: bytes, *, etag: str | None = None, crear: bool = False) -> None:
    """Sube el archivo. crear=True falla (Conflicto) si ya existe; con etag, falla si cambió."""
    headers = {}
    if etag and not crear:
        headers["If-Match"] = etag
    if len(datos) > LIMITE_SUBIDA_SIMPLE:
        return _subir_por_sesion(carpeta_real, nombre, datos, headers, crear)

    url = _url_item(carpeta_real, nombre, sufijo=":/content")
    if crear:
        url += "?@microsoft.graph.conflictBehavior=fail"
    resp = _http("PUT", url, headers={**headers, "Content-Type": "application/octet-stream"}, data=datos)
    _verificar_subida(resp, nombre)


def _subir_por_sesion(carpeta_real, nombre, datos, headers, crear):
    cuerpo = {"item": {"@microsoft.graph.conflictBehavior": "fail" if crear else "replace"}}
    resp = _http(
        "POST", _url_item(carpeta_real, nombre, sufijo=":/createUploadSession"),
        headers=headers, json=cuerpo,
    )
    _verificar_subida(resp, nombre, esperado=(200,))
    destino = resp.json()["uploadUrl"]
    n = len(datos)
    # La URL de la sesión ya va firmada: se sube SIN el encabezado Authorization.
    r = requests.put(
        destino, data=datos, timeout=300,
        headers={"Content-Length": str(n), "Content-Range": f"bytes 0-{n - 1}/{n}"},
    )
    _verificar_subida(r, nombre, esperado=(200, 201))


def _verificar_subida(resp, nombre: str, esperado=(200, 201)) -> None:
    if resp.status_code in esperado:
        return
    if resp.status_code in (409, 412, 423):
        raise Conflicto(f"Conflicto al guardar '{nombre}' (HTTP {resp.status_code}).")
    raise _error_graph(resp, f"No se pudo guardar '{nombre}'")


# ----------------------------------------------------------------------------
# API DE ALTO NIVEL
# ----------------------------------------------------------------------------
def editar_libro(carpeta: str, nombre: str, funcion, *, crear_desde_plantilla: bool = True, intentos: int = 6):
    """Descarga el libro `nombre` (o lo crea desde la plantilla), ejecuta
    `funcion(workbook)` sobre él y lo sube de vuelta. Devuelve lo que devuelva
    `funcion`.

    `funcion` se vuelve a ejecutar sobre una descarga fresca si hay conflicto,
    así que debe depender solo del libro que recibe (no de estado externo).
    """
    ultimo = None
    for intento in range(intentos):
        try:
            info = buscar_archivo(carpeta, nombre)
            if info is None:
                if not crear_desde_plantilla:
                    raise ArchivoNoEncontrado(f"No encontré '{nombre}' en la carpeta '{carpeta}'.")
                plantilla = buscar_plantilla(carpeta)
                datos, _ = descargar(plantilla["_carpeta"], plantilla["name"])
                carpeta_real, nombre_real, etag, crear = plantilla["_carpeta"], nombre, None, True
                if not nombre_real.lower().endswith(".xlsx"):
                    nombre_real += ".xlsx"
            else:
                carpeta_real, nombre_real, crear = info["_carpeta"], info["name"], False
                datos, etag = descargar(carpeta_real, nombre_real)

            libro = load_workbook(io.BytesIO(datos))
            resultado = funcion(libro)
            salida = io.BytesIO()
            libro.save(salida)
            subir(carpeta_real, nombre_real, salida.getvalue(), etag=etag, crear=crear)
            return resultado
        except Conflicto as e:
            ultimo = e
            time.sleep(min(1.5 * (intento + 1), 8))
    raise ArchivoOcupado(
        f"No se pudo guardar en '{nombre}' después de {intentos} intentos: el archivo está "
        "abierto o siendo editado en Excel por otra persona. Ciérralo e inténtalo de nuevo "
        "(no se perdió nada de lo ya guardado)."
    ) from ultimo


def leer_libro(carpeta: str, nombre: str):
    """Libro abierto en modo lectura, o None si el archivo todavía no existe."""
    info = buscar_archivo(carpeta, nombre)
    if info is None:
        return None
    datos, _ = descargar(info["_carpeta"], info["name"])
    return load_workbook(io.BytesIO(datos), read_only=True, data_only=True)


def texto_celda(v) -> str:
    """Convierte el valor de una celda a texto como lo hacía gspread.get_all_values()."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v)


def valores_hoja(ws) -> list[list[str]]:
    return [[texto_celda(c) for c in fila] for fila in ws.iter_rows(values_only=True)]


def leer_hojas(carpeta: str, nombre: str, titulos: tuple | list) -> dict[str, list[list[str]]]:
    """{titulo_de_hoja: valores} de las hojas pedidas que existan. {} si no existe el libro."""
    libro = leer_libro(carpeta, nombre)
    if libro is None:
        return {}
    try:
        return {t: valores_hoja(libro[t]) for t in titulos if t in libro.sheetnames}
    finally:
        libro.close()


@st.cache_data(ttl=20, show_spinner=False)
def leer_hojas_cacheado(carpeta: str, nombre: str, titulos: tuple) -> dict[str, list[list[str]]]:
    """Igual que leer_hojas pero con caché de 20 s (el wizard vuelve a ejecutarse en cada clic)."""
    return leer_hojas(carpeta, nombre, titulos)
