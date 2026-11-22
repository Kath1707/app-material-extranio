"""Simula Microsoft Graph en memoria para probar sin internet ni credenciales."""
import re
import sys
from pathlib import Path
from urllib.parse import unquote

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
PLANTILLAS = Path(__file__).resolve().parent / "plantillas"

# carpeta de OneDrive -> archivos de plantilla que contiene
ESTRUCTURA = {
    "Control material extranio": ["Control materiales extranios - Plantilla base.xlsx"],
}
CARPETA = "Control material extranio"
PLANTILLA = "Control materiales extranios - Plantilla base.xlsx"


def foto_prueba(ancho=800, alto=600, formato="JPEG", orientacion=None) -> bytes:
    """Foto sintética: mitad izquierda roja, derecha azul (sirve para ver la orientación)."""
    import io
    from PIL import Image

    img = Image.new("RGB", (ancho, alto), (200, 30, 30))
    img.paste((30, 60, 200), (ancho // 2, 0, ancho, alto))
    buf = io.BytesIO()
    if orientacion:
        exif = Image.Exif()
        exif[0x0112] = orientacion
        img.save(buf, format=formato, exif=exif)
    else:
        img.save(buf, format=formato)
    return buf.getvalue()


class Resp:
    def __init__(self, status=200, datos=None, contenido=b"", headers=None):
        self.status_code, self._datos, self.content, self.headers = status, datos, contenido, headers or {}

    def json(self):
        return self._datos


class FakeGraph:
    def __init__(self):
        self.archivos = {}   # "Calidad/Carpeta/archivo.xlsx" -> {"datos": bytes, "v": int}
        self.llamadas = []
        self.fallar_con = None  # código HTTP con el que fallan las llamadas (simula Graph caído/sin permiso)
        self.interferir = None  # callable(ruta) -> se ejecuta justo antes de aceptar un PUT (simula otro usuario)
        for carpeta, nombres in ESTRUCTURA.items():
            for n in nombres:
                self.archivos[f"Calidad/{carpeta}/{n}"] = {"datos": (PLANTILLAS / n).read_bytes(), "v": 1}

    def etag(self, ruta):
        return f'"{{ABC}},{self.archivos[ruta]["v"]}"'

    def http(self, metodo, url, headers=None, reintentos=0, **kw):
        base = re.match(r".*/drive/root:/(.*)$", url).group(1)
        ruta, _, query = base.partition("?")
        ruta = unquote(ruta)
        sufijo = ""
        for s in (":/children", ":/content", ":/createUploadSession"):
            if ruta.endswith(s):
                ruta, sufijo = ruta[: -len(s)], s
        self.llamadas.append((metodo, ruta, sufijo))
        headers = headers or {}
        if self.fallar_con:
            return Resp(self.fallar_con, {"error": {"message": "Access denied to /users/x/drive (interno)"}})

        if sufijo == ":/children":
            prefijo = ruta + "/"
            hijos, carpetas = [], set()
            for k, a in self.archivos.items():
                if k.startswith(prefijo):
                    resto = k[len(prefijo):]
                    if "/" in resto:
                        carpetas.add(resto.split("/")[0])
                    else:
                        hijos.append({"name": resto, "file": {}, "size": len(a["datos"]), "eTag": self.etag(k)})
            hijos += [{"name": c, "folder": {}} for c in sorted(carpetas)]
            if not hijos:
                return Resp(404, {"error": {"message": "no existe"}})
            return Resp(200, {"value": hijos})

        if metodo == "GET" and sufijo == "":
            if ruta not in self.archivos:
                return Resp(404, {"error": {"message": "no existe"}})
            return Resp(200, {"eTag": self.etag(ruta), "size": len(self.archivos[ruta]["datos"])})

        if metodo == "GET" and sufijo == ":/content":
            if ruta not in self.archivos:
                return Resp(404, {"error": {"message": "no existe"}})
            return Resp(200, contenido=self.archivos[ruta]["datos"])

        if metodo == "PUT" and sufijo == ":/content":
            if self.interferir:
                f, self.interferir = self.interferir, None
                f(ruta)
            existe = ruta in self.archivos
            if "conflictBehavior=fail" in query and existe:
                return Resp(409, {"error": {"message": "ya existe"}})
            if "If-Match" in headers and (not existe or headers["If-Match"] != self.etag(ruta)):
                return Resp(412, {"error": {"message": "precondición"}})
            v = self.archivos[ruta]["v"] + 1 if existe else 1
            self.archivos[ruta] = {"datos": kw["data"], "v": v}
            return Resp(200 if existe else 201, {})
        raise AssertionError(f"llamada no simulada: {metodo} {url}")


@pytest.fixture
def nube(monkeypatch):
    import onedrive_client as od

    g = FakeGraph()
    monkeypatch.setenv("ONEDRIVE_USER", "tu_correo@empresa.pe")
    monkeypatch.setenv("ONEDRIVE_ROOT", "Calidad")
    monkeypatch.setattr(od, "_http", g.http)
    monkeypatch.setattr(od, "time", type("T", (), {"sleep": staticmethod(lambda s: None), "time": staticmethod(__import__("time").time)}))
    od._cache_carpetas.clear()
    od.leer_hojas_cacheado.clear()
    return g
