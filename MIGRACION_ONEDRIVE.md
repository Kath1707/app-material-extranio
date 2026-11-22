# Historial en OneDrive (Excel) — Control de materiales extraños

La app puede guardar en Google Sheets (valor por defecto) o en archivos Excel de la carpeta
**Calidad** del OneDrive corporativo, el mismo que usan las apps de
`Kath1707/formatos-digitales-calidad`. El cambio se hace **solo con los Secrets de Streamlit**
de esta app, sin tocar el código.

## Secrets (nivel superior, sin sección)

```toml
AZURE_TENANT_ID     = "..."
AZURE_CLIENT_ID     = "..."
AZURE_CLIENT_SECRET = "..."                  # el VALOR del secreto, no su ID
ONEDRIVE_USER       = "tu_correo@empresa.pe"
ONEDRIVE_ROOT       = "Calidad"              # opcional

BACKEND = "onedrive"      # o, solo para esta app: BACKEND_MATERIAL = "onedrive"
```

Mantén `ROOT_FOLDER_ID` y `[gcp_service_account]` mientras quieras poder volver a Google.
**Para volver a Google**: borra la línea `BACKEND...` o ponla en `"google"`.

## Cómo se guarda

| Qué | Dónde |
|---|---|
| Carpeta | `Calidad/Control material extranio` |
| Plantilla (debe estar en esa carpeta) | `Control materiales extranios - Plantilla base.xlsx` |
| Archivo del mes (se crea solo copiando la plantilla) | `<Mes> <Año> - Material extranio.xlsx` |
| Hoja (pestaña) | `AAAA-MM-DD`, copiada de `Hoja 1` |

- Se llenan primero las filas libres 5 a 16; después se inserta una fila encima de la leyenda
  "VºBº COORDINADORA DE CALIDAD", que baja junto con la nota legal.
- Cada elemento marcado es una fila. Todos los de un mismo registro se guardan en **una sola**
  subida.
- **Foto**: se acepta solo JPG/PNG de hasta 5 MB. Se corrige la orientación del celular, se
  quitan los metadatos (GPS, modelo) y se guarda una miniatura (máx. 150×120 px) en la
  columna K de cada fila del registro. **La foto original no se guarda en ningún lado.**
- No borres ni renombres `Hoja 1` ni cambies el texto de la leyenda.
- Si dos personas guardan a la vez, la app reintenta sobre la versión más nueva y no se pierde
  ninguno. Si alguien tiene el Excel abierto, puede pedir que esperen un minuto.

## Seguridad (el link es público)

- La app solo **escribe**: nunca muestra, lista ni ofrece descargar nada de OneDrive.
- Ante cualquier error el usuario ve un mensaje genérico; el detalle técnico queda solo en el
  log de Streamlit Cloud ("Manage app" → logs).
- Los textos se recortan (p. ej. 100 caracteres el producto, 500 las observaciones) y lo que
  empieza con `=`, `+`, `-` o `@` se guarda como texto, nunca como fórmula.

## Pruebas

`pip install -r requirements.txt pytest` y luego `pytest`. Usan la plantilla de
`tests/plantillas/` y un OneDrive simulado; no necesitan internet ni credenciales.
