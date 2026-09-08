# -*- coding: utf-8 -*-
"""
Version en la nube de exportar_stock.py, pensada para correr dentro de
GitHub Actions (no en tu computador). Descarga el Excel directamente desde
Google Drive usando el ID del archivo, sincroniza fotos y documentos desde
una carpeta de Drive (via cuenta de servicio), y genera docs/data.json +
docs/images + docs/docs.

Variables de entorno requeridas:
  DRIVE_FILE_ID              - ID del archivo Excel (Settings > Variables)
  DRIVE_MEDIA_FOLDER_ID       - ID de la carpeta de Drive con fotos/documentos (Variables)
  DRIVE_NEWS_FOLDER_ID        - ID de la carpeta de Drive con noticias (Variables)
  GOOGLE_SERVICE_ACCOUNT_JSON - contenido completo del JSON de la cuenta de servicio (Secrets)
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession

try:
    import docx
except ImportError:
    docx = None

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

COL_CODIGO = "B"
COL_TIPO = "C"
COL_CATEGORIA = "D"
COL_NOMBRE = "F"
COL_PRESENTACION = "G"
COL_DESCRIPCION = "H"
COL_SALDO = "J"
COL_FECHA_INGRESO = "Z"

FILA_INICIO = 5
DIAS_CONSIDERADO_NUEVO = 28

RUTA_TEMPORAL = "planilla_descargada.xlsm"
RUTA_JSON = Path("docs/data.json")
CARPETA_IMAGENES = Path("docs/images")
CARPETA_DOCS = Path("docs/docs")
CARPETA_NOTICIAS_IMG = Path("docs/noticias/images")
RUTA_NOTICIAS_JSON = Path("docs/noticias.json")
MAX_NOTICIAS = 20

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".webp"}
EXTENSIONES_DOC = {".pdf", ".txt", ".html", ".htm", ".doc", ".docx"}
EXTENSIONES_DOC_NOTICIA = {".txt", ".html", ".htm", ".docx", ".doc"}


def descargar_de_google_drive(file_id: str, destino: str):
    """
    Descarga un archivo publico ('cualquiera con el enlace') desde Google Drive,
    manejando la pantalla de confirmacion que Google muestra para archivos
    con macros o de mayor tamano.
    """
    url = "https://docs.google.com/uc?export=download"
    session = requests.Session()

    respuesta = session.get(url, params={"id": file_id}, stream=True)

    token = None
    for clave, valor in respuesta.cookies.items():
        if clave.startswith("download_warning"):
            token = valor
            break

    if token:
        respuesta = session.get(
            url, params={"id": file_id, "confirm": token}, stream=True
        )

    if respuesta.status_code != 200:
        print(f"ERROR: Google Drive respondio con codigo {respuesta.status_code}")
        sys.exit(1)

    with open(destino, "wb") as f:
        for chunk in respuesta.iter_content(32768):
            if chunk:
                f.write(chunk)


def valor_texto(celda):
    v = celda.value
    if isinstance(v, str):
        return v.strip()
    return v or ""


def leer_productos_con_stock(ruta_excel: str):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws = wb["Hoja1"]

    productos = []
    filas_vacias = 0
    total_filas_con_codigo = 0

    for fila in range(FILA_INICIO, ws.max_row + 1):
        codigo = ws[f"{COL_CODIGO}{fila}"].value
        if codigo is None or str(codigo).strip() == "":
            filas_vacias += 1
            continue

        total_filas_con_codigo += 1
        saldo = ws[f"{COL_SALDO}{fila}"].value
        try:
            saldo = float(saldo) if saldo is not None else 0
        except (TypeError, ValueError):
            saldo = 0

        if saldo > 0:
            fecha_ingreso_val = ws[f"{COL_FECHA_INGRESO}{fila}"].value
            es_nuevo = False
            if isinstance(fecha_ingreso_val, datetime):
                dias = (datetime.now() - fecha_ingreso_val).days
                es_nuevo = 0 <= dias <= DIAS_CONSIDERADO_NUEVO

            productos.append({
                "codigo": str(codigo).strip(),
                "nombre": valor_texto(ws[f"{COL_NOMBRE}{fila}"]),
                "descripcion": valor_texto(ws[f"{COL_DESCRIPCION}{fila}"]),
                "saldo": int(saldo) if saldo == int(saldo) else saldo,
                "nuevo": es_nuevo,
                "tipo": valor_texto(ws[f"{COL_TIPO}{fila}"]),
                "categoria": valor_texto(ws[f"{COL_CATEGORIA}{fila}"]),
                "presentacion": valor_texto(ws[f"{COL_PRESENTACION}{fila}"]),
            })

    wb.close()
    print(f"  Filas revisadas: {FILA_INICIO} a {ws.max_row}")
    print(f"  Filas con codigo encontradas: {total_filas_con_codigo}")
    print(f"  Filas vacias encontradas (toleradas, no cortan la lectura): {filas_vacias}")
    print(f"  Productos con stock > 0: {len(productos)}")
    return productos


def listar_archivos_carpeta_drive(folder_id: str, session: AuthorizedSession, con_fecha=False):
    """
    Lista los archivos dentro de una carpeta de Google Drive compartida con
    la cuenta de servicio, usando la API de Drive autenticada.
    """
    archivos = []
    url = "https://www.googleapis.com/drive/v3/files"
    page_token = None
    campos = "id, name" + (", createdTime" if con_fecha else "")

    while True:
        params = {
            "q": f"'{folder_id}' in parents and trashed = false",
            "fields": f"nextPageToken, files({campos})",
            "pageSize": 1000,
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        }
        if page_token:
            params["pageToken"] = page_token

        r = session.get(url, params=params)
        if r.status_code != 200:
            print(f"ERROR listando carpeta de Drive: {r.status_code} - {r.text}")
            sys.exit(1)

        data = r.json()
        archivos.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break

    return archivos


def descargar_archivo_drive_api(file_id: str, session: AuthorizedSession, destino: Path):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    r = session.get(url, params={"alt": "media"}, stream=True)
    if r.status_code != 200:
        print(f"  ADVERTENCIA: no se pudo descargar {destino.name} ({r.status_code})")
        return False
    with open(destino, "wb") as f:
        for chunk in r.iter_content(32768):
            if chunk:
                f.write(chunk)
    return True


RE_IMG = re.compile(r"^(.*?)img(\d*)$")
RE_DOC = re.compile(r"^(.*?)doc$")


def extraer_codigo_y_tipo(nombre_archivo: str):
    """
    'CODIGOimg.jpg'  -> (CODIGO, 'img', 0)   -- foto principal
    'CODIGOimg2.jpg' -> (CODIGO, 'img', 2)   -- foto adicional (2, 3, 4...)
    'CODIGOdoc.pdf'  -> (CODIGO, 'doc', 0)
    Devuelve (None, None, None, None) si no calza con el patron esperado.
    """
    base, ext = os.path.splitext(nombre_archivo)
    ext = ext.lower()

    m = RE_IMG.match(base)
    if m and ext in EXTENSIONES_IMAGEN:
        orden = int(m.group(2)) if m.group(2) else 1
        return m.group(1), "img", ext, orden

    m = RE_DOC.match(base)
    if m and ext in EXTENSIONES_DOC:
        return m.group(1), "doc", ext, 0

    return None, None, None, None


def sincronizar_media(folder_id: str, session: AuthorizedSession):
    """
    Descarga fotos y documentos desde la carpeta de Drive.
    Devuelve:
      imagenes: {codigo: [ruta1, ruta2, ...]} en orden (principal primero)
      documentos: {codigo: {"ruta": ..., "titulo":..., "texto": ...}}
    """
    imagenes_raw = {}   # codigo -> lista de (orden, ruta)
    documentos = {}

    archivos = listar_archivos_carpeta_drive(folder_id, session)
    print(f"Archivos encontrados en la carpeta de Drive: {len(archivos)}")

    for archivo in archivos:
        codigo, tipo, ext, orden = extraer_codigo_y_tipo(archivo["name"])
        if codigo is None:
            print(f"  Omitido (no calza con el patron CODIGOimg/CODIGOdoc): {archivo['name']}")
            continue

        if tipo == "img":
            sufijo = "" if orden <= 1 else str(orden)
            destino = CARPETA_IMAGENES / f"{codigo}img{sufijo}{ext}"
            if descargar_archivo_drive_api(archivo["id"], session, destino):
                imagenes_raw.setdefault(codigo, []).append((orden, f"images/{codigo}img{sufijo}{ext}"))
        elif tipo == "doc":
            destino = CARPETA_DOCS / f"{codigo}doc{ext}"
            if descargar_archivo_drive_api(archivo["id"], session, destino):
                info_doc = {"ruta": f"docs/{codigo}doc{ext}", "titulo": "", "texto": ""}
                if ext in (".txt", ".html", ".htm", ".docx", ".doc"):
                    titulo, cuerpo = extraer_texto_documento(destino, ext)
                    info_doc["titulo"] = titulo
                    info_doc["texto"] = cuerpo
                documentos[codigo] = info_doc

    imagenes = {}
    for codigo, lista in imagenes_raw.items():
        lista.sort(key=lambda t: t[0])
        imagenes[codigo] = [ruta for _, ruta in lista]

    return imagenes, documentos


def extraer_texto_documento(ruta: Path, ext: str):
    """
    Devuelve (titulo, cuerpo) leyendo un .txt, .html, .docx o .doc.
    La primera linea/parrafo se usa como titulo.
    """
    lineas = []

    if ext == ".docx":
        if docx is None:
            print("  ADVERTENCIA: falta la libreria python-docx, no se puede leer .docx")
            return "", ""
        d = docx.Document(str(ruta))
        lineas = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    elif ext == ".doc":
        import subprocess
        try:
            resultado = subprocess.run(
                ["antiword", str(ruta)], capture_output=True, text=True, timeout=30
            )
            if resultado.returncode != 0:
                print(f"  ADVERTENCIA: antiword no pudo leer {ruta.name}: {resultado.stderr.strip()}")
                return "", ""
            lineas = [l.strip() for l in resultado.stdout.splitlines() if l.strip()]
        except FileNotFoundError:
            print("  ADVERTENCIA: falta instalar 'antiword' para leer archivos .doc")
            return "", ""
    elif ext in (".html", ".htm"):
        import re
        contenido = ruta.read_text(encoding="utf-8", errors="ignore")
        texto = re.sub("<[^>]+>", "\n", contenido)
        lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    else:  # .txt
        contenido = ruta.read_text(encoding="utf-8", errors="ignore")
        lineas = [l.strip() for l in contenido.splitlines() if l.strip()]

    if not lineas:
        return "", ""

    titulo = lineas[0]
    cuerpo = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
    return titulo, cuerpo


def sincronizar_noticias(folder_id: str, session: AuthorizedSession):
    """
    Descarga noticias (foto + texto) desde una carpeta de Drive.
    Cada noticia son 2 archivos: IDimg.<ext> e IDdoc.<ext>.
    Devuelve una lista de noticias ordenadas de mas nueva a mas vieja.
    """
    CARPETA_NOTICIAS_IMG.mkdir(parents=True, exist_ok=True)
    (CARPETA_NOTICIAS_IMG / ".gitkeep").touch(exist_ok=True)

    archivos = listar_archivos_carpeta_drive(folder_id, session, con_fecha=True)
    print(f"Archivos encontrados en la carpeta de noticias: {len(archivos)}")

    por_id = {}
    for archivo in archivos:
        base, ext = os.path.splitext(archivo["name"])
        ext = ext.lower()
        if base.endswith("img") and ext in EXTENSIONES_IMAGEN:
            noticia_id = base[:-3]
            por_id.setdefault(noticia_id, {})["img"] = (archivo, ext)
        elif base.endswith("doc") and ext in EXTENSIONES_DOC_NOTICIA:
            noticia_id = base[:-3]
            por_id.setdefault(noticia_id, {})["doc"] = (archivo, ext)

    noticias = []
    for noticia_id, partes in por_id.items():
        if "doc" not in partes:
            print(f"  Omitido '{noticia_id}': falta el archivo de texto (doc)")
            continue

        archivo_doc, ext_doc = partes["doc"]
        ruta_doc_temp = Path(f"_tmp_{noticia_id}{ext_doc}")
        if not descargar_archivo_drive_api(archivo_doc["id"], session, ruta_doc_temp):
            continue
        titulo, cuerpo = extraer_texto_documento(ruta_doc_temp, ext_doc)
        ruta_doc_temp.unlink(missing_ok=True)

        imagen_rel = None
        if "img" in partes:
            archivo_img, ext_img = partes["img"]
            destino_img = CARPETA_NOTICIAS_IMG / f"{noticia_id}{ext_img}"
            if descargar_archivo_drive_api(archivo_img["id"], session, destino_img):
                imagen_rel = f"noticias/images/{noticia_id}{ext_img}"

        noticias.append({
            "id": noticia_id,
            "titulo": titulo or noticia_id,
            "texto": cuerpo,
            "imagen": imagen_rel,
            "fecha": archivo_doc.get("createdTime", ""),
        })

    noticias.sort(key=lambda n: n["fecha"], reverse=True)
    return noticias[:MAX_NOTICIAS]


def main():
    file_id = os.environ.get("DRIVE_FILE_ID")
    if not file_id:
        print("ERROR: falta la variable de entorno DRIVE_FILE_ID")
        sys.exit(1)

    print("Descargando planilla desde Google Drive...")
    descargar_de_google_drive(file_id, RUTA_TEMPORAL)

    print("Leyendo productos con stock...")
    productos = leer_productos_con_stock(RUTA_TEMPORAL)
    print(f"Se encontraron {len(productos)} productos con stock.")

    media_folder_id = os.environ.get("DRIVE_MEDIA_FOLDER_ID")
    news_folder_id = os.environ.get("DRIVE_NEWS_FOLDER_ID")
    service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    imagenes, documentos = {}, {}

    CARPETA_IMAGENES.mkdir(parents=True, exist_ok=True)
    CARPETA_DOCS.mkdir(parents=True, exist_ok=True)
    (CARPETA_IMAGENES / ".gitkeep").touch(exist_ok=True)
    (CARPETA_DOCS / ".gitkeep").touch(exist_ok=True)

    session = None
    if service_account_json and (media_folder_id or news_folder_id):
        info = json.loads(service_account_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        session = AuthorizedSession(creds)

    if media_folder_id and session:
        print("Sincronizando fotos y documentos desde Drive...")
        imagenes, documentos = sincronizar_media(media_folder_id, session)
        print(f"Fotos sincronizadas: {len(imagenes)} | Documentos sincronizados: {len(documentos)}")
    else:
        print("Sin DRIVE_MEDIA_FOLDER_ID / GOOGLE_SERVICE_ACCOUNT_JSON: se omite sincronizacion de fotos y documentos.")

    if news_folder_id and session:
        print("Sincronizando noticias desde Drive...")
        noticias = sincronizar_noticias(news_folder_id, session)
        print(f"Noticias sincronizadas: {len(noticias)}")
        with open(RUTA_NOTICIAS_JSON, "w", encoding="utf-8") as f:
            json.dump({"noticias": noticias}, f, ensure_ascii=False, indent=2)
    else:
        print("Sin DRIVE_NEWS_FOLDER_ID: se omite sincronizacion de noticias.")

    for p in productos:
        cod = p["codigo"]
        if cod in imagenes:
            p["imagenes"] = imagenes[cod]
            p["imagen"] = imagenes[cod][0]   # portada, para la tarjeta del listado
        if cod in documentos:
            doc = documentos[cod]
            p["documento"] = doc["ruta"]
            if doc["texto"]:
                p["informacion"] = doc["texto"]
            if doc["titulo"]:
                p["informacion_titulo"] = doc["titulo"]

    RUTA_JSON.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "generado": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "productos": productos,
    }
    with open(RUTA_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"Archivo generado: {RUTA_JSON}")
    os.remove(RUTA_TEMPORAL)


if __name__ == "__main__":
    main()
