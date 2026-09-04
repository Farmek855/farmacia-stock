# -*- coding: utf-8 -*-
"""
Version en la nube de exportar_stock.py, pensada para correr dentro de
GitHub Actions (no en tu computador). Descarga el Excel directamente desde
Google Drive usando el ID del archivo, y genera docs/data.json.

Requiere la variable de entorno DRIVE_FILE_ID (se configura en GitHub,
Settings > Secrets and variables > Actions > Variables).
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import requests

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
    fila = FILA_INICIO
    while True:
        codigo = ws[f"{COL_CODIGO}{fila}"].value
        if codigo is None or str(codigo).strip() == "":
            siguiente = ws[f"{COL_CODIGO}{fila + 1}"].value
            if siguiente is None or str(siguiente).strip() == "":
                break
            fila += 1
            continue

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
        fila += 1

    wb.close()
    return productos


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
