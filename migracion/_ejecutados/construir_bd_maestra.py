# -*- coding: utf-8 -*-
"""
Construye BD_MAESTRA_COCO.xlsx a partir de la base limpia
(BD_Indicadores_Dashboard_COCO (5).xlsx) y las 8 plantillas de carga.

Salida: un único libro con las pestañas que el Google Sheet maestro necesita:
  - Modelo (las lee parseWorkbook del dashboard):
      Diccionario, BD_Indicadores, TRM, Analisis_Reconciliacion,
      Analisis_Churn, Calc_LTV_Fin
  - Nueva:  Pipeline_Comercial (estructura de la plantilla Comercial_Ventas)
  - Control: CATALOGOS (área→códigos y listas de validación), LOG (vacía)

Uso:
    python construir_bd_maestra.py
Luego: subir BD_MAESTRA_COCO.xlsx a Google Drive → abrir con Google Sheets
       → Archivo → Guardar como hoja de cálculo de Google → renombrar
       "BD_MAESTRA_COCO". Ese Sheet es la fuente de verdad.
"""

# ---------------------------------------------------------------------------
# SEGURO. Este script ya se ejecuto y no forma parte del cierre mensual. Escribe
# sobre BD_MAESTRA_COCO.xlsx SIN pedir --escribir, asi que un doble clic bastaba
# para volver a aplicar un cambio que ya esta aplicado. Ver _ejecutados/LEEME.md.
import sys as _sys
if "--si-se-lo-que-hago" not in _sys.argv:
    print(__doc__ or "")
    print("=" * 74)
    print("  DETENIDO. Este script ya se ejecuto: correrlo otra vez duplica su efecto.")
    print("  Lo que hace hoy lo hace cargar_mes.py, con controles.")
    print("  Si de verdad hace falta: copia la base primero y agrega")
    print("  --si-se-lo-que-hago")
    print("=" * 74)
    raise SystemExit(1)
# ---------------------------------------------------------------------------

import os
import sys
import openpyxl
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
FUENTE = r"C:\Users\andre\Downloads\BD_Indicadores_Dashboard_COCO (5).xlsx"
PLANTILLAS = os.path.join(AQUI, "..", "plantillas")
SALIDA = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")

# Hojas del modelo que se copian tal cual desde la base limpia
HOJAS_MODELO = [
    "Diccionario", "BD_Indicadores", "TRM",
    "Analisis_Reconciliacion", "Analisis_Churn", "Calc_LTV_Fin",
]

# Mapeo área → plantilla (los códigos permitidos salen de la hoja Listas col A)
AREAS_PLANTILLA = {
    "Finanzas":         "Plantilla_Carga_Finanzas.xlsx",
    "Ingresos_MRR":     "Plantilla_Carga_Ingresos_MRR.xlsx",
    "Comercial_Ventas": "Plantilla_Carga_Comercial_Ventas.xlsx",
    "Marketing":        "Plantilla_Carga_Marketing.xlsx",
    "Customer_Success": "Plantilla_Carga_Customer_Success.xlsx",
    "Cartera_Liquidez": "Plantilla_Carga_Cartera_Liquidez.xlsx",
    "Costos_Gastos":    "Plantilla_Carga_Costos_Gastos.xlsx",
    "Internacional":    "Plantilla_Carga_Internacional.xlsx",
}

PIPELINE_HEADER = ["cliente", "comercial", "tipo", "prob", "mes",
                   "mrr_cop", "impl_cop", "citas", "valor_cop", "ponderado_cop"]

LOG_HEADER = ["timestamp", "usuario", "accion", "area", "hoja",
              "filas_afectadas", "detalle"]


def copiar_hoja(ws_src, ws_dst):
    """Copia valores celda a celda (sin estilos; Google Sheets no los necesita)."""
    for row in ws_src.iter_rows():
        for cell in row:
            if cell.value is not None:
                ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)


def leer_codigos_area(ruta_plantilla):
    wb = openpyxl.load_workbook(ruta_plantilla, data_only=True)
    ws = wb["Listas"]
    codigos = []
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
        v = row[0]
        if v is not None and str(v).strip():
            codigos.append(str(v).strip())
    return codigos


def main():
    if not os.path.exists(FUENTE):
        sys.exit(f"ERROR: no se encontró la base limpia: {FUENTE}")

    src = openpyxl.load_workbook(FUENTE, data_only=True)
    out = openpyxl.Workbook()
    out.remove(out.active)

    # --- 1. Hojas del modelo, tal cual ---
    for nombre in HOJAS_MODELO:
        if nombre not in src.sheetnames:
            sys.exit(f"ERROR: la base limpia no tiene la hoja '{nombre}'")
        copiar_hoja(src[nombre], out.create_sheet(nombre))
        print(f"  [ok] {nombre}: {src[nombre].max_row} filas copiadas")

    # --- 2. Pipeline_Comercial (nueva, con encabezado; datos si la plantilla trae) ---
    ws_pipe = out.create_sheet("Pipeline_Comercial")
    ws_pipe.append(PIPELINE_HEADER)
    ruta_com = os.path.join(PLANTILLAS, AREAS_PLANTILLA["Comercial_Ventas"])
    wb_com = openpyxl.load_workbook(ruta_com, data_only=True)
    n_pipe = 0
    if "Pipeline_Comercial" in wb_com.sheetnames:
        for row in wb_com["Pipeline_Comercial"].iter_rows(min_row=2, values_only=True):
            # fila real = tiene cliente; las filas de relleno solo traen 0/0 en fórmulas
            if row[0] is not None and str(row[0]).strip():
                ws_pipe.append(list(row[:len(PIPELINE_HEADER)]))
                n_pipe += 1
    print(f"  [ok] Pipeline_Comercial: encabezado + {n_pipe} filas")

    # --- 3. CATALOGOS: área→códigos + listas de validación ---
    ws_cat = out.create_sheet("CATALOGOS")
    ws_cat.append(["area", "codigo_indicador"])
    total = 0
    for area, plantilla in AREAS_PLANTILLA.items():
        ruta = os.path.join(PLANTILLAS, plantilla)
        if not os.path.exists(ruta):
            sys.exit(f"ERROR: falta la plantilla {ruta}")
        for cod in leer_codigos_area(ruta):
            ws_cat.append([area, cod])
            total += 1
    # Listas de validación en columnas D-G (misma fuente que usan las plantillas)
    ws_cat["D1"], ws_cat["E1"], ws_cat["F1"], ws_cat["G1"] = (
        "paises", "escenarios", "monedas", "companias")
    listas = {
        "D": ["Colombia", "Costa Rica", "Guatemala", "Panamá", "Regional", "Otros"],
        "E": ["Real", "Presupuesto"],
        "F": ["COP", "USD"],
        "G": ["Coco Colombia", "COCO Tecnologías"],
    }
    for col, valores in listas.items():
        for i, v in enumerate(valores, start=2):
            ws_cat[f"{col}{i}"] = v
    print(f"  [ok] CATALOGOS: {total} pares área→código + listas de validación")

    # --- 4. LOG (vacía, solo encabezado) ---
    ws_log = out.create_sheet("LOG")
    ws_log.append(LOG_HEADER)
    print("  [ok] LOG: encabezado creado")

    out.save(SALIDA)
    print(f"\nListo → {SALIDA}")
    print("Pestañas:", ", ".join(out.sheetnames))


if __name__ == "__main__":
    main()
