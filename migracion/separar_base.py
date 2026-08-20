# -*- coding: utf-8 -*-
"""
SEPARAR LA BASE EN DOS — operativa y trazabilidad

Problema que resuelve
---------------------
El libro tiene 38 hojas, pero el tablero solo lee 10. La misma cifra vive en varias
capas (BD_Indicadores, BD_PYG_OFICIAL, Subsidiarias_PyG, vistas por área) y ya causó
desalineaciones: se actualiza una y se olvida otra. Separar elimina la posibilidad.

Cómo se decidió el corte (verificado, no supuesto)
-------------------------------------------------
1. Se leyó el parser del tablero: las únicas hojas que abre son las de OPERATIVA.
2. Se mapearon las fórmulas del libro: solo dos hojas son referenciadas por otras
   (Diccionario y Calc_LTV_Fin) -> ambas se quedan en la operativa.
3. Las hojas EEFF_* y los puentes de gastos no tienen NINGUNA fórmula cruzada:
   son datos ya calculados, por lo que moverlas no rompe nada.

Qué hace
--------
  BD_MAESTRA_COCO.xlsx        -> queda solo con las hojas operativas
  BD_TRAZABILIDAD_COCO.xlsx   -> recibe el resto (fuentes, puentes, vistas, controles)

Las hojas movidas se convierten a VALORES: son vistas y soportes, no cálculos vivos.
Así no quedan enlaces externos rotos apuntando al libro original.

Uso:  python separar_base.py             -> vista previa (no escribe)
      python separar_base.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
TRAZA = os.path.join(AQUI, "BD_TRAZABILIDAD_COCO.xlsx")

# --- Hojas que el tablero abre (verificado en parseWorkbook) ---
LEE_EL_TABLERO = [
    "Diccionario", "BD_Indicadores", "TRM", "TRM_Peru",
    "Analisis_Churn", "Analisis_Reconciliacion", "Calc_LTV_Fin",
    "Pipeline_Comercial", "CATALOGOS", "BD_GASTOS_RESUMEN_MENSUAL",
]
# --- Hojas que necesita la capa de escritura de Google (Apps Script) ---
USA_APPS_SCRIPT = ["LOG"]

# --- Cadena de producción de gastos ---
# Según DB_README son capas 3-5: el detalle homologado y el mapa alimentan el puente,
# y este produce BD_GASTOS_RESUMEN_MENSUAL, que es lo que el tablero lee. Es una cadena
# ACTIVA, no archivo: se queda completa en la base operativa para no separar la fábrica
# del producto en el cierre mensual.
CADENA_GASTOS = ["BD_GASTOS_HOMOLOGADOS", "MAP_DASHBOARD_GASTOS", "BD_GASTOS_DASHBOARD_BRIDGE"]

OPERATIVAS = LEE_EL_TABLERO + USA_APPS_SCRIPT + CADENA_GASTOS

MOTIVO = {
    "BD_PYG_OFICIAL": "soporte financiero; el tablero solo lo usa como etiqueta de origen",
    "PyG_Oficial": "repite cifras que ya están en BD_Indicadores",
    "PyG_Mensual_Col": "repite cifras que ya están en BD_Indicadores",
    "Subsidiarias_PyG": "resumen por subsidiaria; el tablero no lo abre",
    "Finanzas": "vista de solo lectura regenerada desde BD_Indicadores",
    "Contabilidad": "vista de solo lectura regenerada desde BD_Indicadores",
    "Ventas": "vista de solo lectura regenerada desde BD_Indicadores",
    "Marketing": "vista de solo lectura regenerada desde BD_Indicadores",
    "Customer_Success": "vista de solo lectura regenerada desde BD_Indicadores",
    "SaaS_Revenue": "vista de solo lectura regenerada desde BD_Indicadores",
    "BD_GASTOS_HOMOLOGADOS": "detalle previo al resumen de gastos",
    "BD_GASTOS_DASHBOARD_BRIDGE": "puente intermedio de gastos",
    "MAP_DASHBOARD_GASTOS": "mapa de homologación de gastos",
    "DB_README": "documentación de la base",
    "DB_CHECKS_ORGANIZACION": "controles de organización",
}


def main():
    escribir = "--escribir" in sys.argv
    if not os.path.exists(BASE):
        sys.exit(f"No encuentro la base: {BASE}")

    wb = openpyxl.load_workbook(BASE, data_only=True)   # data_only: guarda VALORES
    todas = list(wb.sheetnames)
    quedan = [h for h in todas if h in OPERATIVAS]
    mueven = [h for h in todas if h not in OPERATIVAS]
    faltan = [h for h in OPERATIVAS if h not in todas]

    print("=" * 84)
    print("  SEPARAR LA BASE —", "ESCRITURA" if escribir else "VISTA PREVIA (no escribe)")
    print("=" * 84)
    print(f"\n  BD_MAESTRA_COCO.xlsx  ->  se queda con {len(quedan)} hoja(s) OPERATIVAS:\n")
    for h in quedan:
        ws = wb[h]
        n = sum(1 for r in ws.iter_rows(values_only=True) if any(v is not None for v in r))
        marca = ("tablero" if h in LEE_EL_TABLERO
                 else "Apps Script" if h in USA_APPS_SCRIPT else "cadena de gastos")
        print(f"     {h:<32}{n:>7} filas   [{marca}]")
    if faltan:
        print(f"\n     AVISO: no existen en el libro: {', '.join(faltan)}")

    print(f"\n  BD_TRAZABILIDAD_COCO.xlsx  ->  recibe {len(mueven)} hoja(s):\n")
    for h in mueven:
        ws = wb[h]
        n = sum(1 for r in ws.iter_rows(values_only=True) if any(v is not None for v in r))
        motivo = MOTIVO.get(h, "fuente / soporte de auditoría")
        print(f"     {h:<32}{n:>7} filas   {motivo}")

    print("\n" + "-" * 84)
    print("  Nada se pierde: todo lo movido queda íntegro en el libro de trazabilidad.")
    if not escribir:
        print("  Para aplicar:  python separar_base.py --escribir\n")
        return

    # ---------- respaldo ----------
    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", f"BD_MAESTRA_COCO_antes_separar_{sello}.xlsx")
    shutil.copy2(BASE, copia)

    # ---------- libro de trazabilidad (valores) ----------
    wt = openpyxl.Workbook()
    wt.remove(wt.active)
    for h in mueven:
        origen = wb[h]
        destino = wt.create_sheet(h[:31])
        for fila in origen.iter_rows(values_only=True):
            destino.append(list(fila))
    wt.save(TRAZA)

    # ---------- base operativa: quitar las movidas (conserva fórmulas y tablas) ----------
    wb2 = openpyxl.load_workbook(BASE)          # sin data_only: preserva fórmulas
    for h in mueven:
        if h in wb2.sheetnames:
            del wb2[h]
    wb2.save(BASE)

    print(f"\n  Respaldo:      {os.path.basename(copia)}")
    print(f"  Trazabilidad:  {os.path.basename(TRAZA)}  ({len(mueven)} hojas)")
    print(f"  Base operativa: {len(quedan)} hojas")
    print("\n  Siguiente: corre Revisar_Base.bat y abre el tablero para confirmar.\n")


if __name__ == "__main__":
    main()
