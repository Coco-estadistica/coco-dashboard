# -*- coding: utf-8 -*-
"""
Carga julio 2026 de "gasto_area" (el agregado de 5 categorias -- Financieros,
TI, Proveedores, Comercial, Administracion -- que alimenta el grafico
"Costos y gastos por area" y evita que julio caiga entero en "Sin desglose
contable"). Es la SUMA de lo que ya se cargo por separado en
cargar_gastos_colombia_julio.py, mas Financieros que sale del P&G oficial:

  Administracion = gasto_admin (6 de 7 categorias, julio)      =  95,247,095
  Comercial      = gasto_comercial (6 de 6, julio)              = 120,916,831
  Proveedores    = costo_proveedor (2 de 5, julio)               = 114,769,552
  TI             = costo_operacion (4 de 5) + gasto_desarrollo   = 264,158,213
  Financieros    = pyg_gastos_financieros oficial julio Colombia =   5,353,621

Cada una queda por debajo del total real de su categoria (por los huecos ya
documentados: Diversos admin, AWS/IVA/Otros proveedores, Viajes operaciones)
-- la franja "Sin desglose contable" en el grafico sigue cubriendo esa
diferencia contra el P&G oficial, como ya hace para los demas meses.

Uso:  python cargar_gasto_area_julio.py             -> vista previa
      python cargar_gasto_area_julio.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
PERIODO = "2026-07"

COMENTARIO = ("Julio 2026. Suma de gasto_admin+gasto_comercial+costo_proveedor+"
              "costo_operacion+gasto_desarrollo (parcial, ver esos codigos) + "
              "pyg_gastos_financieros oficial. Alimenta 'Costos y gastos por area'; "
              "la franja gris del grafico cubre lo que aun falta vs. el P&G oficial.")

DATOS = [
    ("Administración", 95247095),
    ("Comercial", 120916831),
    ("Proveedores", 114769552),
    ("TI", 264158213),
    ("Financieros", 5353621),
]


def main():
    escribir = "--escribir" in sys.argv
    wb = openpyxl.load_workbook(BASE)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) + 1 for c in H if c}

    existentes = {}
    for f in range(4, ws.max_row + 1):
        cod = ws.cell(f, ix["codigo_indicador"]).value
        det = ws.cell(f, ix["detalle"]).value
        per = ws.cell(f, ix["periodo"]).value
        pais = ws.cell(f, ix["pais"]).value
        if cod == "gasto_area" and pais == "Colombia":
            existentes[(det, str(per))] = f

    siguiente = ws.max_row + 1
    reporte = []
    total = 0

    for det, valor in DATOS:
        total += valor
        clave = (det, PERIODO)
        if clave in existentes:
            f = existentes[clave]
            antes = ws.cell(f, ix["valor"]).value
            reporte.append("  UPDATE gasto_area  %-14s %s -> %s" % (det, antes, valor))
        else:
            f = siguiente
            siguiente += 1
            ws.cell(f, ix["periodo"]).value = PERIODO
            ws.cell(f, ix["pais"]).value = "Colombia"
            ws.cell(f, ix["compania"]).value = "Coco Colombia"
            ws.cell(f, ix["moneda"]).value = "COP"
            ws.cell(f, ix["escenario"]).value = "Real"
            ws.cell(f, ix["periodicidad"]).value = "Mensual"
            reporte.append("  INSERT gasto_area  %-14s -> %s" % (det, valor))
        ws.cell(f, ix["codigo_indicador"]).value = "gasto_area"
        ws.cell(f, ix["detalle"]).value = det
        ws.cell(f, ix["valor"]).value = valor
        ws.cell(f, ix["unidad"]).value = "COP"
        ws.cell(f, ix["comentario"]).value = COMENTARIO

    print("=" * 90)
    print("  CARGA gasto_area - JULIO 2026 (Colombia)")
    print("=" * 90)
    print()
    for l in reporte:
        print(l)
    print("\n  Total gasto_area julio (parcial):", format(total, ","))

    if not escribir:
        print("\n  Vista previa. Para aplicar agrega --escribir\n")
        return

    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", "BD_MAESTRA_COCO_antes_gasto_area_julio_" + sello + ".xlsx")
    shutil.copy2(BASE, copia)
    wb.save(BASE)
    print("\n  Respaldo:", os.path.basename(copia))
    print("  Base actualizada.\n")


if __name__ == "__main__":
    main()
