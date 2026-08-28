# -*- coding: utf-8 -*-
"""
Agrega el presupuesto 2026 (combinado, original de enero, revisado de mayo)
a la tabla tbl_BD_Indicadores del libro maestro, como filas escenario='Presupuesto'.

- codigo_indicador='ingresos_totales' (mismo código que usa 'Real' -> comparables
  directamente sin tocar el dashboard).
- Combinado (segmento vacío = versión oficial aplicada): Original ene-abr, Revisado may-dic.
- Original enero / Revisado mayo (segmento con el nombre de versión): el plan completo
  de cada versión, para trazabilidad y comparación de versiones. Estas dos NO deben
  sumarse por defecto junto con el combinado -> el dashboard las excluye de las sumas
  generales (mismo mecanismo que usa para el bloque internacional).

Valores tomados literalmente de la tabla entregada (fuente: Finanzas COCO), en COP.
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

import openpyxl

RUTA = r"C:\Users\andre\OneDrive - CFOcus\01_CFOCUS_CLIENTES\COCO\1.0 Coco Digital\COCO_Dashboard_Cloud\migracion\BD_COCO_2026.xlsx"
HOJA = "BD_Indicadores"
TABLA = "tbl_BD_Indicadores"

MESES = ["2026-01","2026-02","2026-03","2026-04","2026-05","2026-06",
         "2026-07","2026-08","2026-09","2026-10","2026-11","2026-12"]

COMBINADO = [744928000,1417284000,1584956000,1436764000,1328900000,955991587,
             1047406944,1573626071,1665041429,1756456786,1847872143,1939287500]
ORIGINAL  = [744928000,1417284000,1584956000,1436764000,1612320000,1502236000,
             1552340000,1534356000,1697052000,1679816000,1692868000,1767984000]
REVISADO  = [612247701,744218137,787041187,750867662,1328900000,955991587,
             1047406944,1573626071,1665041429,1756456786,1847872143,1939287500]

VERSION_APLICADA = ["Original 4.5M"]*4 + ["Revisado 4.0M"]*8  # ene-abr / may-dic


def fila_base():
    return {"pais":"Colombia","compania":"Coco Colombia","moneda":"COP",
            "escenario":"Presupuesto","codigo_indicador":"ingresos_totales",
            "unidad":"COP","area":"Finanzas","area_responsable":"Finanzas",
            "periodicidad":"Mensual"}


def construir_filas():
    filas = []
    for i, per in enumerate(MESES):
        b = fila_base(); b["periodo"] = per
        # 1) Combinado = versión oficial aplicada ese mes (segmento vacío -> es lo que
        #    compara por defecto contra 'Real' en ppto_vs_real_ingresos)
        f = dict(b); f["segmento"] = ""; f["detalle"] = ""
        f["valor"] = COMBINADO[i]
        f["comentario"] = f"Presupuesto combinado 2026 · versión aplicada: {VERSION_APLICADA[i]} · TRM plan {'4.000' if i<4 else '3.690'}"
        filas.append(f)
        # 2) Versión Original (aprobada enero 2026, TRM plan 4.000, meta anual $4.5M)
        f = dict(b); f["segmento"] = "Original enero"; f["detalle"] = ""
        f["valor"] = ORIGINAL[i]
        f["comentario"] = "Presupuesto original aprobado en enero 2026 · TRM plan 4.000 · meta anual USD 4.5M"
        filas.append(f)
        # 3) Versión Revisada (aprobada mayo 2026, TRM plan 3.690, meta anual $4.0M)
        f = dict(b); f["segmento"] = "Revisado mayo"; f["detalle"] = ""
        f["valor"] = REVISADO[i]
        f["comentario"] = "Presupuesto revisado en mayo 2026 · TRM plan 3.690 · meta anual USD 4.0M"
        filas.append(f)
    return filas


def main():
    wb = openpyxl.load_workbook(RUTA)  # sin data_only: preserva fórmulas del resto del libro
    ws = wb[HOJA]
    tabla = ws.tables[TABLA]

    encabezados = [c.value for c in ws[3]]
    print("Encabezados:", encabezados)

    fila_inicio = ws.max_row + 1
    filas = construir_filas()
    print(f"Filas a insertar: {len(filas)} (desde la fila {fila_inicio})")

    for i, f in enumerate(filas):
        r = fila_inicio + i
        for col_idx, nombre_col in enumerate(encabezados, start=1):
            ws.cell(row=r, column=col_idx, value=f.get(nombre_col, ""))

    fila_fin = fila_inicio + len(filas) - 1
    col_letra_fin = tabla.ref.split(":")[1][0] if not tabla.ref.split(":")[1][1].isdigit() else tabla.ref.split(":")[1][:2]
    # reconstruir ref preservando la columna final (N) y la fila inicial (3)
    col_ini, _ = tabla.ref.split(":")
    col_letra_ini = ''.join(ch for ch in col_ini if ch.isalpha())
    col_letra_fin_real = ''.join(ch for ch in tabla.ref.split(":")[1] if ch.isalpha())
    fila_ini_tabla = ''.join(ch for ch in col_ini if ch.isdigit())
    nuevo_ref = f"{col_letra_ini}{fila_ini_tabla}:{col_letra_fin_real}{fila_fin}"
    print("Ref de tabla:", tabla.ref, "->", nuevo_ref)
    tabla.ref = nuevo_ref

    wb.save(RUTA)
    print(f"\nListo. {len(filas)} filas de presupuesto agregadas a {TABLA}.")
    print(f"Nueva última fila con datos: {fila_fin}")


if __name__ == "__main__":
    main()
