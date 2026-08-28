# -*- coding: utf-8 -*-
"""
Corrección de cierre a junio 2026 — tres ajustes, todos con fuente verificable.

1) INGRESOS COLOMBIA ene–may (escenario Real, COP)
   La base traía cifras redondeadas que sumaban 4.463.453.314. El estado oficial
   (PDF "Estados Financieros Junio 2026", pág. 2) da 4.472.342.747 de ingresos
   operacionales netos del semestre, y las cifras mensuales del Modelo suman
   EXACTAMENTE ese total (diferencia 0). Se adoptan las del Modelo.

2) COSTA RICA junio (bloque "Consolidación USD")
   Los 137.010 USD de ingresos NO corresponden a junio: entran después (indicación
   del CFO). Además el balance de la filial (Balance Costa Rica Junio 2026.xlsx) no
   registra ninguna cuenta de ingresos y arroja pérdida del ejercicio de 22.674,27.
   Se ponen los ingresos de junio en 0 y se recalcula la utilidad de Costa Rica y
   del Consolidado de ese mes.

3) ESTADO CONSOLIDADO H1 (bloque "Acumulado H1 2026")
   Se reemplaza por la versión definitiva del "Estado de Resultados Combinado por
   País — Acumulado a junio 2026", que incorpora la reclasificación de honorarios
   de Perú pagados por EE.UU. (10.557 USD) y ajusta gastos de ventas. La utilidad
   neta combinada pasa de (7.112) a (17.700) USD.

Uso:  python corregir_junio_2026.py            -> vista previa
      python corregir_junio_2026.py --escribir -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_COCO_2026.xlsx")
HOJA = "BD_Indicadores"
TABLA = "tbl_BD_Indicadores"
SEG_MENSUAL = "Consolidación USD"
SEG_H1 = "Acumulado H1 2026"

FUENTE_COL = "Ajustado al cierre oficial (PDF Estados Financieros Junio 2026); serie mensual del Modelo, suma = 4.472.342.747"
FUENTE_CR = "Ingresos de 137.010 USD reversados: no corresponden a junio (entran después). Balance de la filial no registra ingresos."
FUENTE_H1 = ("Estado de Resultados Combinado por País · acumulado a junio 2026 (versión definitiva, "
             "incluye reclasificación de honorarios de Perú pagados por EE.UU.).")

# --- 1) Ingresos Colombia mensuales (COP) ---
INGRESOS_COL = {
    "2026-01": 612247701, "2026-02": 744218137, "2026-03": 787041187,
    "2026-04": 750867662, "2026-05": 826514746, "2026-06": 751453314,
}

# --- 3) Estado combinado definitivo: codigo -> {pais: valor USD} (None = "-") ---
H1 = {
    "pyg_servicio_software":        {"EE.UU.":32997, "Perú":4679,  "Costa Rica":None,  "Colombia":1398926, "Consolidado":1436602},
    "pyg_devoluciones":             {"EE.UU.":None,  "Perú":None,  "Costa Rica":None,  "Colombia":-123034, "Consolidado":-123034},
    "pyg_ingresos_operacionales":   {"EE.UU.":32997, "Perú":4679,  "Costa Rica":None,  "Colombia":1275892, "Consolidado":1313567},
    "pyg_costo_ventas":             {"EE.UU.":None,  "Perú":None,  "Costa Rica":None,  "Colombia":605005,  "Consolidado":605005},
    "pyg_gasto_administracion":     {"EE.UU.":30736, "Perú":1600,  "Costa Rica":21792, "Colombia":196717,  "Consolidado":250845},
    "pyg_gasto_ventas":             {"EE.UU.":25246, "Perú":None,  "Costa Rica":None,  "Colombia":240203,  "Consolidado":265449},
    "pyg_gasto_proyectos":          {"EE.UU.":11500, "Perú":16797, "Costa Rica":None,  "Colombia":163589,  "Consolidado":191885},
    "pyg_reclasif_honorarios_peru": {"EE.UU.":10557, "Perú":None,  "Costa Rica":None,  "Colombia":None,    "Consolidado":10557},
    "pyg_ingresos_no_operacionales":{"EE.UU.":None,  "Perú":None,  "Costa Rica":None,  "Colombia":2876,    "Consolidado":2876},
    "pyg_gastos_financieros":       {"EE.UU.":15,    "Perú":265,   "Costa Rica":206,   "Colombia":9915,    "Consolidado":10401},
    "pyg_utilidad_neta":            {"EE.UU.":-45057,"Perú":-13983,"Costa Rica":-21998,"Colombia":63339,   "Consolidado":-17700},
}
COMPANIA = {"EE.UU.":"Coco LLC","Perú":"Coco Perú","Costa Rica":"Coco Costa Rica",
            "Colombia":"Coco Colombia","Consolidado":"Coco Consolidado"}


def col(enc, nombre):
    return enc.index(nombre) + 1


def main():
    escribir = "--escribir" in sys.argv
    wb = openpyxl.load_workbook(BASE)
    ws = wb[HOJA]
    enc = [c.value for c in ws[3]]
    C = {n: col(enc, n) for n in enc}
    cambios = []

    def val(r, n):
        return ws.cell(row=r, column=C[n]).value

    def set_(r, n, v):
        ws.cell(row=r, column=C[n], value=v)

    # ---------- 1) Ingresos Colombia ----------
    for r in range(4, ws.max_row + 1):
        if (val(r, "codigo_indicador") == "ingresos_totales" and val(r, "escenario") == "Real"
                and val(r, "pais") == "Colombia" and not val(r, "segmento")):
            per = str(val(r, "periodo"))
            if per in INGRESOS_COL and val(r, "valor") != INGRESOS_COL[per]:
                cambios.append(("1·Colombia", per, "ingresos_totales", val(r, "valor"), INGRESOS_COL[per], r))
                if escribir:
                    set_(r, "valor", INGRESOS_COL[per]); set_(r, "comentario", FUENTE_COL)

    # ---------- 2) Costa Rica junio ----------
    cr_gastos = 0.0
    for r in range(4, ws.max_row + 1):
        if (val(r, "segmento") == SEG_MENSUAL and val(r, "pais") == "Costa Rica"
                and str(val(r, "periodo")) == "2026-06"):
            cod = val(r, "codigo_indicador")
            if cod in ("pyg_gasto_administracion", "pyg_gasto_ventas", "pyg_gasto_proyectos", "pyg_gastos_financieros"):
                cr_gastos += (val(r, "valor") or 0)
    cr_util_nueva = -cr_gastos
    delta_cr = None
    for r in range(4, ws.max_row + 1):
        if (val(r, "segmento") == SEG_MENSUAL and val(r, "pais") == "Costa Rica"
                and str(val(r, "periodo")) == "2026-06"):
            cod = val(r, "codigo_indicador")
            if cod == "pyg_ingresos_operacionales" and (val(r, "valor") or 0) != 0:
                cambios.append(("2·CostaRica", "2026-06", cod, val(r, "valor"), 0, r))
                if escribir:
                    set_(r, "valor", 0); set_(r, "comentario", FUENTE_CR)
            elif cod == "pyg_utilidad_neta":
                delta_cr = cr_util_nueva - (val(r, "valor") or 0)
                cambios.append(("2·CostaRica", "2026-06", cod, val(r, "valor"), round(cr_util_nueva, 2), r))
                if escribir:
                    set_(r, "valor", round(cr_util_nueva, 2)); set_(r, "comentario", FUENTE_CR)

    # Consolidado de junio: se ajusta por el mismo delta
    for r in range(4, ws.max_row + 1):
        if (val(r, "segmento") == SEG_MENSUAL and val(r, "pais") == "Consolidado"
                and str(val(r, "periodo")) == "2026-06"):
            cod = val(r, "codigo_indicador")
            if cod == "pyg_ingresos_operacionales":
                nuevo = (val(r, "valor") or 0) - 137010
                cambios.append(("2·Consolidado", "2026-06", cod, val(r, "valor"), round(nuevo, 2), r))
                if escribir:
                    set_(r, "valor", round(nuevo, 2)); set_(r, "comentario", FUENTE_CR)
            elif cod == "pyg_utilidad_neta" and delta_cr is not None:
                nuevo = (val(r, "valor") or 0) + delta_cr
                cambios.append(("2·Consolidado", "2026-06", cod, val(r, "valor"), round(nuevo, 2), r))
                if escribir:
                    set_(r, "valor", round(nuevo, 2)); set_(r, "comentario", FUENTE_CR)

    # ---------- 3) Estado combinado H1 ----------
    existentes = {}
    for r in range(4, ws.max_row + 1):
        if val(r, "segmento") == SEG_H1:
            existentes[(val(r, "codigo_indicador"), val(r, "pais"))] = r
    fila_libre = ws.max_row + 1
    for cod, porpais in H1.items():
        for pais, v in porpais.items():
            if v is None:
                continue
            clave = (cod, pais)
            if clave in existentes:
                r = existentes[clave]
                if val(r, "valor") != v:
                    cambios.append(("3·H1", pais, cod, val(r, "valor"), v, r))
                    if escribir:
                        set_(r, "valor", v); set_(r, "comentario", FUENTE_H1)
            else:
                cambios.append(("3·H1 NUEVA", pais, cod, None, v, fila_libre))
                if escribir:
                    fila = {"periodo":"2026-06","pais":pais,"compania":COMPANIA[pais],"moneda":"USD",
                            "escenario":"Real","codigo_indicador":cod,"segmento":SEG_H1,"detalle":"",
                            "valor":v,"unidad":"USD","comentario":FUENTE_H1,"area":"Internacional",
                            "area_responsable":"Finanzas","periodicidad":"Semestral"}
                    for i, n in enumerate(enc, start=1):
                        ws.cell(row=fila_libre, column=i, value=fila.get(n, ""))
                fila_libre += 1

    # ---------- salida ----------
    print("=" * 92)
    print("  CORRECCIÓN DE CIERRE A JUNIO 2026 —", "ESCRITURA" if escribir else "VISTA PREVIA (no escribe)")
    print("=" * 92)
    print(f"  {'bloque':<16}{'quién':<14}{'indicador':<32}{'antes':>16}{'después':>16}")
    print("  " + "-" * 90)
    for b, q, cod, a, d, _ in cambios:
        fa = "—" if a is None else (f"{a:,.2f}" if isinstance(a, float) else f"{a:,}")
        fd = f"{d:,.2f}" if isinstance(d, float) else f"{d:,}"
        print(f"  {b:<16}{q:<14}{cod:<32}{fa:>16}{fd:>16}")
    print(f"\n  Total de ajustes: {len(cambios)}")

    if not escribir:
        print("\n  Nada se escribió. Para aplicar:  python corregir_junio_2026.py --escribir\n")
        return

    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", f"BD_COCO_2026_antes_correccion_{sello}.xlsx")
    shutil.copy2(BASE, copia)

    t = ws.tables[TABLA]
    ini, fin = t.ref.split(":")
    ci = "".join(c for c in ini if c.isalpha()); fi = "".join(c for c in ini if c.isdigit())
    cf = "".join(c for c in fin if c.isalpha())
    t.ref = f"{ci}{fi}:{cf}{max(fila_libre - 1, ws.max_row)}"
    wb.save(BASE)
    print(f"\n  Respaldo: {os.path.basename(copia)}")
    print(f"  Base actualizada. Tabla -> {t.ref}\n")


if __name__ == "__main__":
    main()
