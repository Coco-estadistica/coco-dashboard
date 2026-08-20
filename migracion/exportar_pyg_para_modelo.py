# -*- coding: utf-8 -*-
"""
Exporta el P&G mensual de EE.UU./Peru/Costa Rica/Colombia/Consolidado (ene-jul 2026)
en el mismo layout del modelo financiero (bloques por mes, columnas por entidad),
para pegar directo en "Consolidado" del modelo.

Fuente: BD_Indicadores, segmento "Consolidacion USD", ya recargado desde los
auxiliares contables el 13-ago-2026.

Colombia y Consolidado NO tienen ene-mar (ver nota en la hoja): esas celdas
quedan como "s/d", nunca en blanco silencioso ni en cero.

Uso: python exportar_pyg_para_modelo.py
"""
import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
SALIDA = os.path.join(AQUI, "..", "P&G_Filiales_Colombia_Consolidado_ene-jul2026.xlsx")

MESES = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07"]
NOMBRE_MES = {"2026-01": "ENERO", "2026-02": "FEBRERO", "2026-03": "MARZO", "2026-04": "ABRIL",
              "2026-05": "MAYO", "2026-06": "JUNIO", "2026-07": "JULIO"}
ENTIDADES = ["EE.UU.", "Perú", "Costa Rica", "Colombia", "Consolidado"]
ETQ_ENT = {"EE.UU.": "USA", "Perú": "PERU", "Costa Rica": "COSTA RICA",
           "Colombia": "COLOMBIA", "Consolidado": "CONSOLIDADO"}
MONEDA_ORIGEN = {"EE.UU.": "USD", "Perú": "USD/PEN", "Costa Rica": "USD",
                 "Colombia": "COP/USD", "Consolidado": "TOTAL USD"}
LINEAS = [
    ("pyg_ingresos_operacionales", "Ingresos Ordinarios", "ing"),
    ("pyg_ingresos_no_operacionales", "Otros Ingresos", "ing"),
    (None, "", None),
    ("pyg_costo_ventas", "Costo del servicio", "gas"),
    (None, "", None),
    ("pyg_gasto_administracion", "Administración", "gas"),
    ("pyg_gasto_ventas", "Ventas", "gas"),
    ("pyg_gasto_proyectos", "Operación / Proyectos", "gas"),
    (None, "", None),
    ("pyg_gastos_financieros", "Gastos financieros", "gas"),
    (None, "", None),
    ("pyg_utilidad_neta", "UTILIDAD / (PÉRDIDA) NETA", "res"),
]

AZUL = "1E3A5F"
AZUL_H = "17324C"
GRIS = "F2F4F7"
BLANCO = "FFFFFF"
ROJO = "C42828"


def cargar_valores():
    wb = openpyxl.load_workbook(BASE, data_only=True)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) for c in H if c}
    filas = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not any(r):
            continue
        d = {c: r[ix[c]] for c in ix}
        if d.get("codigo_indicador"):
            filas.append(d)

    wt = ws  # TRM vive en otra hoja
    trm_ws = wb["TRM"]
    trm = {}
    for r in trm_ws.iter_rows(values_only=True):
        if r[0] and str(r[0]) in MESES:
            trm[str(r[0])] = r[1]

    def valor(pais, per, cod):
        for d in filas:
            if (d.get("pais") == pais and str(d["periodo"]) == per and d["codigo_indicador"] == cod
                    and (d.get("segmento") or "") == "Consolidación USD"):
                return d.get("valor")
        return None

    return valor, trm


def construir():
    valor, trm = cargar_valores()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado"
    ws.sheet_view.showGridLines = False

    f_titulo = Font(bold=True, size=13, color=BLANCO)
    f_sub = Font(italic=True, size=9, color="D2DCEB")
    f_head = Font(bold=True, size=10, color=BLANCO)
    f_ent = Font(bold=True, size=9, color=BLANCO)
    f_label = Font(size=10)
    f_label_b = Font(bold=True, size=10)
    f_num = Font(size=10)
    f_num_b = Font(bold=True, size=10)
    f_sd = Font(size=10, italic=True, color="94A3B8")
    fill_azul = PatternFill("solid", fgColor=AZUL)
    fill_azul_h = PatternFill("solid", fgColor=AZUL_H)
    fill_gris = PatternFill("solid", fgColor=GRIS)
    fill_res = PatternFill("solid", fgColor="E8EEF6")
    borde_fino = Border(bottom=Side(style="thin", color="D8DEE6"))

    ws.merge_cells("A1:C2")
    ws["A1"] = "P&G Filiales + Colombia + Consolidado — enero a julio 2026"
    ws["A1"].font = f_titulo
    ws["A1"].fill = fill_azul
    ws["A1"].alignment = Alignment(vertical="center")
    for c in ("B1", "C1", "A2", "B2", "C2"):
        ws[c].fill = fill_azul

    fila_nota = 3
    ws.merge_cells(f"A{fila_nota}:C{fila_nota}")
    ws[f"A{fila_nota}"] = ("Fuente: auxiliares contables homologados (BD_Indicadores, 13-ago-2026). "
                            "\"s/d\" = Colombia y Consolidado no tienen enero-marzo (ver nota abajo).")
    ws[f"A{fila_nota}"].font = Font(size=8.5, italic=True, color="6B7280")

    fila_ent = 5
    fila_moneda = 6
    fila_trm = 7
    fila_labels = 8
    fila0 = 9

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    col = 3  # C en adelante
    bloques = {}
    for mes in MESES:
        c0 = col
        ws.merge_cells(start_row=fila_ent - 1, start_column=c0, end_row=fila_ent - 1, end_column=c0 + len(ENTIDADES) - 1)
        celda_mes = ws.cell(row=fila_ent - 1, column=c0, value=NOMBRE_MES[mes])
        celda_mes.font = f_head
        celda_mes.fill = fill_azul_h
        celda_mes.alignment = Alignment(horizontal="center")
        for j, ent in enumerate(ENTIDADES):
            cc = c0 + j
            ce = ws.cell(row=fila_ent, column=cc, value=ETQ_ENT[ent])
            ce.font = f_ent
            ce.fill = fill_azul
            ce.alignment = Alignment(horizontal="center")
            cm = ws.cell(row=fila_moneda, column=cc, value=MONEDA_ORIGEN[ent])
            cm.font = Font(size=8, italic=True, color="6B7280")
            cm.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(cc)].width = 13
        ws.cell(row=fila_trm, column=c0, value=("TRM COP/USD %s: %.2f" % (mes[-2:], trm.get(mes, 0))))
        ws.cell(row=fila_trm, column=c0).font = Font(size=8, color="6B7280")
        bloques[mes] = c0
        col += len(ENTIDADES) + 1  # una columna de separación entre meses

    r = fila0
    for cod, etiqueta, tipo in LINEAS:
        if cod is None:
            r += 1
            continue
        celda_l = ws.cell(row=r, column=2, value=etiqueta)
        celda_l.font = f_label_b if tipo == "res" else f_label
        if tipo == "res":
            for cc in range(1, col):
                ws.cell(row=r, column=cc).fill = fill_res
        for mes in MESES:
            c0 = bloques[mes]
            for j, ent in enumerate(ENTIDADES):
                v = valor(ent, mes, cod)
                cc = ws.cell(row=r, column=c0 + j)
                if v is None:
                    cc.value = "s/d"
                    cc.font = f_sd
                    cc.alignment = Alignment(horizontal="center")
                else:
                    cc.value = round(v, 2)
                    cc.number_format = '#,##0.00;[RED](#,##0.00)'
                    cc.font = f_num_b if tipo == "res" else f_num
        r += 1

    for cc in range(1, col):
        ws.cell(row=fila0 - 1, column=cc).border = borde_fino

    ws.freeze_panes = "C9"

    r += 2
    ws.cell(row=r, column=2, value="Notas").font = Font(bold=True, size=10)
    r += 1
    notas = [
        "Costo del servicio y Operación/Proyectos: en cero en las 3 filiales — su contabilidad no distingue esas cuentas por separado (solo Administración/Ventas/Financieros).",
        "Colombia y Consolidado, enero-marzo: \"s/d\". El ingreso de Colombia guardado para esos meses es de gestión y no concilia con el PDF oficial (USD 514.030 vs USD 630.459 esperados); se dejó sin llenar en vez de mezclar dos criterios distintos.",
        "La reclasificación de honorarios de Perú (ajuste de consolidación, vive en el H1 oficial) no está incluida en esta vista mensual.",
    ]
    for n in notas:
        ws.cell(row=r, column=2, value="• " + n).font = Font(size=9, color="4B5563")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 26
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    wb.save(SALIDA)
    print("Generado:", os.path.abspath(SALIDA))


if __name__ == "__main__":
    construir()
