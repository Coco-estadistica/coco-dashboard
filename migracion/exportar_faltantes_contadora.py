# -*- coding: utf-8 -*-
"""
Exporta a Excel los indicadores con hueco en BD_Indicadores (fuera de Ventas,
Marketing y Customer Success) EXCLUYENDO los que ya tienen algo en el modelo
financiero (esos se llenan aparte, del modelo). Queda listo para enviarle a
la contadora, con una columna por mes para que lo diligencie directo.
"""
import json
import os

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(AQUI, "..", "Indicadores_faltantes_para_contadora.xlsx")
MESES = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07"]
NOMBRE_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]

# Excluidos: ya tienen algo aprovechable en el modelo financiero (Seg. Ppto Jun-26
# y el informe Burn & Runway). Esos se completan desde ahi, no se le piden a la
# contadora.
EXCLUIR_EN_MODELO = {
    "ingresos_totales", "ebitda", "utilidad_neta", "margen_bruto_pct", "margen_ebitda",
    "burn_rate_neto", "runway_meses",
}

GRUPOS = [
    ("Finanzas generales", [
        "gastos_totales", "egresos_totales",
    ]),
    ("MRR / Ingresos recurrentes", [
        "arr", "arpa", "mrr_final", "ingresos_recurrentes", "ingresos_no_recurrentes",
        "new_mrr", "expansion_mrr", "contraction_mrr", "churned_mrr", "retained_mrr",
        "reactivacion_mrr", "churn_mrr_mensual",
    ]),
    ("Cartera y liquidez", [
        "cartera_total", "cartera_vencida", "cartera_sana", "cartera_vencida_cliente",
        "dso", "dpo", "ccc", "liquidez_total", "cartera_cobro_legal", "indice_morosidad",
        "cartera_mayor_60_pct", "cartera_arr_pct", "concentracion_cartera", "write_off_ratio",
    ]),
    ("Costos y gastos por área", [
        "gasto_area", "gasto_admin", "gasto_comercial", "gasto_desarrollo",
        "costo_operacion", "costo_proveedor", "costo_proveedor_ventas_pct",
        "salario_area_pct", "distribucion_ingresos",
    ]),
]

AZUL = "1E3A5F"
AZUL_H = "17324C"
GRIS = "F5F7FA"
VERDE_S = "D7EFDD"
ROJO_S = "FBE0DF"


def main():
    matriz = json.load(open(os.path.join(AQUI, "matriz_faltantes.json"), encoding="utf-8"))
    nombres = json.load(open(os.path.join(AQUI, "nombres_dic.json"), encoding="utf-8"))
    por_codigo = {r[0]: r for r in matriz}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Faltantes"
    ws.sheet_view.showGridLines = False

    f_tit = Font(bold=True, size=13, color="FFFFFF")
    f_sub = Font(italic=True, size=9, color="6B7280")
    f_grupo = Font(bold=True, size=11, color="FFFFFF")
    f_head = Font(bold=True, size=9.5, color="FFFFFF")
    f_cod = Font(size=9, color="6B7280", italic=True)
    f_nom = Font(size=10)
    fill_azul = PatternFill("solid", fgColor=AZUL)
    fill_azul_h = PatternFill("solid", fgColor=AZUL_H)
    fill_gris = PatternFill("solid", fgColor=GRIS)
    borde = Border(bottom=Side(style="thin", color="D8DEE6"))
    borde_celda = Border(top=Side(style="thin", color="E5E9F0"), bottom=Side(style="thin", color="E5E9F0"),
                          left=Side(style="thin", color="E5E9F0"), right=Side(style="thin", color="E5E9F0"))

    ws.merge_cells("A1:K2")
    ws["A1"] = "Indicadores por completar — para diligenciar por mes"
    ws["A1"].font = f_tit
    ws["A1"].fill = fill_azul
    ws["A1"].alignment = Alignment(vertical="center", horizontal="left", indent=1)
    for col in range(1, 12):
        ws.cell(1, col).fill = fill_azul
        ws.cell(2, col).fill = fill_azul

    ws.merge_cells("A3:K3")
    ws["A3"] = ("No incluye Ventas/Comercial, Marketing ni Customer Success (se manejan aparte), "
                "ni lo que ya sale del modelo financiero (ingresos, EBITDA, utilidad neta, "
                "márgenes, burn y runway — esos se completan desde ahí).")
    ws["A3"].font = f_sub

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 15
    for i in range(7):
        ws.column_dimensions[get_column_letter(5 + i)].width = 9
    ws.column_dimensions["L"].width = 24

    fila = 5
    ws.cell(fila, 2, "Indicador").font = f_head
    ws.cell(fila, 3, "Unidad").font = f_head
    ws.cell(fila, 4, "Último dato en BD").font = f_head
    for j, m in enumerate(NOMBRE_MES):
        c = ws.cell(fila, 5 + j, m)
        c.font = f_head
        c.alignment = Alignment(horizontal="center")
    ws.cell(fila, 12, "Fuente sugerida").font = f_head
    for col in range(2, 13):
        ws.cell(fila, col).fill = fill_azul_h
    ws.row_dimensions[fila].height = 20
    fila += 1

    total_filas = 0
    for titulo_grupo, codigos in GRUPOS:
        codigos_incluir = [c for c in codigos if c not in EXCLUIR_EN_MODELO]
        if not codigos_incluir:
            continue
        ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=12)
        cg = ws.cell(fila, 2, titulo_grupo)
        cg.font = f_grupo
        cg.fill = fill_azul
        cg.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[fila].height = 20
        fila += 1

        for cod in codigos_incluir:
            nombre, unidad, fuente = nombres.get(cod, (cod, "", ""))
            r = por_codigo.get(cod)
            ultimo = r[3] if r else "nunca"
            faltan = set(r[4]) if r else set(MESES)

            ws.cell(fila, 2, nombre or cod).font = f_nom
            ws.cell(fila, 2).alignment = Alignment(indent=1)
            sub = ws.cell(fila, 2)
            ws.cell(fila, 3, unidad or "").font = f_cod
            ws.cell(fila, 3).alignment = Alignment(horizontal="center")
            et = ws.cell(fila, 4, ultimo)
            et.font = f_cod
            et.alignment = Alignment(horizontal="center")

            for j, m in enumerate(MESES):
                c = ws.cell(fila, 5 + j)
                if m in faltan:
                    c.value = ""
                    c.fill = PatternFill("solid", fgColor=ROJO_S)
                else:
                    c.value = "OK"
                    c.fill = PatternFill("solid", fgColor=VERDE_S)
                    c.font = Font(size=8, color="2F6B45")
                    c.alignment = Alignment(horizontal="center")
                c.border = borde_celda

            ws.cell(fila, 12, fuente or "").font = Font(size=8.5, color="6B7280")
            ws.cell(fila, 12).alignment = Alignment(wrap_text=True)

            # fila auxiliar con el codigo tecnico, chiquita
            fila += 1
            total_filas += 1

    ws.freeze_panes = "E6"

    # leyenda
    fila += 1
    ws.cell(fila, 2, "OK = ya está cargado · vacío en rojo = falta ese mes").font = Font(size=8.5, italic=True, color="6B7280")

    wb.save(SALIDA)
    print("Generado:", os.path.abspath(SALIDA), "-", total_filas, "indicadores")


if __name__ == "__main__":
    main()
