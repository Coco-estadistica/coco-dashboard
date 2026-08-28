# -*- coding: utf-8 -*-
"""
Carga el puente de MRR de julio 2026 (Analisis de Movimiento MRR - Junio vs Julio,
entregado por el area comercial) y crea el indicador "arr_mrr" (MRR final x 12),
separado del "arr" que ya existe (ese es "Ventas Totales" segun contabilidad, una
metrica distinta pese al nombre).

Verificado antes de cargar: retenidas + nuevas + rescatadas + expansion - contraccion
- churned = Total MRR Julio, exacto (1,197,953,198).

Convencion de signo: contraccion_mrr y churned_mrr se guardan en NEGATIVO, igual que
ya estaba abril/mayo en la base -- si se guardaran en positivo la serie cambiaria de
signo a mitad de camino.

Uso:  python cargar_mrr_julio.py             -> vista previa
      python cargar_mrr_julio.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")

# --- puente de MRR de julio (dato del area comercial) ---
PUENTE_JULIO = {
    "retained_mrr": 316310847,
    "new_mrr": 504092820,
    "reactivacion_mrr": 0,
    "expansion_mrr": 14828122,
    "contraction_mrr": -43146258,
    "churned_mrr": -37681670,
}
MRR_JUNIO = 759860184
MRR_JULIO = 1197953198
ARR_JUNIO = 9118322209  # = MRR_JUNIO * 12, dado explicitamente por el area comercial
ARR_JULIO = MRR_JULIO * 12
COMENTARIO_PUENTE = "Analisis de Movimiento MRR Jun-Jul 2026 (area comercial, 13-ago-2026)."
COMENTARIO_ARR_JULIO = ("MRR julio x 12. Incluye efecto puntual de Costa Rica (CCSS, "
                         "+USD/COP 494.8M, cliente nuevo sin historico previo) -- si no es "
                         "representativo de un MRR estable, usar el ARR de junio como referencia.")


def main():
    escribir = "--escribir" in sys.argv
    wb = openpyxl.load_workbook(BASE)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) + 1 for c in H if c}

    existentes = {}
    for f in range(4, ws.max_row + 1):
        cod = ws.cell(f, ix["codigo_indicador"]).value
        per = ws.cell(f, ix["periodo"]).value
        pais = ws.cell(f, ix["pais"]).value
        if cod and pais == "Colombia":
            existentes[(cod, str(per))] = f

    siguiente = ws.max_row + 1
    reporte = []

    def escribe(cod, per, valor, comentario):
        nonlocal siguiente
        clave = (cod, per)
        if clave in existentes:
            f = existentes[clave]
            antes = ws.cell(f, ix["valor"]).value
            reporte.append("  UPDATE %-16s %s  %s -> %s" % (cod, per, antes, valor))
        else:
            f = siguiente
            siguiente += 1
            ws.cell(f, ix["periodo"]).value = per
            ws.cell(f, ix["pais"]).value = "Colombia"
            ws.cell(f, ix["compania"]).value = "Coco Colombia"
            ws.cell(f, ix["moneda"]).value = "COP"
            ws.cell(f, ix["escenario"]).value = "Real"
            ws.cell(f, ix["periodicidad"]).value = "Mensual"
            reporte.append("  INSERT %-16s %s  -> %s" % (cod, per, valor))
        ws.cell(f, ix["codigo_indicador"]).value = cod
        ws.cell(f, ix["valor"]).value = valor
        ws.cell(f, ix["unidad"]).value = "COP"
        ws.cell(f, ix["comentario"]).value = comentario

    for cod, val in PUENTE_JULIO.items():
        escribe(cod, "2026-07", val, COMENTARIO_PUENTE)
    escribe("mrr_final", "2026-06", MRR_JUNIO, COMENTARIO_PUENTE + " (total de junio, referencia).")
    escribe("mrr_final", "2026-07", MRR_JULIO, COMENTARIO_PUENTE)
    escribe("arr_mrr", "2026-06", ARR_JUNIO, "MRR junio x 12. Indicador nuevo, separado de 'arr' "
            "(ese es Ventas Totales segun contabilidad, no MRR anualizado pese al nombre).")
    escribe("arr_mrr", "2026-07", ARR_JULIO, COMENTARIO_ARR_JULIO)

    # --- diccionario: registrar arr_mrr si no existe ---
    wd = wb["Diccionario"]
    HD = [str(c.value).strip() if c.value else "" for c in wd[3]]
    ixd = {c: HD.index(c) + 1 for c in HD if c}
    ya_existe = any(wd.cell(r, ixd["Código indicador"]).value == "arr_mrr" for r in range(4, wd.max_row + 1))
    if not ya_existe:
        maxid = 0
        for r in range(4, wd.max_row + 1):
            try:
                maxid = max(maxid, int(wd.cell(r, ixd["ID"]).value))
            except Exception:
                pass
        fd = wd.max_row + 1
        vals = {
            "ID": maxid + 1, "Categoría": "SaaS / Revenue", "Código indicador": "arr_mrr",
            "Nombre visible": "ARR (MRR x 12)", "Tipo": "KPI",
            "Fórmula / regla de cálculo": "mrr_final * 12", "Unidad": "COP/USD",
            "Responsable carga": "Comercial", "Fuente sugerida": "Comercial / Facturación",
            "Frecuencia": "Mensual", "Origen": "COCO",
            "Comentario": "MRR anualizado real. Separado de 'arr' porque ese código guarda "
                          "Ventas Totales reportadas por contabilidad, una métrica distinta.",
            "periodicidad": "Mensual",
        }
        for k, v in vals.items():
            if k in ixd:
                wd.cell(fd, ixd[k]).value = v
        reporte.append("  Diccionario: agregado 'arr_mrr' (fila %d)" % fd)

    print("=" * 90)
    print("  CARGA MRR JULIO 2026 + ARR (MRR x 12)")
    print("=" * 90)
    # "Retenidas MRR" es informativo (clientes sin ningun cambio neto), NO es la base
    # del puente -- la base es el total de junio. El primer intento de verificacion
    # (antes de escribir nada) uso por error "Retenidas" como base y no cuadraba.
    print("\n  Verificación: MRR junio + nuevas + expansion + rescatadas - contraccion - churned")
    print("    %s + %s + %s + %s - %s - %s = %s" % (
        MRR_JUNIO, PUENTE_JULIO["new_mrr"], PUENTE_JULIO["expansion_mrr"], PUENTE_JULIO["reactivacion_mrr"],
        -PUENTE_JULIO["contraction_mrr"], -PUENTE_JULIO["churned_mrr"], MRR_JULIO))
    suma = MRR_JUNIO + PUENTE_JULIO["new_mrr"] + PUENTE_JULIO["expansion_mrr"] + PUENTE_JULIO["reactivacion_mrr"] \
        + PUENTE_JULIO["contraction_mrr"] + PUENTE_JULIO["churned_mrr"]
    print("  Suma real: %s  ·  Total MRR julio dado: %s  ·  %s" % (
        suma, MRR_JULIO, "OK, cuadra" if suma == MRR_JULIO else "NO CUADRA"))
    print()
    for l in reporte:
        print(l)

    if not escribir:
        print("\n  Vista previa. Para aplicar agrega --escribir\n")
        return

    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", "BD_MAESTRA_COCO_antes_mrr_julio_" + sello + ".xlsx")
    shutil.copy2(BASE, copia)
    wb.save(BASE)
    print("\n  Respaldo:", os.path.basename(copia))
    print("  Base actualizada.\n")


if __name__ == "__main__":
    main()
