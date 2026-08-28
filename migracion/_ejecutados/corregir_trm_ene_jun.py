# -*- coding: utf-8 -*-
"""
Reemplaza la TRM "implicita" de enero a junio 2026, que estaba en 4.170 FIJO
los seis meses. Seis meses con la misma tasa al peso no es una tasa implicita:
era un valor de relleno que nunca se sustituyo.

El problema que causaba
-----------------------
La hoja de TRM solo afecta a los indicadores guardados en COP que el tablero
muestra en USD (MRR, ARR, cartera, gastos por area). Al corregir SOLO julio
(3.268,95) el 14-ago, la serie quedo partida: junio convertia a 4.170 y julio
a 3.268,95, y el MRR de julio aparecia creciendo 52,3% cuando en pesos crece
19,4%. Unos 33 puntos eran puro efecto cambiario.

De donde sale cada tasa
-----------------------
La propia hoja dice que la implicita es "derivada de ingresos COP / ingresos
USD". Donde existe el bloque de consolidacion de Colombia, se deriva exacto:

    abr  750.867.662 / 207.564,28 = 3.617,52
    may  826.514.746 / 222.539,06 = 3.714,02
    jun  751.453.314 / 214.378,26 = 3.505,27   <- era la de JUNIO
    jul  679.890.552 / 207.984,38 = 3.268,95   (ya corregida)

(Nota: 3.505,27 es la tasa de junio. Se habia usado por error como la de julio
hasta la correccion de esta manana.)

Enero a marzo NO son derivables: Colombia no tiene bloque en USD para esos
meses. Se usa la TRM promedio de mercado, que ya estaba en la misma hoja y cae
dentro del rango de las derivadas (3.269-3.714), muy lejos del 4.170.

Lo que NO cambia
----------------
El P&G oficial (codigos pyg_*) tiene su propio bloque "Consolidacion USD"
guardado; no se convierte con esta tasa. Ni el tablero de P&G ni la
presentacion de resultados cambian una sola cifra por este ajuste.

Uso:  python corregir_trm_ene_jun.py             -> vista previa
      python corregir_trm_ene_jun.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")

# periodo -> (tasa, origen)
NUEVAS = {
    "2026-01": (3703.00,  "TRM promedio de mercado (no derivable: sin bloque USD de Colombia)"),
    "2026-02": (3675.64,  "TRM promedio de mercado (no derivable: sin bloque USD de Colombia)"),
    "2026-03": (3717.56,  "TRM promedio de mercado (no derivable: sin bloque USD de Colombia)"),
    "2026-04": (3617.52,  "Derivada: ingresos COP / ingresos USD de la consolidacion"),
    "2026-05": (3714.02,  "Derivada: ingresos COP / ingresos USD de la consolidacion"),
    "2026-06": (3505.27,  "Derivada: ingresos COP / ingresos USD de la consolidacion"),
}
# control: la derivada debe reproducir el USD guardado
CONTROL = {
    "2026-04": (750867662, 207564.28),
    "2026-05": (826514746, 222539.06),
    "2026-06": (751453314, 214378.26),
}


def main():
    escribir = "--escribir" in sys.argv

    print("=" * 92)
    print("  CORREGIR TRM IMPLICITA  ene-jun 2026   (4.170 fijo -> tasa real del mes)")
    print("=" * 92)
    print("\n  Control de las derivadas:")
    for per, (cop, usd) in CONTROL.items():
        calc = cop / usd
        prop = NUEVAS[per][0]
        estado = "OK" if abs(calc - prop) < 0.02 else "NO CUADRA"
        print("    %s  %s / %s = %.2f   vs propuesta %.2f   %s"
              % (per, format(cop, ","), format(usd, ","), calc, prop, estado))
        if estado != "OK":
            print("\n  Se aborta.\n")
            return

    wb = openpyxl.load_workbook(BASE)
    wt = wb["TRM"]

    fila = {}
    for r in range(1, wt.max_row + 1):
        p = str(wt.cell(r, 1).value)
        if p in NUEVAS and p not in fila:
            fila[p] = r

    print("\n  Cambios en la hoja TRM:")
    for per, (tasa, origen) in NUEVAS.items():
        r = fila.get(per)
        if not r:
            print("    AVISO %s no encontrado en la hoja" % per)
            continue
        antes = wt.cell(r, 2).value
        print("    %s  %s -> %s   (%s)" % (per, antes, tasa, origen.split(":")[0]))
        wt.cell(r, 2).value = tasa
        wt.cell(r, 3).value = origen

    # --- impacto sobre el MRR, que es lo que se ve en el tablero ---
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) for c in H if c}
    mrr = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[ix["codigo_indicador"]] == "mrr_final" and row[ix["pais"]] == "Colombia":
            mrr[str(row[ix["periodo"]])] = row[ix["valor"]]

    print("\n  Efecto sobre 'MRR final por mes' (USD):")
    for per in ["2026-04", "2026-05", "2026-06", "2026-07"]:
        cop = mrr.get(per)
        if not cop:
            continue
        vieja = 4170.0 if per != "2026-07" else 3268.95
        nueva = NUEVAS.get(per, (3268.95, ""))[0]
        print("    %s   antes USD %9s   ahora USD %9s"
              % (per, format(int(cop / vieja), ","), format(int(cop / nueva), ",")))
    j6 = mrr.get("2026-06"); j7 = mrr.get("2026-07")
    if j6 and j7:
        antes = (j7 / 3268.95) / (j6 / 4170.0) - 1
        ahora = (j7 / 3268.95) / (j6 / 3505.27) - 1
        print("\n    Variacion jul vs jun en USD:  antes %+.1f%%   ahora %+.1f%%   (en COP: %+.1f%%)"
              % (antes * 100, ahora * 100, (j7 / j6 - 1) * 100))

    if not escribir:
        print("\n  Vista previa. Para aplicar agrega --escribir\n")
        return

    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", "BD_MAESTRA_COCO_antes_trm_ene_jun_" + sello + ".xlsx")
    shutil.copy2(BASE, copia)
    wb.save(BASE)
    print("\n  Respaldo:", os.path.basename(copia))
    print("  Base actualizada.\n")


if __name__ == "__main__":
    main()
