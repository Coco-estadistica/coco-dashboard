# -*- coding: utf-8 -*-
"""
Actualiza la serie de MRR con la tabla completa jul-2025 a jul-2026 entregada
por el area comercial (14-ago-2026).

QUE SE ENCONTRO
---------------
La tabla define el total del mes como:

    Total MRR = Retenidas + Nuevas + Rescatadas + Expansion

Verificado: cuadra EXACTO en los 13 meses. NO es un puente encadenado (probado
tambien: encadenar el total del mes anterior da diferencias de hasta 302 MM).
Es decir, "Retenidas" ya viene neta de contraccion y churn -- esas dos lineas
se reportan como informacion de lo perdido, no se vuelven a restar.

Once de los trece meses (jul-2025 a may-2026) YA estaban en la base con estos
mismos valores, al peso. Solo cambian dos meses:

  2026-06  no tenia ningun componente del puente (solo el total). Se cargan los
           seis. Y su total baja de 759.860.184 a 699.523.252.
  2026-07  los componentes coinciden, pero el total baja de 1.197.953.198 a
           835.231.789.

POR QUE CAMBIAN ESOS DOS TOTALES
--------------------------------
Venian del informe anterior ("Analisis de Movimiento MRR - Junio vs Julio"),
que uso una definicion DISTINTA: encadenaba el total de junio y le restaba
contraccion y churn. Con esa formula julio daba 1.197.953.198.

Los otros once meses de la base siempre usaron la definicion de esta tabla. Es
decir, junio y julio eran los dos unicos meses inconsistentes con su propia
serie. Se alinean con los once restantes, no al reves.

arr_mrr (= mrr_final x 12) se recalcula para los trece meses: hasta hoy solo
existia para junio y julio, y con dos totales que cambian.

Uso:  python actualizar_mrr_serie_completa.py             -> vista previa
      python actualizar_mrr_serie_completa.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")

MESES = ["2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12",
         "2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07"]

SERIES = {
    "retained_mrr":    [665215212, 645143880, 599113206, 621530361, 647442178, 533747198,
                        554242648, 626918165, 666914432, 685135673, 638091793, 656948972, 316310847],
    "new_mrr":         [0, 30328404, 66758720, 9744000, 4576586, 121130467,
                        42197000, 67164125, 20433810, 19704840, 57840476, 0, 504092820],
    "expansion_mrr":   [33230383, 10257072, 20728046, 32161155, 30488403, 7435487,
                        62692450, 139839642, 60430077, 37926081, 10796596, 42574280, 14828122],
    "contraction_mrr": [-66478867, -53026479, -47500121, -40805021, -46520913, -63176870,
                        -37202034, -11016136, -25206134, -17709872, -43867662, -22143677, -43146258],
    "churned_mrr":     [0, -11465728, -1414500, -906843, -5800000, -1685823,
                        -35983665, -9736599, -11358174, -1358174, 0, -1573424, -37681670],
}
# la fila "Rescatadas" viene en blanco salvo sep-2025 y jul-2026, ambos en cero.
# Como el total cuadra sin ella en los trece meses, los blancos son cero de
# hecho; aun asi NO se rellenan: la fuente no los reporta y un blanco no es un
# cero confirmado. Los dos que si vienen ya estan en la base.
TOTAL = [698445595, 685729356, 686599972, 663435516, 682507167, 662313152,
         659132098, 833921932, 747778319, 742766594, 706728865, 699523252, 835231789]

COMENTARIO = ("Serie MRR jul-2025 a jul-2026 (area comercial, 14-ago-2026). "
              "Total del mes = Retenidas + Nuevas + Rescatadas + Expansion; contraccion y "
              "churn se reportan como perdida del periodo, ya netas dentro de Retenidas.")
COMENTARIO_ARR = "MRR del mes x 12. Recalculado sobre la serie unificada del 14-ago-2026."


def main():
    escribir = "--escribir" in sys.argv

    # --- verificacion antes de tocar nada ---
    print("=" * 96)
    print("  ACTUALIZACION SERIE MRR  jul-2025 a jul-2026")
    print("=" * 96)
    print("\n  Verificacion: Total = Retenidas + Nuevas + Expansion")
    fallo = False
    for i, m in enumerate(MESES):
        s = SERIES["retained_mrr"][i] + SERIES["new_mrr"][i] + SERIES["expansion_mrr"][i]
        if s != TOTAL[i]:
            fallo = True
            print("    %s  NO CUADRA  %s vs %s" % (m, format(s, ","), format(TOTAL[i], ",")))
    if fallo:
        print("\n  Se aborta: la tabla no cuadra consigo misma.\n")
        return
    print("    los 13 meses cuadran exacto\n")

    wb = openpyxl.load_workbook(BASE)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) + 1 for c in H if c}

    existentes = {}
    for f in range(4, ws.max_row + 1):
        cod = ws.cell(f, ix["codigo_indicador"]).value
        per = str(ws.cell(f, ix["periodo"]).value)
        if cod and ws.cell(f, ix["pais"]).value == "Colombia":
            existentes[(cod, per)] = f

    siguiente = ws.max_row + 1
    inserta, actualiza, iguales = [], [], 0

    def escribe(cod, per, valor, comentario):
        nonlocal siguiente, iguales
        clave = (cod, per)
        if clave in existentes:
            f = existentes[clave]
            antes = ws.cell(f, ix["valor"]).value
            if antes is not None and abs(antes - valor) < 0.5:
                iguales += 1
                return
            actualiza.append("  UPDATE %-16s %s  %s -> %s"
                             % (cod, per, format(antes, ",") if antes is not None else "—",
                                format(valor, ",")))
        else:
            f = siguiente
            siguiente += 1
            existentes[clave] = f
            ws.cell(f, ix["periodo"]).value = per
            ws.cell(f, ix["pais"]).value = "Colombia"
            ws.cell(f, ix["compania"]).value = "Coco Colombia"
            ws.cell(f, ix["moneda"]).value = "COP"
            ws.cell(f, ix["escenario"]).value = "Real"
            ws.cell(f, ix["periodicidad"]).value = "Mensual"
            inserta.append("  INSERT %-16s %s  -> %s" % (cod, per, format(valor, ",")))
        ws.cell(f, ix["codigo_indicador"]).value = cod
        ws.cell(f, ix["valor"]).value = valor
        ws.cell(f, ix["unidad"]).value = "COP"
        ws.cell(f, ix["comentario"]).value = comentario

    for cod, vals in SERIES.items():
        for i, m in enumerate(MESES):
            escribe(cod, m, vals[i], COMENTARIO)
    for i, m in enumerate(MESES):
        escribe("mrr_final", m, TOTAL[i], COMENTARIO)
        escribe("arr_mrr", m, TOTAL[i] * 12, COMENTARIO_ARR)

    for l in inserta:
        print(l)
    if inserta:
        print()
    for l in actualiza:
        print(l)
    print("\n  %d filas nuevas · %d actualizadas · %d ya estaban igual (no se tocan)"
          % (len(inserta), len(actualiza), iguales))

    if not escribir:
        print("\n  Vista previa. Para aplicar agrega --escribir\n")
        return

    os.makedirs(os.path.join(AQUI, "respaldos"), exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, "respaldos", "BD_MAESTRA_COCO_antes_mrr_serie_" + sello + ".xlsx")
    shutil.copy2(BASE, copia)
    wb.save(BASE)
    print("\n  Respaldo:", os.path.basename(copia))
    print("  Base actualizada.\n")


if __name__ == "__main__":
    main()
