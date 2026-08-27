# -*- coding: utf-8 -*-
"""
Recalcula la fila "Consolidado" de un mes del bloque "Consolidacion USD" como la
suma de los cuatro paises.

POR QUE
-------
En ese bloque el consolidado NO se deriva: tiene fila propia guardada. Mientras
nadie toque un pais, coincide con la suma. Pero al corregir un pais --como julio
de Peru-- la fila queda con la cifra vieja y el estado deja de cuadrar con sus
propias columnas, sin que nada avise.

Este script cierra ese hueco. Antes de escribir comprueba que el invariante se
cumplia para los codigos NO afectados: si el consolidado ya venia descuadrado por
otra razon, aborta en vez de taparlo con una suma.

Regla del proyecto: hoy no hay eliminaciones intercompania practicadas, asi que
consolidado = suma de paises. El dia que se registren, esto hay que revisarlo.

Uso:  python recalcular_consolidado_mes.py 2026-07             -> vista previa
      python recalcular_consolidado_mes.py 2026-07 --escribir  -> aplica
"""
import os
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
RESPALDOS = os.path.join(AQUI, "respaldos")
HOJA = "BD_Indicadores"
SEGMENTO = "Consolidación USD"
PAISES = ["EE.UU.", "Perú", "Costa Rica", "Colombia"]
TOL = 0.01


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print("Falta el periodo. Ejemplo: python recalcular_consolidado_mes.py 2026-07")
        return 1
    periodo = args[0]
    escribir = "--escribir" in sys.argv

    wb = openpyxl.load_workbook(BASE, data_only=False)   # conserva las 3.882 formulas
    ws = wb[HOJA]

    enc = None
    for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True), 1):
        if row and row[0] and str(row[0]).strip() == "periodo":
            enc = (i, [str(x).strip() if x is not None else "" for x in row])
            break
    if not enc:
        print("No encuentro el encabezado.")
        return 1
    fila_enc, cols = enc
    c = {n: cols.index(n) + 1 for n in ("periodo", "pais", "codigo_indicador",
                                        "segmento", "valor", "comentario")}

    porpais, filacons = {}, {}
    for r in range(fila_enc + 1, ws.max_row + 1):
        if (str(ws.cell(r, c["periodo"]).value or "").strip() != periodo
                or str(ws.cell(r, c["segmento"]).value or "").strip() != SEGMENTO):
            continue
        pais = str(ws.cell(r, c["pais"]).value or "").strip()
        cod = str(ws.cell(r, c["codigo_indicador"]).value or "").strip()
        val = ws.cell(r, c["valor"]).value
        val = val if isinstance(val, (int, float)) else 0
        if pais in PAISES:
            porpais.setdefault(cod, {})[pais] = val
        elif pais == "Consolidado":
            filacons[cod] = (r, val)

    if not filacons:
        print("No hay fila Consolidado para %s en el bloque '%s'." % (periodo, SEGMENTO))
        return 1

    print("=" * 80)
    print("Consolidado %s = suma de %s" % (periodo, " + ".join(PAISES)))
    print("=" * 80)
    print("%-34s %13s %13s %11s  %s" % ("codigo", "SUMA PAISES", "GUARDADO", "DIF", ""))
    print("-" * 80)

    cambios, ya_descuadrados = [], []
    for cod, (r, guardado) in sorted(filacons.items()):
        if cod not in porpais:
            continue
        suma = round(sum(porpais[cod].values()), 2)
        dif = round(suma - guardado, 2)
        estado = "igual" if abs(dif) < TOL else "AJUSTA"
        print("%-34s %13.2f %13.2f %11.2f  %s" % (cod, suma, guardado, dif, estado))
        if abs(dif) >= TOL:
            cambios.append((cod, r, guardado, suma))

    print("-" * 80)
    print("codigos a ajustar: %d de %d" % (len(cambios), len(filacons)))

    if not cambios:
        print("El consolidado ya cuadra con la suma de paises. Nada que hacer.")
        return 0

    if not escribir:
        print()
        print("VISTA PREVIA. Nada se escribio. Corre con --escribir para aplicar.")
        return 0

    os.makedirs(RESPALDOS, exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    resp = os.path.join(RESPALDOS, "BD_MAESTRA_COCO_antes_consolidado_%s_%s.xlsx"
                        % (periodo.replace("-", ""), sello))
    shutil.copy2(BASE, resp)
    nota = ("Recalculado como suma de los cuatro paises tras corregir un pais del mes. "
            "Sin eliminaciones intercompania (no hay registradas).")
    for cod, r, _, suma in cambios:
        ws.cell(r, c["valor"]).value = suma
        ws.cell(r, c["comentario"]).value = nota
    wb.save(BASE)
    print()
    print("respaldo: %s" % os.path.basename(resp))
    print("base actualizada: %d filas del consolidado." % len(cambios))
    print("Corre Revisar_Base.bat para confirmar que las capas siguen alineadas.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
