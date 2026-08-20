# -*- coding: utf-8 -*-
"""
AUDITORIA DE ALCANCE POR INDICADOR

Por que existe
--------------
El filtro Pais del tablero tiene tres mundos:

  Colombia            -> facts en COP, segmento vacio
  EE.UU./Peru/CR/Consolidado -> bloque 'Consolidacion USD'
  Todos (__ALL__)     -> NO filtra por pais, y ADEMAS excluye 'Consolidacion USD'

Esa tercera regla es el problema: como Costa Rica no tiene ni una fila fuera del
bloque consolidado, y EE.UU./Peru tienen 12 cada uno, 'Todos' termina leyendo
98% Colombia y presentandolo como si fuera el consolidado de los cuatro paises.

Este script clasifica cada codigo segun donde vive realmente, para saber que
puede consolidarse de verdad y que no. La clasificacion se deduce de los datos,
no de una lista escrita a mano: si manana Costa Rica carga MRR, el codigo cambia
de grupo solo.

Uso:  python auditar_alcance.py
      python auditar_alcance.py --csv   -> ademas escribe alcance_indicadores.csv
"""
import csv
import os
import sys
from collections import defaultdict

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
INTL = {"Consolidación USD", "PROINNOVATE", "Original enero",
        "Revisado mayo", "Acumulado H1 2026", "Modelo"}
CONSOL = "Consolidación USD"
PAISES = ["Colombia", "EE.UU.", "Perú", "Costa Rica"]


def main():
    wb = openpyxl.load_workbook(BASE, read_only=True, data_only=True)
    filas = list(wb["BD_Indicadores"].iter_rows(values_only=True))
    H = list(filas[2])
    ix = {c: H.index(c) for c in H if c}
    hechos = [r for r in filas[3:] if r[ix["codigo_indicador"]]]

    normal = defaultdict(set)   # codigo -> paises fuera del bloque consolidado
    consol = defaultdict(set)   # codigo -> paises dentro del bloque consolidado
    ultimo = defaultdict(str)   # codigo -> ultimo periodo con dato
    for r in hechos:
        cod = r[ix["codigo_indicador"]]
        pais = r[ix["pais"]]
        seg = r[ix["segmento"]] or ""
        per = str(r[ix["periodo"]] or "")
        if r[ix["valor"]] is not None and per > ultimo[cod]:
            ultimo[cod] = per
        if seg == CONSOL:
            consol[cod].add(pais)
        elif seg not in INTL:
            normal[cod].add(pais)

    grupos = defaultdict(list)
    for cod in sorted(set(normal) | set(consol)):
        n, c = normal[cod], consol[cod]
        if "Consolidado" in c:
            g = "A. CONSOLIDABLE"          # existe el bloque consolidado real
        elif len(c) > 1:
            g = "B. VARIOS PAISES"         # filiales si, Consolidado no
        elif n == {"Colombia"} and not c:
            g = "C. SOLO COLOMBIA"         # el caso que hoy se disfraza
        elif not n and not c:
            g = "E. SIN DATOS"
        else:
            g = "D. PARCIAL"
        grupos[g].append((cod, sorted(n), sorted(c), ultimo[cod]))

    print("=" * 86)
    print("  ALCANCE REAL DE CADA INDICADOR")
    print("=" * 86)
    for g in sorted(grupos):
        print(f"\n  {g}  ({len(grupos[g])} codigos)")
        for cod, n, c, ult in grupos[g]:
            donde = ("consol:" + "/".join(c)) if c else ("normal:" + "/".join(n) or "-")
            print(f"     {cod:34s} ult {ult or 's/d':8s} {donde[:44]}")

    tot = sum(len(v) for v in grupos.values())
    sc = len(grupos.get("C. SOLO COLOMBIA", []))
    print("\n" + "=" * 86)
    print(f"  {tot} codigos · {sc} son SOLO COLOMBIA")
    print("  Esos son los que hoy, con el filtro en 'Todos', se muestran como si")
    print("  fueran el consolidado de los cuatro paises.")
    print("=" * 86)

    if "--csv" in sys.argv:
        sal = os.path.join(AQUI, "alcance_indicadores.csv")
        with open(sal, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["grupo", "codigo", "paises_mundo_normal",
                        "paises_bloque_consolidado", "ultimo_periodo"])
            for g in sorted(grupos):
                for cod, n, c, ult in grupos[g]:
                    w.writerow([g, cod, "/".join(n), "/".join(c), ult])
        print(f"\n  Escrito: {os.path.basename(sal)}")


if __name__ == "__main__":
    main()
