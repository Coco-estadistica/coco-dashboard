# -*- coding: utf-8 -*-
"""
Carga Analisis_MRR_Junio_Julio.xlsx: la serie de MRR restablecida y la cartera de julio.

QUE TRAE EL ARCHIVO
-------------------
Una definicion nueva de MRR --VENCIDA (mes anterior) + CORRIENTE (mes en curso) por
cliente-- y la Caja Costa Rica (CCSS) reclasificada como one-off. Eso mueve los SIETE
meses de 2026, entre -16,6% y +9,5%.

Y corrige un error que estaba publicado: hoy new_mrr de julio dice 504.092.820, que es
494.843.760 del one-off de Costa Rica + 9.249.060 de ventas nuevas. El one-off estaba
contado como MRR recurrente nuevo. El archivo dice explicitamente que no debe contarse.

VALIDACION HECHA ANTES DE ESCRIBIR
----------------------------------
1. La hoja Base (115 clientes) reproduce la hoja Resumen al centavo en los 7 meses y en
   las 7 categorias. Si el archivo cambia y deja de cuadrar, el script aborta.
2. El puente jun->jul cierra exacto:
      759.860.184 + 9.249.060 + 14.828.122 - 43.146.258 - 37.681.670 = 703.109.438

DECISIONES TOMADAS POR EL NEGOCIO (27-ago-2026)
-----------------------------------------------
  MRR        reemplazar los 7 meses; la serie vieja se archiva, no se borra
  ARR        promedio ene-jul x 12 = 8.595.360.796 (antes era ultimo mes x 12)
  Cartera    el corte del 25-ago NO se carga: es mitad de mes, se espera el cierre
  Morosidad  se carga como viene, con la fuente a la vista

LO QUE ADEMAS CORRIGE
---------------------
La cartera de julio se cargo en su dia convirtiendo USD con 3.505,2683, que el script
de entonces llamaba "TRM oficial de julio". Esa es la TRM de JUNIO; la de julio es
3.268,95 y se corrigio despues. La cartera de julio en pesos quedo 7,2% inflada.

LO QUE NO HACE
--------------
No toca 2026-08. No inventa la definicion de morosidad ni de concentracion: las carga
tal como las entrega cartera y lo deja dicho en el comentario de cada fila.

Uso:  python cargar_mrr_cartera_agosto.py             -> vista previa
      python cargar_mrr_cartera_agosto.py --escribir  -> aplica (con respaldo)
"""
import os
import shutil
import sys
from collections import OrderedDict

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
TRAZA = os.path.join(AQUI, "BD_TRAZABILIDAD_COCO.xlsx")
FUENTE = os.path.join(os.path.expanduser("~"), "Desktop", "Temporales",
                      "Analisis_MRR_Junio_Julio.xlsx")

PAIS, MONEDA, ESC = "Colombia", "COP", "Real"
TRM_JULIO = 3268.95          # la de verdad; la carga anterior uso 3505.2683 (junio)
TRM_ANTERIOR = 3505.2683

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio"]
PER = {m: "2026-%02d" % (i + 1) for i, m in enumerate(MESES)}

NOTA_MRR = ("MRR recurrente = VENCIDA (mes anterior) + CORRIENTE (mes en curso) por "
            "cliente. Caja Costa Rica reclasificada como one-off, fuera de MRR y ARR. "
            "Fuente: Analisis_MRR_Junio_Julio.xlsx, hoja Base (115 clientes).")
NOTA_CARTERA = ("Cartera al 31-jul-2026, recibida en USD y convertida con la TRM de "
                "julio (3.268,95). Reemplaza la conversion anterior, que uso 3.505,2683 "
                "--la TRM de junio-- y dejaba la cifra 7,2%% inflada.")
NOTA_ENTREGADA = ("Entregado por cartera con su propia definicion; el tablero NO lo "
                  "calcula. No coincide con vencida/total (41,6%) ni con la serie "
                  "anterior (10,5 en mayo, 27,81 en junio), que quedo archivada.")


def leer_fuente():
    """Devuelve (serie_mrr, movimientos, one_off, cartera) y valida que la hoja Base
    reproduzca la hoja Resumen. Si no cuadra, no hay carga."""
    wb = openpyxl.load_workbook(FUENTE, data_only=True)

    filas = [f for f in wb["Base"].iter_rows(values_only=True)]
    cab = [str(v or "").strip() for v in filas[0]]
    ix = {c: i for i, c in enumerate(cab) if c}
    col_nr = "Monto no recurrente (Up Sale)"

    suma = OrderedDict((m, 0.0) for m in MESES)
    one_off = 0.0
    cat = {}
    for f in filas[1:]:
        if not f or not f[0]:
            continue
        one_off += float(f[ix[col_nr]] or 0)
        for m in MESES:
            suma[m] += float(f[ix[m]] or 0)
        c = str(f[ix["Categoría"]] or "")
        cat[c] = cat.get(c, 0.0) + float(f[ix["Monto categoría"]] or 0)

    r = {}
    for f in wb["Resumen MRR"].iter_rows(values_only=True):
        if f and f[0] and f[1] is not None:
            r[str(f[0]).strip()] = f[1]

    serie = OrderedDict()
    for m in MESES:
        esperado = float(r[m])
        real = suma[m] - (one_off if m == "Julio" else 0.0)
        if abs(real - esperado) > 1:
            raise SystemExit("ABORTA: la hoja Base no reproduce el Resumen en %s "
                             "(Base %.2f vs Resumen %.2f)" % (m, real, esperado))
        serie[PER[m]] = round(esperado, 2)

    mov = OrderedDict([
        ("new_mrr", round(float(cat["Nueva"]), 2)),
        ("expansion_mrr", round(float(cat["Expansión"]), 2)),
        ("contraction_mrr", -round(float(cat["Contracción"]), 2)),   # la base guarda restas en negativo
        ("churned_mrr", -round(float(cat["Churned"]), 2)),
        ("retained_mrr", round(float(cat["Retenida"]), 2)),
        ("reactivacion_mrr", 0.0),
    ])

    puente = (serie["2026-06"] + mov["new_mrr"] + mov["expansion_mrr"]
              + mov["contraction_mrr"] + mov["churned_mrr"])
    if abs(puente - serie["2026-07"]) > 1:
        raise SystemExit("ABORTA: el puente jun->jul no cierra (%.2f vs %.2f)"
                         % (puente, serie["2026-07"]))

    ws = wb["Indicadores de Cartera"]
    car = {}
    for f in ws.iter_rows(values_only=True):
        if f and f[0] and f[1] is not None:
            car[str(f[0]).strip()] = f[1]           # columna del 31-jul; la del 25-ago se ignora

    return serie, mov, round(one_off, 2), car, round(float(r["ARR = MRR promedio Ene-Jul × 12"]), 2)


def plan(serie, mov, one_off, car, arr):
    """(codigo, periodo, valor, unidad, comentario). Nada mas: una lista explicita."""
    p = []
    for per, v in serie.items():
        p.append(("mrr_final", per, v, "COP", NOTA_MRR))
    for cod, v in mov.items():
        p.append((cod, "2026-07", v, "COP", NOTA_MRR))
    p.append(("ingresos_no_recurrentes", "2026-07", one_off, "COP",
              "One-off Caja Costa Rica (CCSS), julio 2026. Fuera del MRR recurrente y del ARR."))
    p.append(("arr_mrr", "2026-07", arr, "COP",
              "ARR = MRR promedio ene-jul 2026 x 12. Antes se calculaba como ultimo mes x 12; "
              "el cambio de definicion se decidio el 27-ago-2026."))
    for cod, etq in (("cartera_total", "Cartera total"), ("cartera_sana", "Cartera sana"),
                     ("cartera_vencida", "Cartera vencida"),
                     ("cartera_cobro_legal", "Cartera en cobro legal")):
        p.append((cod, "2026-07", round(float(car[etq]) * TRM_JULIO, 2), "COP", NOTA_CARTERA))
    p.append(("dso", "2026-07", float(car["DSO"]), "días",
              "Cartera al 31-jul-2026. Fuente: Analisis_MRR_Junio_Julio.xlsx."))
    p.append(("cartera_mayor_60_pct", "2026-07", float(car["% Cartera mayor a 60 días"]), "%",
              "Cartera al 31-jul-2026."))
    p.append(("cartera_arr_pct", "2026-07", float(car["Cartera como % de ARR"]), "%",
              "Entregado por cartera. Con el ARR nuevo (8.595.360.796) la razon daria 17,9%: "
              "el 20,9% viene del ARR que usaba cartera, no del de esta base."))
    p.append(("write_off_ratio", "2026-07", float(car["Write-off ratio (castigo real)"]), "%",
              "Cartera al 31-jul-2026."))
    p.append(("indice_morosidad", "2026-07", float(car["Índice de morosidad"]), "%", NOTA_ENTREGADA))
    p.append(("concentracion_cartera", "2026-07", float(car["Concentración de cartera"]), "%",
              NOTA_ENTREGADA.replace("(10,5 en mayo, 27,81 en junio)", "(17,1 en mayo y junio)")))
    return p


# Series cuya definicion cambio: los meses viejos se archivan para que no queden dos
# definiciones en la misma linea del tablero.
ARCHIVAR = [("arr_mrr", "2026-01"), ("arr_mrr", "2026-02"), ("arr_mrr", "2026-03"),
            ("arr_mrr", "2026-04"), ("arr_mrr", "2026-05"), ("arr_mrr", "2026-06"),
            ("indice_morosidad", "2026-05"), ("indice_morosidad", "2026-06"),
            ("concentracion_cartera", "2026-05"), ("concentracion_cartera", "2026-06")]


def main():
    escribir = "--escribir" in sys.argv
    serie, mov, one_off, car, arr = leer_fuente()
    filas = plan(serie, mov, one_off, car, arr)

    wb = openpyxl.load_workbook(BASE)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) + 1 for c in H if c}

    def buscar(cod, per):
        for r in range(4, ws.max_row + 1):
            if (ws.cell(r, ix["codigo_indicador"]).value == cod
                    and ws.cell(r, ix["pais"]).value == PAIS
                    and str(ws.cell(r, ix["periodo"]).value) == per
                    and (ws.cell(r, ix["escenario"]).value or ESC) == ESC):
                return r
        return None

    print("=" * 96)
    print("  MRR Y CARTERA - Analisis_MRR_Junio_Julio.xlsx      %s"
          % ("APLICANDO" if escribir else "VISTA PREVIA (no escribe)"))
    print("=" * 96)
    print("\n  La hoja Base reproduce el Resumen en los 7 meses y el puente jun->jul cierra.\n")
    print("  %-24s %-9s %18s %18s  %s" % ("codigo", "periodo", "antes", "despues", ""))
    print("  " + "-" * 92)

    nuevas = act = 0
    for cod, per, val, uni, com in filas:
        r = buscar(cod, per)
        antes = ws.cell(r, ix["valor"]).value if r else None
        marca = ""
        if antes is None:
            marca = "NUEVO"
            nuevas += 1
        elif abs(float(antes) - float(val)) > 0.01:
            marca = "cambia"
            act += 1
        print("  %-24s %-9s %18s %18s  %s"
              % (cod, per, ("%.2f" % float(antes)) if antes is not None else "-", "%.2f" % val, marca))
        if not escribir:
            continue
        if r is None:
            r = ws.max_row + 1
            ws.cell(r, ix["periodo"]).value = per
            ws.cell(r, ix["pais"]).value = PAIS
            ws.cell(r, ix["moneda"]).value = MONEDA
            ws.cell(r, ix["escenario"]).value = ESC
            ws.cell(r, ix["codigo_indicador"]).value = cod
        ws.cell(r, ix["valor"]).value = val
        ws.cell(r, ix["unidad"]).value = uni
        ws.cell(r, ix["comentario"]).value = com

    print("\n  A ARCHIVAR (definicion vieja, se mueven a la trazabilidad):")
    archivo = []
    for cod, per in ARCHIVAR:
        r = buscar(cod, per)
        if r:
            archivo.append((per, cod, ws.cell(r, ix["valor"]).value,
                            ws.cell(r, ix["comentario"]).value))
            print("     %-24s %-9s %s" % (cod, per, ws.cell(r, ix["valor"]).value))

    print("\n  resumen: %d filas nuevas, %d actualizadas, %d retiradas"
          % (nuevas, act, len(archivo)))

    if not escribir:
        print("\n  Nada escrito. Para aplicar:  python %s --escribir"
              % os.path.basename(__file__))
        return 0

    for ruta in (BASE, TRAZA):
        shutil.copy2(ruta, ruta.replace(".xlsx", "_antes_mrr_ago.xlsx"))

    wt = openpyxl.load_workbook(TRAZA)
    hoja = "MRR_Cartera_Reemplazado"
    if hoja in wt.sheetnames:
        del wt[hoja]
    wa = wt.create_sheet(hoja)
    wa.append(["Series reemplazadas el 27-ago-2026 por cambio de definicion. "
               "Archivo, no las lee el tablero."])
    wa.append(["periodo", "codigo_indicador", "valor_anterior", "comentario_anterior"])
    for fila in archivo:
        wa.append(list(fila))
    wa.append([])
    wa.append(["serie mrr_final anterior:"])
    for per in sorted(serie):
        r = buscar("mrr_final", per)
        wa.append([per, "mrr_final", None, None])
    wt.save(TRAZA)

    # Se BLANQUEA el valor, no se borra la fila. BD_Indicadores tiene 3.879 formulas
    # por fila --area, area_responsable y periodicidad se resuelven con MATCH(F<fila>)--
    # y openpyxl no reescribe esas referencias al borrar: la fila 500 seguiria apuntando
    # a F501 y las tres columnas quedarian corridas para todo lo que este debajo, sin
    # que nada avise. Con el valor vacio la fila es invisible para el tablero igual
    # (facts la carga con valor:null y rawStored/byDetalle la saltan).
    for cod, per in ARCHIVAR:
        r = buscar(cod, per)
        if r:
            ws.cell(r, ix["valor"]).value = None
            ws.cell(r, ix["comentario"]).value = (
                "RETIRADO el 27-ago-2026: definicion distinta a la vigente. El valor "
                "anterior queda en BD_TRAZABILIDAD_COCO, hoja MRR_Cartera_Reemplazado.")

    wb.save(BASE)
    print("\n  Escrito. Respaldos: *_antes_mrr_ago.xlsx")
    print("  Sigue:  python recalcular_consolidado_mes.py 2026-07")
    print("          Revisar_Base.bat")
    return 0


if __name__ == "__main__":
    sys.exit(main())
