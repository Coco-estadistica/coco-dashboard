# -*- coding: utf-8 -*-
"""
Corrige julio 2026 de Peru en BD_Indicadores (bloque "Consolidacion USD").

POR QUE
-------
El archivo "Julio 2026 PERU.xlsx" que se habia cargado NO era julio: era un
reexporte de junio. Se comprueba en el saldo inicial de la clase 4 - INGRESOS,
que abria en -78.065,95 (el inicial de junio) en vez de -84.268,20 (el cierre de
junio). De ahi salio el hallazgo C-04 de la auditoria: los mismos 6.202,25 soles
aparecian dos veces, convertidos con dos tasas distintas --3,4050 en junio y
2,2204 en julio-- y ademas la base quedo con dos cifras contradictorias para el
mismo mes: pyg_servicio_software 1.821,51 contra pyg_ingresos_operacionales
2.793,35, con devoluciones en cero.

El archivo nuevo ("EF Peru.zip", 27-ago-2026) si es julio: encadena con junio y
trae movimiento propio. Verificado: la cadena de saldos cierra en los 7 meses.

MAPEO CUENTA -> LINEA
---------------------
Deducido de los meses que NO cambiaron (abril, mayo y junio) y validado contra lo
que el tablero ya tenia cargado. Coincide al centavo en los tres:

    ingresos       4170
    administracion 5110
    ventas         5210 + 5235 + 5255 + 5295
    financieros    5305
    proyectos      (Peru no tiene)
    4295           = ProInnovate, va a su bloque aparte y NO al consolidado

ALCANCE
-------
Solo se tocan las filas de Peru, periodo 2026-07, segmento "Consolidacion USD".
NO se toca el bloque "Acumulado H1 2026": ese es el estado combinado oficial
publicado y cambiarlo es una decision de negocio, no de cargue. Marzo tambien
quedo legible por primera vez (el archivo viejo era texto tabulado con extension
.xlsx) pero vive dentro del H1, asi que queda pendiente por la misma razon.

Uso:  python cargar_peru_julio_corregido.py             -> vista previa
      python cargar_peru_julio_corregido.py --escribir  -> aplica, con respaldo
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

PERIODO = "2026-07"
PAIS = "Perú"
SEGMENTO = "Consolidación USD"
TRM_JULIO = 3.400          # PEN por USD. Fuente: hoja TRM_Peru, BCRP promedio mensual
FUENTE = ("Julio 2026 corregido desde EF Peru.zip (27-ago-2026): el archivo anterior era "
          "un reexporte de junio. Convertido a 3,400 PEN/USD (BCRP promedio mensual).")

# Movimiento del mes en soles, tomado del balance de julio
PEN = {
    "4170": 9511.36,   # servicio de software (credito)
    "5110": 862.40,    # honorarios -> administracion
    "5235": 6858.00,   # servicios   -> ventas
    "5295": 4793.60,   # diversos    -> ventas
    "5305": 92.80,     # financieros
}

ing = PEN["4170"] / TRM_JULIO
admin = PEN["5110"] / TRM_JULIO
ventas = (PEN["5235"] + PEN["5295"]) / TRM_JULIO
fin = PEN["5305"] / TRM_JULIO
neta = ing - admin - ventas - fin

# la base tambien guarda los subtotales: si no se actualizan quedan con la cifra
# vieja y la tabla deja de cuadrar con sus propias lineas
bruta = ing            # Peru no tiene costo de ventas
operativa = ing - admin - ventas

NUEVO = {
    "pyg_utilidad_bruta": round(bruta, 2),
    "pyg_utilidad_operativa": round(operativa, 2),
    "pyg_costo_ventas": 0.0,
    "pyg_servicio_software": round(ing, 2),
    "pyg_ingresos_operacionales": round(ing, 2),
    "pyg_devoluciones": 0.0,
    "pyg_gasto_administracion": round(admin, 2),
    "pyg_gasto_ventas": round(ventas, 2),
    "pyg_gasto_proyectos": 0.0,
    "pyg_ingresos_no_operacionales": 0.0,
    "pyg_gastos_financieros": round(fin, 2),
    "pyg_utilidad_neta": round(neta, 2),
}


def main():
    escribir = "--escribir" in sys.argv

    # data_only=False es obligatorio: BD_Indicadores tiene 3.879 formulas y
    # Calc_LTV_Fin otras 3. Con data_only=True, openpyxl las reemplaza por su
    # valor cacheado y al guardar se pierden en TODO el libro, incluidas las
    # hojas que este script no toca. Las columnas que leemos son valores planos,
    # asi que no hace falta el modo de valores.
    wb = openpyxl.load_workbook(BASE, data_only=False)
    ws = wb[HOJA]

    enc = None
    for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True), 1):
        if row and row[0] and str(row[0]).strip() == "periodo":
            enc = (i, [str(x).strip() if x is not None else "" for x in row])
            break
    if not enc:
        print("No encuentro el encabezado de BD_Indicadores.")
        return 1
    fila_enc, cols = enc
    c = {n: cols.index(n) + 1 for n in ("periodo", "pais", "codigo_indicador",
                                        "segmento", "valor", "comentario", "moneda")}

    encontradas, cambios = 0, []
    for r in range(fila_enc + 1, ws.max_row + 1):
        if (str(ws.cell(r, c["periodo"]).value or "").strip() != PERIODO
                or str(ws.cell(r, c["pais"]).value or "").strip() != PAIS
                or str(ws.cell(r, c["segmento"]).value or "").strip() != SEGMENTO):
            continue
        cod = str(ws.cell(r, c["codigo_indicador"]).value or "").strip()
        if cod not in NUEVO:
            continue
        encontradas += 1
        antes = ws.cell(r, c["valor"]).value
        despues = NUEVO[cod]
        cambios.append((cod, antes, despues, r))
        if escribir:
            ws.cell(r, c["valor"]).value = despues
            ws.cell(r, c["comentario"]).value = FUENTE

    print("=" * 78)
    print("PERU julio 2026 -- correccion del bloque Consolidacion USD")
    print("=" * 78)
    print("%-32s %14s %14s %12s" % ("codigo", "ANTES", "DESPUES", "DIF"))
    print("-" * 78)
    for cod, a, d, _ in sorted(cambios):
        a = a if isinstance(a, (int, float)) else 0
        print("%-32s %14.2f %14.2f %12.2f" % (cod, a, d, d - a))
    print("-" * 78)
    print("filas encontradas: %d de %d codigos" % (encontradas, len(NUEVO)))
    faltan = set(NUEVO) - {x[0] for x in cambios}
    if faltan:
        print("SIN FILA EN LA BASE (no se crean, se avisan): %s" % ", ".join(sorted(faltan)))

    # control interno: la utilidad debe cerrar con sus propias lineas
    chk = (NUEVO["pyg_ingresos_operacionales"] - NUEVO["pyg_gasto_administracion"]
           - NUEVO["pyg_gasto_ventas"] - NUEVO["pyg_gasto_proyectos"]
           - NUEVO["pyg_gastos_financieros"])
    print()
    print("control: ingresos - gastos = %.2f  vs utilidad neta %.2f  -> %s"
          % (chk, NUEVO["pyg_utilidad_neta"],
             "OK" if abs(chk - NUEVO["pyg_utilidad_neta"]) < 0.01 else "NO CUADRA"))

    if not escribir:
        print()
        print("VISTA PREVIA. Nada se escribio. Corre con --escribir para aplicar.")
        return 0

    os.makedirs(RESPALDOS, exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    resp = os.path.join(RESPALDOS, "BD_MAESTRA_COCO_antes_peru_julio_%s.xlsx" % sello)
    shutil.copy2(BASE, resp)
    print()
    print("respaldo: %s" % os.path.basename(resp))
    wb.save(BASE)
    print("base actualizada.")
    print("Corre Revisar_Base.bat para confirmar que las capas siguen alineadas.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
