# -*- coding: utf-8 -*-
"""
Regenera la cadena de gastos de BD_MAESTRA_COCO.xlsx.

    BD_GASTOS_HOMOLOGADOS  +  MAP_DASHBOARD_GASTOS
              |
              v
    BD_GASTOS_DASHBOARD_BRIDGE   (una fila por gasto homologado, con su linea de tablero)
              |
              v
    BD_GASTOS_RESUMEN_MENSUAL    (el agregado que el tablero consume)

Hoy esa cadena se rehace a mano. Este script la reproduce exactamente: no cambia
ninguna cifra, solo automatiza el paso.

POR QUE data_only=False
-----------------------
BD_Indicadores tiene 3.879 formulas y Calc_LTV_Fin otras 3. openpyxl cargado con
data_only=True reemplaza cada formula por su valor cacheado, y al guardar las borra
en TODO el libro -- incluidas las hojas que este script tiene prohibido tocar. Se
perderian 3.882 formulas de forma permanente y el unico sitio donde quedarian seria
el respaldo. Las hojas de gastos no tienen ni una formula, asi que leerlas en modo
formula devuelve los mismos numeros y el resto del libro sobrevive.

CODIGOS SIN MAPEO
-----------------
No todos son un hueco. La tabla homologada trae el P&G comun completo, y el mapeo
solo cubre gastos, asi que los ingresos y las conciliaciones quedan fuera A
PROPOSITO. Reportarlos como aviso en cada corrida es ruido que termina en que nadie
lee los avisos. Aqui se separan:

  - exclusion esperada : ingresos (PL-R-*), no operacionales (PL-NO-*),
                         impuestos (PL-T-*) y conciliaciones (PL-*-900)
  - hueco de verdad    : cualquier codigo de gasto sin linea de tablero

OJO CON COLOMBIA
----------------
Los codigos de gasto sin mapeo son de Colombia (rotulados "- COCO"). No es un olvido:
el puente nunca cubrio a Colombia, y por eso Dashboard_COCO.html le inyecta el
detalle real cuando el puente no la trae (ver README seccion 4.2, funcion
monthlyGastosBridgeData). Antes de agregarlos a MAP_DASHBOARD_GASTOS hay que mirar
esa interaccion: el HTML da prioridad al detalle real, asi que no duplicaria, pero
la decision no es mecanica.

Uso:  python regenerar_gastos.py                   -> vista previa, no escribe
      python regenerar_gastos.py --probar          -> corre sobre una copia y compara
      python regenerar_gastos.py --escribir        -> aplica sobre la base, con respaldo
"""
import os
import shutil
import sys
from collections import OrderedDict
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
RESPALDOS = os.path.join(AQUI, "respaldos")

H_MAPA = "MAP_DASHBOARD_GASTOS"
H_DET = "BD_GASTOS_HOMOLOGADOS"
H_PUENTE = "BD_GASTOS_DASHBOARD_BRIDGE"
H_RESUMEN = "BD_GASTOS_RESUMEN_MENSUAL"

COLS_PUENTE = ["periodo", "pais", "entidad", "moneda", "codigo_comun", "rubro_comun",
               "indicador_dashboard", "detalle_dashboard", "area_dashboard",
               "valor_mes", "valor_ytd", "fuente", "calidad_mapeo", "nota"]
COLS_RESUMEN = ["periodo", "pais", "entidad", "moneda", "indicador_dashboard",
                "detalle_dashboard", "area_dashboard", "valor_mes", "valor_ytd",
                "filas_fuente", "nota"]

TOL = 0.01

# Las dos hojas de destino llevan una nota constante. Se reproducen EXACTAS, con
# el "?" incluido: en la base ese caracter es un signo de interrogacion literal
# (codepoint 0x3F), no una enye mal decodificada al leer. Es el problema de
# codificacion heredado que advierte el README seccion 8. Arreglarlo aqui, de
# paso y sin decirlo, rompería la garantia de que el script no cambia nada.
NOTA_PUENTE = "Puente para redise?ar vistas de Costos & Gastos; no reemplaza P&G oficial."
NOTA_RESUMEN = "Agregado desde BD_GASTOS_DASHBOARD_BRIDGE"


def esperado_fuera(codigo):
    """True si ese codigo NO es un gasto y por lo tanto su ausencia del mapeo es
    la conducta correcta, no un hueco que haya que llenar."""
    c = (codigo or "").upper()
    return (c.startswith("PL-R-") or c.startswith("PL-NO-")
            or c.startswith("PL-T-") or c.endswith("-900"))


def leer_hoja(wb, nombre):
    ws = wb[nombre]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return [], []
    hdr = [str(x).strip() if x is not None else "" for x in filas[0]]
    datos = [OrderedDict(zip(hdr, f)) for f in filas[1:] if any(v is not None for v in f)]
    return hdr, datos


def construir(wb):
    _, mapa_filas = leer_hoja(wb, H_MAPA)
    _, detalle = leer_hoja(wb, H_DET)

    mapa = {}
    for m in mapa_filas:
        cod = str(m.get("codigo_comun") or "").strip()
        if cod:
            mapa[cod] = m

    puente = []
    fuera_esperados, huecos = {}, {}
    for d in detalle:
        cod = str(d.get("codigo_comun") or "").strip()
        m = mapa.get(cod)
        if m is None:
            destino = fuera_esperados if esperado_fuera(cod) else huecos
            clave = (cod, str(d.get("rubro_comun") or "").strip())
            destino[clave] = destino.get(clave, 0) + 1
            continue
        if str(m.get("usar_en_grafica") or "").strip().upper() != "SI":
            continue
        puente.append(OrderedDict([
            ("periodo", d.get("periodo")),
            ("pais", d.get("pais")),
            ("entidad", d.get("entidad")),
            ("moneda", d.get("moneda")),
            ("codigo_comun", cod),
            ("rubro_comun", d.get("rubro_comun")),
            ("indicador_dashboard", m.get("indicador_dashboard")),
            ("detalle_dashboard", m.get("detalle_dashboard")),
            ("area_dashboard", m.get("area_dashboard")),
            ("valor_mes", d.get("movimiento_mes") or 0),
            ("valor_ytd", d.get("saldo_presentacion_ytd") or 0),
            ("fuente", d.get("fuente_archivo")),
            ("calidad_mapeo", m.get("estado_mapeo")),
            ("nota", NOTA_PUENTE),
        ]))

    agr = OrderedDict()
    for p in puente:
        k = (p["periodo"], p["pais"], p["indicador_dashboard"],
             p["detalle_dashboard"], p["area_dashboard"])
        if k not in agr:
            agr[k] = OrderedDict([
                ("periodo", p["periodo"]), ("pais", p["pais"]),
                ("entidad", p["entidad"]), ("moneda", p["moneda"]),
                ("indicador_dashboard", p["indicador_dashboard"]),
                ("detalle_dashboard", p["detalle_dashboard"]),
                ("area_dashboard", p["area_dashboard"]),
                ("valor_mes", 0), ("valor_ytd", 0), ("filas_fuente", 0), ("nota", NOTA_RESUMEN),
            ])
        agr[k]["valor_mes"] += p["valor_mes"]
        agr[k]["valor_ytd"] += p["valor_ytd"]
        agr[k]["filas_fuente"] += 1
    # Redondeo a 2 decimales: la hoja guarda 506,5 donde la suma binaria da
    # 506,4999999999. Sin esto el diff marcaba 28 filas cambiadas por decimas de
    # centavo que no existen en la fuente.
    for v in agr.values():
        v["valor_mes"] = round(v["valor_mes"], 2)
        v["valor_ytd"] = round(v["valor_ytd"], 2)
    # El resumen de la base esta ordenado por periodo, pais, indicador y detalle.
    # Agrupar en orden de aparicion daba el mismo total pero otra secuencia de filas,
    # y el diff contra la hoja anterior salia entero aunque nada hubiera cambiado.
    resumen = sorted(agr.values(), key=lambda x: (str(x["periodo"]), str(x["pais"]),
                                                  str(x["indicador_dashboard"] or ""),
                                                  str(x["detalle_dashboard"] or "")))
    return puente, resumen, fuera_esperados, huecos


def escribir_hoja(wb, nombre, cols, filas):
    """Reemplaza las filas de datos conservando el encabezado."""
    ws = wb[nombre]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, f in enumerate(filas, start=2):
        for j, c in enumerate(cols, start=1):
            ws.cell(i, j).value = f.get(c)
    return len(filas)


def main():
    escribir = "--escribir" in sys.argv
    probar = "--probar" in sys.argv

    destino = BASE
    if probar:
        destino = os.path.join(AQUI, "_prueba_regenerar_gastos.xlsx")
        shutil.copy2(BASE, destino)

    # data_only=False: ver el encabezado del modulo
    wb = openpyxl.load_workbook(destino, data_only=False)

    _, resumen_antes = leer_hoja(wb, H_RESUMEN)
    total_antes = sum(r.get("valor_mes") or 0 for r in resumen_antes)

    puente, resumen, fuera, huecos = construir(wb)
    s_puente = sum(p["valor_mes"] for p in puente)
    s_resumen = sum(r["valor_mes"] for r in resumen)

    print("=" * 78)
    print("REGENERACION DE LA CADENA DE GASTOS")
    print("=" * 78)
    print("  puente  : %4d filas   suma valor_mes = %15.2f" % (len(puente), s_puente))
    print("  resumen : %4d filas   suma valor_mes = %15.2f" % (len(resumen), s_resumen))
    print()
    print("  VALIDACION puente vs resumen: dif %.4f  -> %s"
          % (s_puente - s_resumen, "OK" if abs(s_puente - s_resumen) < TOL else "NO CUADRA"))
    if abs(s_puente - s_resumen) >= TOL:
        print("  ABORTA: no se escribe nada.")
        return 1

    # --- SEGURO DE MONEDA -------------------------------------------------------
    # La cadena suma movimiento_mes sin mirar la moneda. Mientras solo hubo filiales
    # eso ya estaba mal --se sumaban 70.163 soles a 113.311 dolares y el total daba
    # 183.474-- pero pasaba desapercibido porque las magnitudes se parecen. Con
    # Colombia dentro deja de disimular: 656 millones de pesos sumados a dolares.
    # Regla 4 del proyecto: no sumar monedas distintas.
    monedas = {}
    for x in puente:
        m = str(x.get("moneda") or "?")
        monedas[m] = monedas.get(m, 0) + (x.get("valor_mes") or 0)
    if len(monedas) > 1:
        print()
        print("  ABORTA: el puente mezcla monedas y esta hoja se suma sin convertir.")
        for m, v in sorted(monedas.items(), key=lambda kv: -abs(kv[1])):
            print("     %-5s %18s" % (m, "{:,.2f}".format(v).replace(",", ".")))
        print("     El total que saldria (%s) no significa nada: son monedas distintas"
              % "{:,.2f}".format(sum(monedas.values())).replace(",", "."))
        print("     sumadas como si fueran la misma.")
        print()
        print("     Para levantar este seguro hay que decidir la conversion: llevar")
        print("     movimiento_mes a USD con la TRM del pais y del mes (hoja TRM para")
        print("     Colombia, TRM_Peru para Peru) antes de agregar. Eso CAMBIA las cifras")
        print("     de la pestana Gastos, asi que es una decision, no un cargue.")
        return 1

    print()
    print("  PRUEBA DE ACEPTACION (contra lo que hay hoy en la base)")
    print("    antes  : %4d filas   %15.2f" % (len(resumen_antes), total_antes))
    print("    despues: %4d filas   %15.2f" % (len(resumen), s_resumen))
    # El total y el numero de filas no bastan: con ellos cuadrando se puede haber
    # perdido una columna de texto o cambiado el orden de las filas. Se compara celda
    # a celda contra lo que hay hoy en las dos hojas de destino.
    def celdas(nombre, cols, filas):
        return [tuple(f.get(c) for c in cols) for f in filas]

    _, puente_antes = leer_hoja(wb, H_PUENTE)
    dif_res = sum(1 for x, y in zip(celdas(H_RESUMEN, COLS_RESUMEN, resumen_antes),
                                    celdas(H_RESUMEN, COLS_RESUMEN, resumen)) if x != y)
    dif_pue = sum(1 for x, y in zip(celdas(H_PUENTE, COLS_PUENTE, puente_antes),
                                    celdas(H_PUENTE, COLS_PUENTE, puente)) if x != y)
    iguales = (len(resumen_antes) == len(resumen) and abs(total_antes - s_resumen) < TOL
               and dif_res == 0 and dif_pue == 0 and len(puente_antes) == len(puente))
    print("    filas distintas: puente %d · resumen %d" % (dif_pue, dif_res))
    print("    -> %s" % ("IDENTICO celda por celda" if iguales
                         else "*** CAMBIA: revisar antes de aplicar ***"))
    if not iguales:
        for x, y in zip(celdas(H_RESUMEN, COLS_RESUMEN, resumen_antes),
                        celdas(H_RESUMEN, COLS_RESUMEN, resumen)):
            if x != y:
                print("       antes: %s" % str(x)[:96])
                print("       ahora: %s" % str(y)[:96])
                break

    if fuera:
        print()
        print("  Fuera del mapeo por diseno (no son gastos): %d codigos, %d filas"
              % (len(fuera), sum(fuera.values())))
        for (c, rub), n in sorted(fuera.items(), key=lambda x: -x[1]):
            print("     %-12s %-36s %4d filas" % (c, rub[:36], n))

    if huecos:
        print()
        print("  *** GASTOS SIN MAPEO -- revisar si deben entrar a %s ***" % H_MAPA)
        for (c, rub), n in sorted(huecos.items(), key=lambda x: -x[1]):
            print("     %-12s %-36s %4d filas" % (c, rub[:36], n))
        print("     OJO: si dicen '- COCO' son de Colombia. El puente nunca la cubrio y")
        print("     el HTML le inyecta el detalle real (README 4.2). Mirar esa interaccion")
        print("     antes de mapearlos.")

    if not (escribir or probar):
        print()
        print("VISTA PREVIA. Nada se escribio. Usa --probar o --escribir.")
        return 0

    if escribir and not iguales:
        print()
        print("ABORTA: el resultado no reproduce la base actual. Corre --probar y revisa.")
        return 1

    if escribir:
        os.makedirs(RESPALDOS, exist_ok=True)
        sello = datetime.now().strftime("%Y%m%d_%H%M%S")
        resp = os.path.join(RESPALDOS, "BD_MAESTRA_COCO_antes_regenerar_gastos_%s.xlsx" % sello)
        shutil.copy2(BASE, resp)
        print()
        print("  respaldo: %s" % os.path.basename(resp))

    n1 = escribir_hoja(wb, H_PUENTE, COLS_PUENTE, puente)
    n2 = escribir_hoja(wb, H_RESUMEN, COLS_RESUMEN, resumen)
    wb.save(destino)

    print("  escritas: %s %d filas · %s %d filas" % (H_PUENTE, n1, H_RESUMEN, n2))
    print("  archivo : %s" % os.path.basename(destino))
    if probar:
        print()
        print("Era una PRUEBA sobre una copia. La base oficial no se toco.")
    else:
        print()
        print("SIGUIENTE PASO: corre Revisar_Base.bat para confirmar que las capas")
        print("de la base siguen alineadas antes de publicar el cierre.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
