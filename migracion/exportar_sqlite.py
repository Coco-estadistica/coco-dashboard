# -*- coding: utf-8 -*-
"""
Exporta BD_MAESTRA_COCO.xlsx a coco.db (SQLite) y valida las reglas del negocio.

QUE ES Y QUE NO ES
------------------
Es un ESPEJO, no un reemplazo. Excel sigue siendo la fuente; el tablero no se toca
ni se entera. Lo que se gana:

  - Reglas escritas como codigo ejecutable, no como costumbre. Hoy "el consolidado es
    la suma de los paises" vive en la cabeza de quien carga; aqui es una consulta que
    falla sola.
  - Consultas de verdad sobre 1.849 hechos, sin abrir Excel.
  - Un artefacto que git PUEDE comparar. El .xlsx es binario, y por eso la regla del
    proyecto es "no edites mientras el otro edita". El volcado .sql se lee y se fusiona.

NO se escribe nada en el .xlsx. Se abre solo para leer, nunca se guarda, asi que las
3.882 formulas del libro no corren ningun riesgo por este script.

QUE VALIDA
----------
Cada control salio de un error real de este proyecto, no de un manual:

  V1  consolidado = suma de los 4 paises        (se rompio al corregir julio de Peru)
  V2  utilidad = ingresos - gastos, por pais    (quedo desalineada en el mismo cambio)
  V3  la cartera concilia: sana + vencida       (revisado 03-sep-2026: antes comparaba
      = cartera_total                            el indice_morosidad cargado contra
                                                 vencida/total, pero eran definiciones
                                                 distintas y el hallazgo era permanente.
                                                 Desde que el tablero calcula la
                                                 morosidad, lo que importa es que el
                                                 denominador cuadre)
  V4  la composicion del NPS suma 100           (se sumaban los tres detalles y daba
                                                 100 todos los meses)
  V5  sin llaves duplicadas                     (misma llave repetida pasa inadvertida
                                                 y se suma dos veces)
  V6  sin periodos futuros con escenario Real

Uso:  python exportar_sqlite.py            -> exporta y valida
      python exportar_sqlite.py --solo-validar
      python exportar_sqlite.py --volcado   -> ademas escribe coco.sql (texto, para git)
"""
import os
import sqlite3
import sys

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
DB = os.path.join(AQUI, "coco.db")
VOLCADO = os.path.join(AQUI, "coco.sql")

TOL = 0.02          # suma binaria de decimales
TOL_PUBLICADO = 1.01  # el bloque "Acumulado H1" viene publicado redondeado a unidades,
                      # asi que comparar contra una suma al centavo siempre difiere por
                      # menos de 1. Eso no es un error: es como se publico el estado.
PAISES_CONSOL = ("EE.UU.", "Perú", "Costa Rica", "Colombia")

# hoja -> (tabla, fila del encabezado)
HOJAS = {
    "BD_Indicadores": ("indicadores", 3),
    "Diccionario": ("diccionario", None),
    "TRM": ("trm", None),
    "TRM_Peru": ("trm_peru", None),
    "BD_GASTOS_HOMOLOGADOS": ("gastos_homologados", 1),
    "BD_GASTOS_DASHBOARD_BRIDGE": ("gastos_puente", 1),
    "BD_GASTOS_RESUMEN_MENSUAL": ("gastos_resumen", 1),
    "MAP_DASHBOARD_GASTOS": ("gastos_mapeo", 1),
}


def norm_col(s):
    s = str(s or "").strip().lower()
    out = "".join(c if (c.isalnum() or c == "_") else "_" for c in s)
    while "__" in out:
        out = out.replace("__", "_")
    out = out.strip("_")
    return out or "col"


def buscar_encabezado(filas, pista):
    if pista:
        return pista - 1
    for i, f in enumerate(filas[:8]):
        if f and sum(1 for x in f if x is not None) >= 2:
            return i
    return 0


def exportar():
    # data_only=True es seguro AQUI porque el libro se abre solo para leer y nunca se
    # guarda. El peligro de perder formulas aparece al llamar wb.save(), cosa que este
    # script no hace en ningun camino.
    wb = openpyxl.load_workbook(BASE, data_only=True)
    if os.path.exists(DB):
        os.remove(DB)
    con = sqlite3.connect(DB)
    cur = con.cursor()
    resumen = []

    for hoja, (tabla, hdr) in HOJAS.items():
        if hoja not in wb.sheetnames:
            resumen.append((tabla, 0, "hoja ausente"))
            continue
        filas = list(wb[hoja].iter_rows(values_only=True))
        if not filas:
            resumen.append((tabla, 0, "vacia"))
            continue
        h = buscar_encabezado(filas, hdr)
        cols, vistos = [], {}
        for x in filas[h]:
            c = norm_col(x)
            vistos[c] = vistos.get(c, 0) + 1
            cols.append(c if vistos[c] == 1 else "%s_%d" % (c, vistos[c]))
        cur.execute('CREATE TABLE "%s" (%s)' % (tabla, ",".join('"%s"' % c for c in cols)))
        datos = [f[:len(cols)] + (None,) * max(0, len(cols) - len(f))
                 for f in filas[h + 1:] if any(v is not None for v in f)]
        cur.executemany('INSERT INTO "%s" VALUES (%s)' % (tabla, ",".join("?" * len(cols))),
                        [tuple(None if v is None else (v if isinstance(v, (int, float)) else str(v))
                               for v in d) for d in datos])
        resumen.append((tabla, len(datos), ",".join(cols[:4]) + "…"))

    # vistas que calculan en vez de guardar
    cur.executescript("""
    CREATE VIEW v_consol_suma AS
      SELECT periodo, codigo_indicador, segmento, ROUND(SUM(valor),2) AS suma_paises
      FROM indicadores
      WHERE pais IN ('EE.UU.','Perú','Costa Rica','Colombia') AND escenario='Real'
      GROUP BY periodo, codigo_indicador, segmento;

    CREATE VIEW v_consol_guardado AS
      SELECT periodo, codigo_indicador, segmento, ROUND(SUM(valor),2) AS guardado
      FROM indicadores WHERE pais='Consolidado' AND escenario='Real'
      GROUP BY periodo, codigo_indicador, segmento;
    """)
    con.commit()
    return con, resumen


def validar(con):
    cur = con.cursor()
    fallos = []

    def chk(nombre, sql, fmt):
        rows = cur.execute(sql).fetchall()
        if rows:
            fallos.append((nombre, len(rows), [fmt(r) for r in rows[:4]]))
        return len(rows)

    n1 = chk("V1 consolidado = suma de paises",
             """SELECT s.periodo, s.codigo_indicador, s.suma_paises, g.guardado
                FROM v_consol_suma s JOIN v_consol_guardado g
                  ON s.periodo=g.periodo AND s.codigo_indicador=g.codigo_indicador
                 AND IFNULL(s.segmento,'')=IFNULL(g.segmento,'')
                WHERE ABS(s.suma_paises-g.guardado) >
                      CASE WHEN IFNULL(s.segmento,'') LIKE 'Acumulado%%' THEN %s ELSE %s END"""
             % (TOL_PUBLICADO, TOL),
             lambda r: "%s %s: paises %.2f vs guardado %.2f" % r)

    n2 = chk("V2 utilidad = ingresos - gastos",
             """WITH p AS (SELECT periodo,pais,segmento,
                    SUM(CASE codigo_indicador WHEN 'pyg_ingresos_operacionales' THEN valor ELSE 0 END)
                  - SUM(CASE WHEN codigo_indicador IN ('pyg_costo_ventas','pyg_gasto_administracion',
                        'pyg_gasto_ventas','pyg_gasto_proyectos','pyg_gastos_financieros',
                        'pyg_reclasif_honorarios_peru') THEN valor ELSE 0 END)
                  + SUM(CASE codigo_indicador WHEN 'pyg_ingresos_no_operacionales' THEN valor ELSE 0 END) AS calc,
                    SUM(CASE codigo_indicador WHEN 'pyg_utilidad_neta' THEN valor ELSE 0 END) AS guard,
                    SUM(CASE codigo_indicador WHEN 'pyg_utilidad_neta' THEN 1 ELSE 0 END) AS tiene
                  FROM indicadores WHERE escenario='Real' GROUP BY periodo,pais,segmento)
                SELECT periodo,pais,ROUND(calc,2),ROUND(guard,2) FROM p
                WHERE tiene>0 AND ABS(calc-guard) >
                      CASE WHEN IFNULL(segmento,'') LIKE 'Acumulado%%' THEN %s ELSE %s END"""
             % (TOL_PUBLICADO, TOL),
             lambda r: "%s %s: calculada %.2f vs guardada %.2f" % r)

    # V3 cambio de sentido el 03-sep-2026. Antes comparaba el indice_morosidad cargado
    # contra vencida/total y lo marcaba como descuadre; pero eran dos definiciones
    # distintas, asi que el hallazgo era permanente y no significaba nada. Conectado al
    # boton de subir habria frenado todas las subidas por un falso positivo, y en dos
    # semanas alguien lo habria apagado -- que es como mueren los controles.
    #
    # Ahora el tablero calcula la morosidad como vencida/cartera_total, y lo que si hay
    # que vigilar es el denominador: cuando sana+vencida no da cartera_total, el
    # porcentaje cambia segun cual se use. En junio de 2026: 27,81% contra 32,18%.
    n3 = chk("V3 la cartera concilia (sana + vencida = total)",
             """WITH c AS (SELECT periodo,
                    SUM(CASE codigo_indicador WHEN 'cartera_sana' THEN valor END) AS san,
                    SUM(CASE codigo_indicador WHEN 'cartera_vencida' THEN valor END) AS ven,
                    SUM(CASE codigo_indicador WHEN 'cartera_total' THEN valor END) AS tot
                  FROM indicadores WHERE escenario='Real' GROUP BY periodo)
                SELECT periodo, ROUND(san+ven,0), ROUND(tot,0),
                       ROUND(ven*100.0/tot,2), ROUND(ven*100.0/(san+ven),2)
                FROM c
                WHERE san IS NOT NULL AND ven IS NOT NULL AND tot>0
                  AND ABS(san+ven-tot) > MAX(1, tot*0.001)""",
             lambda r: "%s: sana+vencida %.0f vs total %.0f -> morosidad %.2f%% sobre total, %.2f%% sobre la suma" % r)

    n4 = chk("V4 la composicion del NPS suma 100",
             """SELECT periodo, ROUND(SUM(valor),2) FROM indicadores
                WHERE codigo_indicador='nps' AND escenario='Real'
                GROUP BY periodo HAVING ABS(SUM(valor)-100) > 0.5""",
             lambda r: "%s: los detalles suman %.2f, no 100" % r)

    # La llave de un hecho incluye el escenario: Real y Presupuesto conviven en el mismo
    # mes para el mismo codigo y eso es correcto, no una duplicacion. Sin ese campo el
    # control marcaba las 6 filas de ingresos_totales de Colombia, que estan bien.
    n5 = chk("V5 sin llaves duplicadas",
             """SELECT periodo,pais,codigo_indicador,IFNULL(escenario,''),
                       IFNULL(segmento,''),IFNULL(detalle,''),COUNT(*)
                FROM indicadores GROUP BY 1,2,3,4,5,6 HAVING COUNT(*)>1""",
             lambda r: "%s %s %s [%s/%s/%s] x%d" % r)

    n6 = chk("V6 sin futuro con escenario Real",
             """SELECT periodo, COUNT(*) FROM indicadores
                WHERE escenario='Real' AND periodo > '2026-07' GROUP BY periodo""",
             lambda r: "%s: %d filas Real en un periodo futuro" % r)

    return fallos, (n1, n2, n3, n4, n5, n6)


def main():
    solo = "--solo-validar" in sys.argv
    if solo and os.path.exists(DB):
        con = sqlite3.connect(DB)
        resumen = []
    else:
        con, resumen = exportar()

    print("=" * 78)
    print("BD_MAESTRA_COCO.xlsx  ->  coco.db")
    print("=" * 78)
    for tabla, n, cols in resumen:
        print("  %-22s %6d filas   %s" % (tabla, n, cols[:44]))

    fallos, cuentas = validar(con)
    print()
    print("  VALIDACIONES")
    print("  " + "-" * 74)
    nombres = ["V1 consolidado = suma de paises", "V2 utilidad = ingresos - gastos",
               "V3 la cartera concilia (sana + vencida = total)",
               "V4 la composicion del NPS suma 100",
               "V5 sin llaves duplicadas", "V6 sin futuro con escenario Real"]
    dic = {f[0]: f for f in fallos}
    for nom in nombres:
        f = dic.get(nom)
        print("  %-46s %s" % (nom, "OK" if not f else ("%d hallazgo(s)" % f[1])))
        if f:
            for d in f[2]:
                print("       %s" % d)
            if f[1] > len(f[2]):
                print("       … y %d mas" % (f[1] - len(f[2])))
    # Esta lista repite los titulos que ya estan en validar(), y se emparejan por texto
    # exacto. Al renombrar V3 el resumen siguio imprimiendo "OK" mientras el control
    # tenia un hallazgo: un control que miente es peor que no tenerlo. Si algun dia
    # vuelven a desalinearse, que se vea.
    huerfanos = [n for n in dic if n not in nombres]
    if huerfanos:
        print()
        print("  AVISO: hay controles cuyo titulo no coincide con esta lista y por eso")
        print("         no se resumieron arriba. Corrige los nombres:")
        for n in huerfanos:
            print("           - %s (%d hallazgo(s))" % (n, dic[n][1]))

    if "--volcado" in sys.argv:
        with open(VOLCADO, "w", encoding="utf-8") as fh:
            for linea in con.iterdump():
                fh.write(linea + "\n")
        print()
        print("  volcado de texto: %s (%.1f MB) -- este si lo puede comparar git"
              % (os.path.basename(VOLCADO), os.path.getsize(VOLCADO) / 1e6))

    con.close()
    print()
    print("  base: %s (%.0f KB)" % (os.path.basename(DB), os.path.getsize(DB) / 1024))
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())
