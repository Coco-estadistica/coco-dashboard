# -*- coding: utf-8 -*-
"""
Reglas que NO dependen del mes. Siguen valiendo en agosto, en septiembre y el ano que
viene: no hay que tocarlas cuando entre un cierre nuevo.

Si una de estas falla, no es que las cifras cambiaron -- es que algo se rompio.

Cada prueba dice, en su docstring, QUE incidente real previene. Todas salen de cosas
que ya pasaron en este proyecto y se descubrieron tarde.
"""
import pytest

from conftest import CONSOL_SEG, H1_SEG, PAISES

TOL = 1.0          # un peso/dolar de tolerancia: la base guarda 2 decimales


# --------------------------------------------------------------- estructura del P&G
def test_utilidad_bruta_es_ingresos_menos_costo(V, meses_pyg):
    """Utilidad bruta = ingresos - costo de ventas, en cada pais y cada mes."""
    malos = []
    for pais in PAISES + ["Consolidado"]:
        for per in meses_pyg:
            ing = V("pyg_ingresos_operacionales", per, pais)
            cv = V("pyg_costo_ventas", per, pais)
            ub = V("pyg_utilidad_bruta", per, pais)
            if None in (ing, cv, ub):
                continue
            if abs((ing - cv) - ub) > TOL:
                malos.append("%s %s: %.2f vs %.2f" % (pais, per, ing - cv, ub))
    assert not malos, "la utilidad bruta no se deriva: " + "; ".join(malos[:5])


def test_utilidad_neta_es_operativa_mas_no_operacionales(V, meses_pyg):
    """Utilidad neta = utilidad operativa + ingresos no op. - gastos financieros."""
    malos = []
    for pais in PAISES + ["Consolidado"]:
        for per in meses_pyg:
            uo = V("pyg_utilidad_operativa", per, pais)
            un = V("pyg_utilidad_neta", per, pais)
            if uo is None or un is None:
                continue
            ino = V("pyg_ingresos_no_operacionales", per, pais) or 0
            gfi = V("pyg_gastos_financieros", per, pais) or 0
            if abs((uo + ino - gfi) - un) > TOL:
                malos.append("%s %s: %.2f vs %.2f" % (pais, per, uo + ino - gfi, un))
    assert not malos, "la utilidad neta no se deriva: " + "; ".join(malos[:5])


def test_consolidado_es_la_suma_de_los_cuatro_paises(V, meses_pyg):
    """El consolidado NO se deriva: tiene fila propia guardada. Si se corrige un pais y
    nadie recompone, deja de cuadrar con sus propias columnas y nada avisa.

    Previene: al corregir julio de Peru, la fila Consolidado quedo desalineada y se
    descubrio por casualidad (commit 40e4825).
    """
    CODES = ["pyg_ingresos_operacionales", "pyg_costo_ventas", "pyg_gasto_administracion",
             "pyg_gasto_ventas", "pyg_gasto_proyectos", "pyg_gastos_financieros",
             "pyg_ingresos_no_operacionales", "pyg_utilidad_neta"]
    malos = []
    for code in CODES:
        for per in meses_pyg:
            cons = V(code, per, "Consolidado")
            if cons is None:
                continue
            suma = sum(V(code, per, p) or 0 for p in PAISES)
            if abs(suma - cons) > TOL:
                malos.append("%s %s: paises %.2f vs consolidado %.2f" % (code, per, suma, cons))
    assert not malos, "el consolidado no cuadra: " + "; ".join(malos[:5])


# --------------------------------------------------------------- higiene de la base
def test_sin_llaves_duplicadas(filas):
    """Dos filas con la misma llave se suman sin que nadie lo pida."""
    from collections import Counter
    c = Counter((str(f.get("codigo_indicador") or "").strip(), str(f.get("periodo")),
                 str(f.get("pais") or ""), str(f.get("escenario") or ""),
                 str(f.get("segmento") or ""), str(f.get("detalle") or ""))
                for f in filas)
    reps = [k for k, n in c.items() if n > 1]
    assert not reps, "%d llaves repetidas, p.ej. %s" % (len(reps), reps[0])


def test_sin_cifras_reales_en_meses_futuros(filas, ultimo_mes):
    """El presupuesto llega a diciembre; lo Real, no. Una cifra Real en un mes que aun
    no cerro es un cargue en el periodo equivocado."""
    futuras = [f for f in filas
               if str(f.get("escenario") or "") == "Real"
               and str(f.get("periodo")) > ultimo_mes
               and f.get("valor") is not None]
    assert not futuras, "%d filas Reales despues de %s: %s" % (
        len(futuras), ultimo_mes,
        sorted({str(f.get("periodo")) + " " + str(f.get("codigo_indicador")) for f in futuras})[:3])


def test_los_gastos_del_pyg_se_guardan_en_positivo(filas):
    """Mezclar signos hace que una suma reste. La convencion es: gastos en positivo, y
    el signo lo pone quien presenta."""
    malos = [(str(f.get("codigo_indicador")), str(f.get("periodo")), str(f.get("pais")))
             for f in filas
             if str(f.get("codigo_indicador") or "").strip().startswith(("pyg_gasto", "pyg_costo"))
             and f.get("valor") is not None and float(f.get("valor")) < 0]
    assert not malos, "gastos con signo negativo: %s" % (malos[:5],)


def test_el_acumulado_del_semestre_no_se_mezcla_con_el_mes(V):
    """Las filas del cierre semestral viven DENTRO de 2026-06, junto a las del mes. Una
    suma que no filtre por segmento devuelve casi siete veces el ingreso real.

    Este test no exige que la trampa desaparezca -- exige que siga siendo detectable,
    para que quien escriba una consulta nueva sepa que tiene que filtrar.
    """
    mes = V("pyg_ingresos_operacionales", "2026-06", "Consolidado", seg=CONSOL_SEG)
    todo = V("pyg_ingresos_operacionales", "2026-06", "Consolidado", seg=None)
    if mes is None or todo is None:
        pytest.skip("no hay junio 2026 en la base")
    assert todo > mes * 2, (
        "el bloque '%s' ya no infla junio: o se movio a su propio periodo --y entonces "
        "hay que borrar este test-- o se perdio" % H1_SEG)


# --------------------------------------------------------------- moneda
def test_la_trm_derivada_reproduce_ingresos_cop_sobre_usd(V, trm, meses_pyg):
    """trm_cop_usd se declara como ingresos COP / ingresos USD de la consolidacion. Si
    no reproduce, alguien cambio una de las dos series sin recalcular la otra.

    Previene: la cartera de julio se cargo con la TRM de junio (3.505 en vez de 3.269) y
    quedo 7,2% inflada durante ocho dias (commit 564503b).
    """
    malos = []
    for per in meses_pyg:
        cop = V("ingresos_totales", per, "Colombia", seg="")
        usd = V("pyg_ingresos_operacionales", per, "Colombia")
        if not cop or not usd or per not in trm or not trm[per]:
            continue
        if abs(cop / usd - float(trm[per])) > 1:
            malos.append("%s: implicita %.2f vs TRM %.2f" % (per, cop / usd, float(trm[per])))
    assert not malos, "la TRM no reproduce: " + "; ".join(malos[:5])


def test_cada_mes_con_pyg_tiene_su_trm(trm, meses_pyg):
    """Sin TRM del mes, el tablero convierte con una tasa de respaldo y nadie se entera."""
    faltan = [p for p in meses_pyg if p not in trm or not trm[p]]
    assert not faltan, "meses con P&G y sin TRM: %s" % faltan


# --------------------------------------------------------------- lo que el tablero promete
def test_los_codigos_mostrados_existen_en_el_diccionario(libro, filas):
    """Un codigo con datos y sin ficha sale en el Catalogo sin nombre ni formula."""
    hojas = list(libro["Diccionario"].iter_rows(values_only=True))
    cab = [str(v or "").strip() for v in hojas[2]]
    i = cab.index("Código indicador")
    fichas = {str(f[i]).strip() for f in hojas[3:] if f and f[i]}
    usados = {str(f.get("codigo_indicador") or "").strip() for f in filas}
    huerfanos = sorted(c for c in usados - fichas if c and not c.startswith("pyg_"))
    assert not huerfanos, "codigos con datos y sin ficha: %s" % huerfanos[:8]
