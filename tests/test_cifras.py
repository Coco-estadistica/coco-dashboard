# -*- coding: utf-8 -*-
"""
Las cifras concretas del ultimo cierre publicado.

ESTE ARCHIVO SE ACTUALIZA CADA MES, Y ESO ES EL PUNTO.
Cuando entre agosto, estos tests van a fallar. No los borres ni los ajustes a ojo:
cada cifra que se mueva hay que mirarla y decir por que se movio. Esa revision ES el
cierre. Si una cambia y no sabes por que, ahi hay un error de cargue.

De donde salen: recalculadas desde BD_Indicadores el 28-ago-2026, contrastadas contra
lo que muestra el tablero en pantalla, pais por pais. No se copiaron de la pantalla:
la pantalla se comparo contra estas.
"""
import pytest

from conftest import PAISES

CIERRE = "2026-07"
TOL = 1.0
TOL_PCT = 0.05


# ---------------------------------------------------------------- P&G por pais, USD
# ingresos, EBITDA (utilidad operativa), utilidad neta
PYG = {
    "Colombia":    (207984.38,   8976.08,   7558.10),
    "EE.UU.":      (  6013.45,  -5988.30,  -6111.82),
    "Perú":        (  2796.65,   -882.87,   -910.15),
    "Costa Rica":  (134104.09, 131056.98, 131026.62),
    "Consolidado": (350898.57, 133162.00, 131562.75),
}


@pytest.mark.parametrize("pais", list(PYG))
def test_pyg_del_cierre(V, pais):
    """Las tres cifras que encabezan cualquier lectura: ingresos, EBITDA y resultado."""
    ing, eb, un = PYG[pais]
    assert abs(V("pyg_ingresos_operacionales", CIERRE, pais) - ing) < TOL, "ingresos de " + pais
    assert abs(V("pyg_utilidad_operativa", CIERRE, pais) - eb) < TOL, "EBITDA de " + pais
    assert abs(V("pyg_utilidad_neta", CIERRE, pais) - un) < TOL, "utilidad neta de " + pais


def test_margen_ebitda_consolidado(V):
    """37,9% es la cifra que encabeza el tablero. Si se mueve, se mueve la lectura."""
    ing = V("pyg_ingresos_operacionales", CIERRE, "Consolidado")
    eb = V("pyg_utilidad_operativa", CIERRE, "Consolidado")
    assert abs(eb / ing * 100 - 37.9) < TOL_PCT


def test_margen_bruto_consolidado(V):
    """71,9%. Se derivaba de una serie de gestion que llevaba siete meses sin
    alimentarse; desde el 28-ago sale del P&G oficial (ingresos - costo de ventas)."""
    ing = V("pyg_ingresos_operacionales", CIERRE, "Consolidado")
    cv = V("pyg_costo_ventas", CIERRE, "Consolidado")
    assert abs((ing - cv) / ing * 100 - 71.9) < TOL_PCT


def test_egresos_y_operativos_difieren_en_los_financieros(V):
    """Dos tarjetas mostraban el MISMO numero con nombres distintos, porque los dos
    codigos apuntaban a la misma funcion. Operativos excluye financieros."""
    cv = V("pyg_costo_ventas", CIERRE, "Consolidado")
    adm = V("pyg_gasto_administracion", CIERRE, "Consolidado")
    ven = V("pyg_gasto_ventas", CIERRE, "Consolidado")
    pro = V("pyg_gasto_proyectos", CIERRE, "Consolidado")
    fin = V("pyg_gastos_financieros", CIERRE, "Consolidado")
    egresos, operativos = cv + adm + ven + pro + fin, cv + adm + ven + pro
    assert abs(egresos - 219567) < TOL
    assert abs(operativos - 217738) < TOL
    assert abs((egresos - operativos) - fin) < 0.01, "la diferencia debe ser exactamente los financieros"


# ---------------------------------------------------------------- MRR
MRR = {"2026-01": 592860893.20, "2026-02": 695884029.36, "2026-03": 734076969.85,
       "2026-04": 754046327.50, "2026-05": 774122622.75, "2026-06": 759860184.11,
       "2026-07": 703109437.70}


@pytest.mark.parametrize("per", sorted(MRR))
def test_serie_de_mrr(V, per):
    """Serie restablecida el 27-ago con la definicion VENCIDA(mes anterior) +
    CORRIENTE(mes en curso), y la Caja Costa Rica fuera del recurrente."""
    assert abs(V("mrr_final", per, "Colombia", seg="") - MRR[per]) < TOL


def test_el_puente_de_mrr_cierra(V):
    """Un puente de MRR que no cierra es lo primero que revisa un inversionista de SaaS.
    Este no cerraba: faltaban 80,8 M en julio porque el one-off de Costa Rica estaba
    contado como venta nueva."""
    jun = V("mrr_final", "2026-06", "Colombia", seg="")
    jul = V("mrr_final", "2026-07", "Colombia", seg="")
    mov = sum(V(c, "2026-07", "Colombia", seg="") or 0
              for c in ("new_mrr", "expansion_mrr", "contraction_mrr", "churned_mrr"))
    assert abs((jun + mov) - jul) < TOL, "el puente no cierra: %.2f vs %.2f" % (jun + mov, jul)


def test_el_one_off_no_entra_al_mrr(V):
    """494.843.760 de la Caja Costa Rica: van a ingresos no recurrentes, no a MRR."""
    assert abs(V("ingresos_no_recurrentes", "2026-07", "Colombia", seg="") - 494843760) < TOL
    assert abs(V("new_mrr", "2026-07", "Colombia", seg="") - 9249060) < TOL


# ---------------------------------------------------------------- cartera
def test_cartera_reconvierte_al_usd_de_origen(V, trm):
    """Se cargo en su dia con la TRM de junio (3.505) creyendo que era la de julio, y
    quedo 7,2% inflada. Estas cuatro cifras tienen que volver a los USD del archivo."""
    fuente = {"cartera_total": 469850, "cartera_vencida": 195494,
              "cartera_sana": 274355, "cartera_cobro_legal": 16342}
    t = float(trm[CIERRE])
    for code, usd in fuente.items():
        cop = V(code, CIERRE, "Colombia", seg="")
        assert abs(cop / t - usd) < TOL, "%s: %.2f COP / %.2f = %.2f, se esperaba %d" % (
            code, cop, t, cop / t, usd)


def test_la_morosidad_es_la_que_entrega_cartera(V):
    """77,3%, NO vencida/total (41,6%). Es una decision de negocio del 27-ago: el
    tablero no la calcula. Si alguien la 'corrige' a 41,6%, este test lo atrapa."""
    assert abs(V("indice_morosidad", CIERRE, "Colombia", seg="") - 77.3) < TOL_PCT


# ---------------------------------------------------------------- crecimiento
def test_crecimiento_organico_de_colombia(V, trm):
    """+38,3% en USD comparando ene-jul contra ene-jul. El tablero mostraba +119%
    porque comparaba Colombia-2025 contra Consolidado-2026."""
    meses = ["2026-%02d" % i for i in range(1, 8)]
    c26 = sum(V("pyg_ingresos_operacionales", p, "Colombia") or 0 for p in meses)
    c25 = 0.0
    for p in meses:
        cop = V("ingresos_totales", p.replace("2026", "2025"), "Colombia", seg="")
        t = trm.get(p.replace("2026", "2025"))
        if cop and t:
            c25 += cop / float(t)
    assert abs((c26 / c25 - 1) * 100 - 38.3) < 0.2, "organico: %.1f%%" % ((c26 / c25 - 1) * 100)
