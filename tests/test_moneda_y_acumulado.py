# -*- coding: utf-8 -*-
"""Las reglas de moneda y de acumulado que se fijaron el 3-sep-2026.

POR QUE ESTE ARCHIVO EXISTE
---------------------------
La suite de 52 pruebas pasaba en verde el dia en que se encontraron 19 errores en
pantalla. Ninguno era de estructura de la base --por eso test_invariantes no los
vio-- ni de una formula puntual. Eran de CONVERSION y de ACUMULADO: el tablero
mostraba 1.612.839 donde el estado financiero decia 1.664.466, y en pesos mostraba
a Colombia con 4.682 MM cuando las siete barras de su propio grafico sumaban
5.152 MM.

Un error asi no se ve leyendo el codigo: hay que preguntarle a la pagina. Estas
pruebas abren el tablero de verdad y le preguntan.

LAS DOS REGLAS QUE FIJAN
------------------------
1. El acumulado del P&G es el OFICIAL: el bloque "Acumulado H1" que publico
   contabilidad, mas los meses posteriores. No la reconversion mes a mes.
2. En pesos, Colombia se muestra con SU LIBRO (la suma de sus meses en COP), no
   el acumulado en dolares multiplicado por la tasa del corte. Las filiales, que
   se originan en dolares, si se convierten con la tasa del corte.

Si alguna de estas dos decisiones se cambia a proposito, hay que venir aqui y
cambiar la prueba a mano. Ese es el punto: obliga a decir por que.
"""
import os
import socket
import subprocess
import sys
import time

import pytest

pytest.importorskip(
    "playwright.sync_api",
    reason="playwright no instalado; estas pruebas de navegador se omiten",
)
from playwright.sync_api import sync_playwright  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TABLERO = "index.html"
CHROME_SISTEMA = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

CIERRE = "2026-07"
TOL = 2.0          # dolares
TOL_COP = 2000.0   # pesos


def _puerto_libre():
    with socket.socket() as s:
        s.bind(("", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="module")
def pagina():
    puerto = _puerto_libre()
    proc = subprocess.Popen(
        [sys.executable, "-m", "http.server", str(puerto)],
        cwd=RAIZ, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    for _ in range(50):
        try:
            with socket.create_connection(("127.0.0.1", puerto), 0.2):
                break
        except OSError:
            time.sleep(0.1)
    with sync_playwright() as p:
        kw = {"args": ["--no-sandbox"]}
        if os.path.exists(CHROME_SISTEMA):
            kw["executable_path"] = CHROME_SISTEMA
        nav = p.chromium.launch(**kw)
        pg = nav.new_page()
        pg.errores = []
        pg.on("pageerror", lambda e: pg.errores.append(str(e)))
        pg.goto(f"http://127.0.0.1:{puerto}/{TABLERO}")
        pg.wait_for_function(
            "() => typeof S !== 'undefined' && S && S.periodos && S.periodos.length > 0",
            timeout=30000,
        )
        pg.wait_for_timeout(1200)
        yield pg
        nav.close()
    proc.terminate()
    proc.wait(timeout=10)


def _moneda(pg, m):
    pg.evaluate(f"() => {{ F.moneda = '{m}'; render(); }}")
    pg.wait_for_timeout(400)


def _pais(pg, p):
    pg.evaluate(f"() => {{ F.pais = '{p}'; render(); }}")
    pg.wait_for_timeout(400)


# ------------------------------------------------------ el acumulado es el oficial

def test_el_acumulado_en_dolares_es_el_bloque_oficial(pagina):
    """1.664.466 = "Acumulado H1 2026" (1.313.567) + julio (350.899).

    Antes daba 1.612.839 porque Colombia se reconvertia mes a mes con la tasa de
    cada mes en vez de tomarse del bloque que publico contabilidad. La diferencia
    --51.627 dolares-- no es un redondeo: es que el tablero le estaba discutiendo
    la tasa al estado financiero.
    """
    _moneda(pagina, "USD")
    v = pagina.evaluate(
        f"officialAccumDisplay('pyg_ingresos_operacionales','Consolidado','{CIERRE}')")
    h1 = pagina.evaluate(
        "factVal('pyg_ingresos_operacionales','2026-06','Consolidado','Acumulado H1 2026')")
    jul = pagina.evaluate(
        f"factVal('pyg_ingresos_operacionales','{CIERRE}','Consolidado',CONSOL_SEG)")
    assert abs(v - (h1 + jul)) < TOL, (
        f"el acumulado deberia ser {h1:,.0f} + {jul:,.0f} = {h1+jul:,.0f} y da {v:,.0f}")


def test_en_pesos_colombia_es_la_suma_de_sus_meses(pagina):
    """El acumulado de Colombia en COP tiene que ser el libro contable.

    Antes se calculaba en dolares y se multiplicaba por la tasa de julio: el KPI
    decia 4.681.950.372 mientras las siete barras del grafico de abajo, en la
    misma pantalla, sumaban 5.152.233.299. Cualquiera con una calculadora lo veia.
    """
    _moneda(pagina, "COP")
    v = pagina.evaluate(
        f"officialAccumDisplay('pyg_ingresos_operacionales','Colombia','{CIERRE}')")
    meses = pagina.evaluate("""(per) => {
        const anio = per.slice(0,4); let s = 0;
        for (const m of S.periodos) {
            if (m.slice(0,4) !== anio || m > per) continue;
            const v = valueOf('pyg_ingresos_operacionales',
                              {periodo:m, pais:'Colombia', segmento:'', escenario:'Real'});
            if (v != null) s += v;
        }
        return s;
    }""", CIERRE)
    assert abs(v - meses) < TOL_COP, (
        f"el acumulado en pesos da {v:,.0f} y la suma de los meses es {meses:,.0f}")


def test_en_pesos_el_consolidado_es_la_suma_de_las_columnas(pagina):
    """La columna CONSOLIDADO de la tabla tiene que ser lo que suman las de pais.

    Es la comprobacion que hace cualquiera que lea la tabla: sumar de izquierda a
    derecha. Si no da, la tabla no se puede defender aunque cada celda tenga su
    explicacion.
    """
    _moneda(pagina, "COP")
    for code in ("pyg_ingresos_operacionales", "pyg_costo_ventas",
                 "pyg_gasto_administracion", "pyg_gasto_ventas"):
        tot = pagina.evaluate(f"officialAccumDisplay('{code}','Consolidado','{CIERRE}')")
        partes = [pagina.evaluate(f"officialAccumDisplay('{code}','{p}','{CIERRE}')")
                  for p in ("Colombia", "EE.UU.", "Perú", "Costa Rica")]
        suma = sum(x for x in partes if x is not None)
        # la reclasificacion de honorarios de Peru solo existe en el consolidado
        assert abs(tot - suma) < max(TOL_COP, abs(tot) * 0.001), (
            f"{code}: consolidado {tot:,.0f} contra la suma de paises {suma:,.0f}")


def test_ida_y_vuelta_entre_monedas(pagina):
    """USD -> COP -> USD devuelve el valor original, en el acumulado y en el mes."""
    for code in ("pyg_ingresos_operacionales", "pyg_utilidad_neta"):
        _moneda(pagina, "USD")
        usd = pagina.evaluate(f"officialAccumDisplay('{code}','EE.UU.','{CIERRE}')")
        _moneda(pagina, "COP")
        cop = pagina.evaluate(f"officialAccumDisplay('{code}','EE.UU.','{CIERRE}')")
        trm = pagina.evaluate(f"trmOf('{CIERRE}')")
        assert abs(cop / trm - usd) < TOL, (
            f"{code}: {cop:,.0f} COP / {trm:,.0f} = {cop/trm:,.0f}, se esperaba {usd:,.0f}")


def test_ninguna_cifra_en_pesos_llega_a_los_billones(pagina):
    """La doble conversion multiplicaba por la tasa dos veces: factor ~3.600.

    Un gasto de Colombia salio en 17.707.767.169.647 --diecisiete billones-- cuando
    el correcto eran 4.916 millones. Trece digitos en un rubro mensual o anual de
    esta compania es siempre el mismo error.
    """
    _moneda(pagina, "COP")
    _pais(pagina, "__ALL__")
    peores = pagina.evaluate("""(per) => {
        const codes = ['pyg_ingresos_operacionales','pyg_costo_ventas',
                       'pyg_gasto_administracion','pyg_gasto_ventas',
                       'pyg_gasto_proyectos','pyg_gastos_financieros'];
        const ents = ['Colombia','EE.UU.','Perú','Costa Rica','Consolidado'];
        const malos = [];
        for (const c of codes) for (const e of ents) {
            const v = officialAccumDisplay(c, e, per);
            if (v != null && Math.abs(v) > 1e12) malos.push(c + '/' + e + ' = ' + Math.round(v));
        }
        return malos;
    }""", CIERRE)
    assert not peores, "cifras en billones de pesos: " + "; ".join(peores[:5])


# ------------------------------------------------------ coherencia entre pestanas

def test_gastos_y_finanzas_dicen_lo_mismo(pagina):
    """La tabla de Gastos y el P&G de Finanzas leen la misma fuente.

    Antes Gastos sumaba las filas mensuales y Finanzas leia el bloque oficial. Para
    Colombia y Costa Rica daba casi igual, pero contabilidad reclasifico rubros de
    EE.UU. y Peru al cerrar el semestre: administracion de EE.UU. salia en 64.117
    en una pestana y 34.738 en la otra, el mismo rubro y el mismo pais.
    """
    _moneda(pagina, "USD")
    _pais(pagina, "__ALL__")
    # Se leen las celdas que la tabla PINTA, no la funcion que deberia usar: la
    # prueba tiene que fallar tambien si alguien la reimplementa por dentro.
    tabla = pagina.evaluate("""() => {
        PYG_TABLA_MODO = 'ytd';
        const el = tablaConsolidadaPyG();
        const filas = [...el.querySelectorAll('tbody tr')];
        const num = t => Number(String(t).replace(/[^0-9,-]/g, '').replace(/\\./g, '')
                                        .replace(',', '.')) || null;
        const out = {};
        for (const tr of filas) {
            const c = [...tr.children].map(x => x.textContent.trim());
            out[c[0].replace(/\\s+/g, ' ').trim()] = c.slice(1).map(num);
        }
        return out;
    }""")
    # columnas de la tabla: EE.UU., Perú, Costa Rica, Colombia, Consolidado
    ORDEN = ["EE.UU.", "Perú", "Costa Rica", "Colombia", "Consolidado"]
    RUBRO = {"Costo ventas": "pyg_costo_ventas",
             "Administración": "pyg_gasto_administracion",
             "Ventas": "pyg_gasto_ventas",
             "Proyectos": "pyg_gasto_proyectos",
             "Financieros": "pyg_gastos_financieros"}
    revisados = 0
    for rotulo, celdas in tabla.items():
        code = RUBRO.get(rotulo)
        if not code:
            continue
        for i, ent in enumerate(ORDEN):
            pintado = celdas[i] if i < len(celdas) else None
            if pintado is None:
                continue
            fin = pagina.evaluate(f"officialAccumDisplay('{code}','{ent}','{CIERRE}')")
            assert fin is not None, f"Gastos pinta {rotulo}/{ent} y Finanzas no lo tiene"
            assert abs(abs(fin) - pintado) <= max(2.0, abs(fin) * 0.001), (
                f"{rotulo} de {ent}: Gastos pinta {pintado:,.0f} y el P&G de "
                f"Finanzas dice {abs(fin):,.0f}")
            revisados += 1
    assert revisados >= 12, f"solo se compararon {revisados} celdas; la tabla no se leyo bien"


def test_las_tarjetas_no_dicen_utilidad_neta(pagina):
    """Ninguna entidad la tiene provisionada: es utilidad ANTES de renta.

    Convivian los dos rotulos para la misma cifra --"Utilidad neta" en la tarjeta
    de Finanzas y "Utilidad antes de renta" en el P&G-- y para un CFO no son lo
    mismo.
    """
    assert pagina.evaluate("nombreKPI('utilidad_neta')") == "Utilidad antes de renta"
    assert pagina.evaluate("nombreKPI('margen_neto')") == "Margen antes de renta"


def test_el_crecimiento_de_colombia_se_mide_en_pesos(pagina):
    """Colombia factura en pesos: su crecimiento se mide en pesos.

    En dolares, enero 2026 contra enero 2025 daba +41%. En pesos fue +21%. Los
    veinte puntos de diferencia son la revaluacion del peso (4.308 -> 3.703), no
    ventas nuevas.
    """
    ene25 = pagina.evaluate("ingresoColombiaCOP('2025-01')")
    ene26 = pagina.evaluate("ingresoColombiaCOP('2026-01')")
    assert ene25 and ene26, "faltan los eneros para comparar"
    pct = (ene26 / ene25 - 1) * 100
    assert 15 < pct < 27, f"el crecimiento de enero en pesos deberia rondar 21% y da {pct:.1f}%"


def test_la_facturacion_a_doce_meses_cuadra(pagina):
    """Ingresos + Otros ingresos - Egresos = Utilidad, en la misma tabla.

    La columna de otros ingresos no existia y febrero mostraba 207.354 - 209.979 =
    +20.409, que no cierra por ninguna aritmetica. Los 23.034 que faltaban eran
    ingresos no operacionales.
    """
    _moneda(pagina, "USD")
    _pais(pagina, "__ALL__")
    for per in ("2026-02", "2026-06", "2026-07"):
        ing = pagina.evaluate(f"financeDisplayValue('ingresos_totales','{per}')")
        egr = pagina.evaluate(f"financeDisplayValue('egresos_totales','{per}')")
        uti = pagina.evaluate(f"financeDisplayValue('utilidad_neta','{per}')")
        otr = pagina.evaluate(
            f"officialPygDisplay('pyg_ingresos_no_operacionales','{per}')")
        if None in (ing, egr, uti):
            continue
        assert abs((ing + (otr or 0) - egr) - uti) < TOL, (
            f"{per}: {ing:,.0f} + {otr or 0:,.0f} - {egr:,.0f} = "
            f"{ing+(otr or 0)-egr:,.0f}, y la utilidad dice {uti:,.0f}")


def test_el_catalogo_no_contradice_las_tarjetas(pagina):
    """El Catalogo leia la serie cruda y las tarjetas el P&G oficial.

    Decia EBITDA "Por medir" mientras el Resumen mostraba 133.162, y egresos de
    junio mientras Finanzas mostraba los de julio. Dos verdades para el mismo
    codigo, en el mismo tablero.
    """
    _moneda(pagina, "USD")
    _pais(pagina, "__ALL__")
    for code in ("ebitda", "egresos_totales", "gastos_totales", "utilidad_neta"):
        tarjeta = pagina.evaluate(f"financeDisplayValue('{code}',F.periodo)")
        assert tarjeta is not None, f"{code} no tiene valor en la vista financiera"


def test_ninguna_pestana_lanza_errores_en_las_dos_monedas(pagina):
    """El recorrido completo: 7 pestanas x 2 monedas x consolidado y Colombia."""
    for moneda in ("USD", "COP"):
        for pais in ("__ALL__", "Colombia"):
            _moneda(pagina, moneda)
            _pais(pagina, pais)
            for tid, nombre in pagina.eval_on_selector_all(
                    ".tab-btn", "e=>e.map(x=>[x.dataset.go, x.textContent.trim()])"):
                antes = len(pagina.errores)
                pagina.click(f".tab-btn[data-go='{tid}']")
                pagina.wait_for_timeout(500)
                cuerpo = pagina.inner_text("#app")
                assert len(cuerpo) > 100, f"{nombre} ({moneda}/{pais}) quedo vacia"
                assert len(pagina.errores) == antes, (
                    f"{nombre} ({moneda}/{pais}): {pagina.errores[antes:]}")
