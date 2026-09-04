# -*- coding: utf-8 -*-
"""
Indicadores del tablero que se calculan, comprobados en un navegador de verdad.

POR QUE ESTE ARCHIVO EXISTE
  El indice de morosidad cambio de origen el 03-sep-2026: antes se mostraba el valor
  que cargaba cartera (77,3% en julio), ahora el tablero lo calcula (41,6%). Es una
  cifra que se presenta, y la diferencia entre las dos es casi la mitad. Un cambio
  accidental aqui no rompe nada visible: solo cambia el numero.

  Tambien se comprueba que el NPS siga oculto, porque esconderlo fue una decision y
  no un descarte: el dato, el derivado y el grafico siguen en el codigo, marcados
  NPS-OFF, y es facil devolverlos sin querer.

COMO SE CORRE
  Necesita playwright, que NO es requisito del resto de los tests. Si no esta
  instalado, este archivo se salta solo:

      pip install playwright && playwright install chromium
      pytest tests/test_indicadores_tablero.py -v
"""
import os
import re
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


def _pestanas(pg):
    return pg.eval_on_selector_all(".tab-btn", "e=>e.map(x=>[x.dataset.go, x.textContent.trim()])")


def _ir_a(pg, aguja):
    for tid, nombre in _pestanas(pg):
        if aguja in nombre.lower():
            pg.click(f".tab-btn[data-go='{tid}']")
            pg.wait_for_timeout(1000)
            return nombre
    return None


# ----------------------------------------------------------------- morosidad

def test_la_morosidad_es_vencida_sobre_cartera_total(pagina):
    ven = pagina.evaluate("rawStored('cartera_vencida',{periodo:F.periodo})")
    tot = pagina.evaluate("rawStored('cartera_total',{periodo:F.periodo})")
    calc = pagina.evaluate("valueOf('indice_morosidad',{periodo:F.periodo})")
    assert ven and tot, "faltan las cifras de cartera para comprobar la formula"
    assert abs(calc - ven / tot * 100) < 0.01, (
        f"la morosidad deberia ser {ven}/{tot}*100 = {ven/tot*100:.2f} y da {calc:.2f}"
    )


def test_la_morosidad_ya_no_toma_el_valor_cargado(pagina):
    """El dato que carga cartera sigue en la base, con otra definicion. No debe usarse."""
    cargado = pagina.evaluate("rawStored('indice_morosidad',{periodo:F.periodo})")
    calc = pagina.evaluate("valueOf('indice_morosidad',{periodo:F.periodo})")
    if cargado is None:
        pytest.skip("no hay valor cargado este mes; nada que distinguir")
    assert abs(calc - cargado) > 0.5, (
        "el tablero esta mostrando el valor cargado en vez de calcularlo"
    )


def test_la_tarjeta_declara_su_formula(pagina):
    """Un porcentaje sin definicion obliga a confiar; con ella se puede rehacer."""
    assert _ir_a(pagina, "cartera"), "no encuentro la pestana de cartera"
    tarjeta = pagina.evaluate("""()=>{
        const c=[...document.querySelectorAll('#app .card')].find(e=>/morosidad/i.test(e.innerText));
        return c ? c.innerText : null; }""")
    assert tarjeta, "no encuentro la tarjeta de morosidad"
    assert "vencida" in tarjeta.lower() and "total" in tarjeta.lower(), (
        f"la tarjeta no muestra la formula:\n{tarjeta}"
    )
    # y el rotulo no puede seguir atribuyendo el dato a cartera
    assert "entregado por cartera" not in tarjeta.lower(), (
        "el titulo sigue diciendo que lo entrega cartera, y ya no es asi"
    )


def test_el_catalogo_describe_la_morosidad_como_derivada(pagina):
    ficha = pagina.evaluate("()=>JSON.stringify(S.catalog['indice_morosidad']||{})")
    assert "Derivado" in ficha, f"el Catalogo no la marca como derivada: {ficha}"
    assert "NO lo calcula" not in ficha, (
        "la ficha conserva el texto viejo que dice que el tablero no lo calcula"
    )


# ----------------------------------------------------------------------- NPS

def test_el_nps_no_aparece_en_las_vistas_de_negocio(pagina):
    con_nps = []
    for tid, nombre in _pestanas(pagina):
        if "cat" in nombre.lower():      # el Catalogo es inventario, no presentacion
            continue
        pagina.click(f".tab-btn[data-go='{tid}']")
        pagina.wait_for_timeout(700)
        if "NPS" in pagina.inner_text("#app"):
            con_nps.append(nombre)
    assert not con_nps, f"el NPS reaparecio en: {con_nps}"


def test_el_catalogo_sigue_inventariando_el_nps(pagina):
    """Ocultarlo es una decision de presentacion; borrarlo del inventario seria otra cosa."""
    assert _ir_a(pagina, "cat"), "no encuentro el Catalogo"
    assert "NPS" in pagina.inner_text("#app")


def test_la_pestana_de_churn_sobrevive_sin_el_nps(pagina):
    """
    OJO con esta: el area "cs" (Customer Success & Churn) esta DEFINIDA en AREAS pero
    areasVisibles() no la devuelve --como tampoco "ventas" ni "marketing"--, asi que el
    grafico del NPS vivia en una pestana que el tablero no pinta. Por eso ocultarlo no
    cambio nada de lo que se ve: el efecto es preventivo, para cuando esa pestana vuelva.
    El test se salta si sigue sin mostrarse, y comprueba de verdad el dia que vuelva.
    """
    nombre = _ir_a(pagina, "churn") or _ir_a(pagina, "customer")
    if not nombre:
        pytest.skip("el area 'cs' no esta entre las visibles; nada que comprobar aun")
    cuerpo = pagina.inner_text("#app")
    assert len(cuerpo) > 200, "la pestana quedo practicamente vacia al quitar el NPS"
    assert "churn" in cuerpo.lower()
    assert "NPS" not in cuerpo, "el NPS reaparecio al volver la pestana"


def test_ninguna_pestana_lanza_errores(pagina):
    for tid, nombre in _pestanas(pagina):
        antes = len(pagina.errores)
        pagina.click(f".tab-btn[data-go='{tid}']")
        pagina.wait_for_timeout(800)
        assert len(pagina.inner_text("#app")) > 100, f"{nombre} quedo vacia"
        assert len(pagina.errores) == antes, f"{nombre}: {pagina.errores[antes:]}"
