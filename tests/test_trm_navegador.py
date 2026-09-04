# -*- coding: utf-8 -*-
"""
La base de TRM, comprobada en un navegador de verdad.

POR QUE ESTE ARCHIVO EXISTE
  Los otros dos tests leen la base y comprueban cifras. Eso no alcanza para la base
  de TRM: el defecto no estaba en los datos sino en la conversion, y solo aparece
  cuando alguien mueve un control y mira que numero queda en pantalla. Estas pruebas
  abren el tablero, cambian la base de TRM y leen las tarjetas.

  Los tres defectos que motivaron el archivo, para que no vuelvan:
    1. La tasa fija se seguia aplicando despues de volver a "Promedio mensual".
    2. Un enlace compartido perdia moneda, pais, escenario y base de TRM: el arranque
       reescribia la direccion antes de leerla.
    3. La casilla de la tasa fija quedaba activa en Burn & Runway, que usa tasa propia.

  QUE SE VERIFICA DE LAS CIFRAS
  El tablero guarda cada indicador en su moneda de origen, asi que la TRM se aplica en
  una direccion u otra segun el dato:
    - INGRESOS MES es nativo USD  -> en COP se MULTIPLICA por la tasa.
    - CARTERA es de Colombia, COP -> en USD se DIVIDE por la tasa.
  Las dos direcciones se comprueban aqui; con una sola, un signo invertido pasaria.

COMO SE CORRE
  Necesita playwright, que NO es requisito del resto de los tests. Si no esta
  instalado, este archivo se salta solo y los demas siguen corriendo:

      pip install playwright && playwright install chromium
      pytest tests/test_trm_navegador.py -v
"""
import os
import re
import socket
import subprocess
import sys
import time

import pytest

pw = pytest.importorskip(
    "playwright.sync_api",
    reason="playwright no instalado; estas pruebas de navegador se omiten",
)
from playwright.sync_api import sync_playwright  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TABLERO = "index.html"
# Chromium del sistema cuando existe (entornos que ya lo traen); si no, el de playwright.
CHROME_SISTEMA = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"


def _puerto_libre():
    with socket.socket() as s:
        s.bind(("", 0))
        return s.getsockname()[1]


def _num(txt):
    """'$1.403.597.160' -> 1403597160.0   (formato es-CO)"""
    limpio = re.sub(r"[^\d.,-]", "", txt)
    return float(limpio.replace(".", "").replace(",", "."))


@pytest.fixture(scope="module")
def servidor():
    """El tablero necesita HTTP: con file:// no puede leer la base."""
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
    yield f"http://127.0.0.1:{puerto}/{TABLERO}"
    proc.terminate()
    proc.wait(timeout=10)


@pytest.fixture(scope="module")
def pagina(servidor):
    with sync_playwright() as p:
        kw = {"args": ["--no-sandbox"]}
        if os.path.exists(CHROME_SISTEMA):
            kw["executable_path"] = CHROME_SISTEMA
        navegador = p.chromium.launch(**kw)
        pg = navegador.new_page()
        pg.navegador = navegador   # para las pruebas que necesitan una segunda pestana
        pg.errores = []
        pg.on("pageerror", lambda e: pg.errores.append(str(e)))
        pg.goto(servidor)
        # S es local al script, no cuelga de window: hay que preguntar con typeof.
        pg.wait_for_function(
            "() => typeof S !== 'undefined' && S && S.periodos && S.periodos.length > 0",
            timeout=30000,
        )
        pg.wait_for_timeout(1200)
        yield pg
        navegador.close()


def _primera_tarjeta(pg):
    return _num(re.search(r"\$[\d.,]+", pg.evaluate("()=>document.querySelectorAll('#app .v.mono')[0].innerText")).group(0))


def _ultima_tarjeta(pg):
    return _num(re.search(r"\$[\d.,]+", pg.evaluate("()=>{const v=document.querySelectorAll('#app .v.mono');return v[v.length-1].innerText}")).group(0))


def _pon_tasa(pg, valor):
    pg.select_option("#f-trmbase", "fija")
    pg.wait_for_timeout(150)
    pg.fill("#f-trm-fijo", str(valor))
    pg.wait_for_timeout(650)


def test_el_selector_ofrece_la_tasa_fija(pagina):
    ops = pagina.eval_on_selector_all("#f-trmbase option", "e=>e.map(x=>x.value)")
    assert "fija" in ops
    assert pagina.is_visible("#f-trmbase"), "con dos o mas opciones el selector debe verse"


@pytest.mark.parametrize("tasa", [4000, 3268.95, 5000, 1000])
def test_dato_nativo_usd_se_multiplica_por_la_tasa(pagina, tasa):
    """INGRESOS MES nace en USD: mostrarlo en COP es multiplicar por la tasa."""
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(400)
    pagina.click("#f-moneda button[data-m='USD']")
    pagina.wait_for_timeout(500)
    base_usd = _primera_tarjeta(pagina)

    pagina.click("#f-moneda button[data-m='COP']")
    pagina.wait_for_timeout(400)
    _pon_tasa(pagina, tasa)

    assert pagina.evaluate("trmOf(F.periodo)") == tasa
    en_pantalla = _primera_tarjeta(pagina)
    esperado = base_usd * tasa
    assert abs(en_pantalla - esperado) / esperado < 0.005, (
        f"con tasa {tasa} se esperaba ~{esperado:,.0f} y hay {en_pantalla:,.0f}"
    )


@pytest.mark.parametrize("tasa", [4000, 5000])
def test_dato_nativo_cop_se_divide_por_la_tasa(pagina, tasa):
    """CARTERA es de Colombia y nace en COP: mostrarla en USD es dividir."""
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(400)
    pagina.click("#f-moneda button[data-m='COP']")
    pagina.wait_for_timeout(500)
    base_cop = _ultima_tarjeta(pagina)

    pagina.click("#f-moneda button[data-m='USD']")
    pagina.wait_for_timeout(400)
    _pon_tasa(pagina, tasa)

    en_pantalla = _ultima_tarjeta(pagina)
    esperado = base_cop / tasa
    assert abs(en_pantalla - esperado) / esperado < 0.01, (
        f"con tasa {tasa} se esperaba ~{esperado:,.0f} y hay {en_pantalla:,.0f}"
    )


def test_la_tasa_fija_no_se_filtra_a_las_otras_bases(pagina):
    """El defecto original: volver a 'Promedio mensual' seguia convirtiendo con la fija."""
    _pon_tasa(pagina, 9999)
    assert pagina.evaluate("trmOf(F.periodo)") == 9999
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(600)
    assert pagina.evaluate("trmOf(F.periodo)") != 9999


def test_vaciar_la_casilla_devuelve_la_tasa_de_la_hoja(pagina):
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(500)
    de_la_hoja = pagina.evaluate("trmOf(F.periodo)")
    _pon_tasa(pagina, 7777)
    assert pagina.evaluate("trmOf(F.periodo)") == 7777
    pagina.fill("#f-trm-fijo", "")
    pagina.wait_for_timeout(650)
    assert pagina.evaluate("trmOf(F.periodo)") == de_la_hoja


@pytest.mark.parametrize("absurdo", ["0", "-500"])
def test_una_tasa_absurda_no_se_aplica(pagina, absurdo):
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(500)
    de_la_hoja = pagina.evaluate("trmOf(F.periodo)")
    pagina.select_option("#f-trmbase", "fija")
    pagina.fill("#f-trm-fijo", absurdo)
    pagina.wait_for_timeout(650)
    assert pagina.evaluate("trmOf(F.periodo)") == de_la_hoja


def test_la_etiqueta_declara_la_tasa_que_se_uso(pagina):
    """Un tablero financiero no puede decir 'promedio mensual' mientras convierte con otra."""
    pagina.select_option("#f-trmbase", "promedio")
    pagina.wait_for_timeout(450)
    assert "fija" not in pagina.evaluate("etiquetaTasa(F.periodo)")

    pagina.select_option("#f-trmbase", "fija")
    pagina.fill("#f-trm-fijo", "")
    pagina.wait_for_timeout(500)
    assert "falta" in pagina.evaluate("etiquetaTasa(F.periodo)").lower(), (
        "elegir 'fija' sin escribir tasa cae al promedio, y hay que decirlo"
    )

    pagina.fill("#f-trm-fijo", "4200")
    pagina.wait_for_timeout(500)
    etiqueta = pagina.evaluate("etiquetaTasa(F.periodo)")
    assert "fija" in etiqueta.lower() and "4.200" in etiqueta


def test_todas_las_pestanas_pintan_con_la_tasa_fija_puesta(pagina):
    _pon_tasa(pagina, 4200)
    pestanas = pagina.eval_on_selector_all(
        ".tab-btn", "e=>e.map(x=>[x.dataset.go, x.textContent.trim()])"
    )
    assert len(pestanas) >= 5
    for tid, nombre in pestanas:
        antes = len(pagina.errores)
        pagina.click(f".tab-btn[data-go='{tid}']")
        pagina.wait_for_timeout(800)
        assert len(pagina.inner_text("#app")) > 100, f"{nombre} quedo vacia"
        assert len(pagina.errores) == antes, f"{nombre} lanzo errores: {pagina.errores[antes:]}"


def test_burn_apaga_la_casilla_porque_trae_su_propia_tasa(pagina):
    pestanas = pagina.eval_on_selector_all(
        ".tab-btn", "e=>e.map(x=>[x.dataset.go, x.textContent.trim()])"
    )
    burn = [t for t, n in pestanas if "burn" in n.lower()]
    if not burn:
        pytest.skip("no hay pestana Burn & Runway en esta base")
    pagina.click(f".tab-btn[data-go='{burn[0]}']")
    pagina.wait_for_timeout(800)
    assert pagina.evaluate("()=>document.getElementById('f-trm-fijo').disabled")
    pagina.click(f".tab-btn[data-go='{pestanas[0][0]}']")
    pagina.wait_for_timeout(500)


def test_un_enlace_compartido_restaura_lo_que_lleva(servidor, pagina):
    """
    Antes se perdia todo: applyPreset('def') termina en render(), y render reescribe la
    direccion con los valores de arranque. Cuando el codigo iba a leer el enlace, ya lo
    habia borrado -- y no solo la tasa: tambien moneda, pais, escenario y periodo.
    """
    pagina.click("#f-moneda button[data-m='COP']")
    pagina.wait_for_timeout(400)
    _pon_tasa(pagina, 4321)
    url = pagina.url
    assert "trm=4321" in url and "trmBase=fija" in url and "moneda=COP" in url

    # Pestana nueva del mismo navegador: abrir otro playwright dentro del fixture no se puede.
    pg2 = pagina.navegador.new_page()
    try:
        pg2.goto(url)
        pg2.wait_for_function(
            "() => typeof S !== 'undefined' && S && S.periodos && S.periodos.length > 0",
            timeout=30000,
        )
        pg2.wait_for_timeout(1500)
        assert pg2.input_value("#f-trmbase") == "fija"
        assert pg2.is_visible("#f-trm-fijo"), "la casilla tiene que aparecer, no solo el valor"
        assert pg2.input_value("#f-trm-fijo") == "4321"
        assert pg2.evaluate("trmOf(F.periodo)") == 4321
        assert pg2.evaluate("F.moneda") == "COP", "el enlace tambien trae la moneda"
    finally:
        pg2.close()
