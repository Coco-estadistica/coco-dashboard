# -*- coding: utf-8 -*-
"""
Genera apps_script/Dashboard.html — el archivo que se pega dentro del proyecto de
Apps Script para que Google sirva el tablero.

Qué hace, además de copiar:
  1. Incrusta burn_runway.json dentro de la página. Dentro de Apps Script no hay
     archivos que descargar, así que el informe de caja debe viajar en el HTML.
  2. Deja una marca de generación para saber de cuándo es la versión publicada.

Cuándo volver a correrlo:
  - Cada vez que cambie el tablero (Dashboard_COCO.html)
  - Cada vez que actualices el modelo y regeneres el informe de Burn & Runway

Uso:  python generar_Dashboard_html.py
Luego: copia TODO el contenido de Dashboard.html y pégalo en el archivo
       "Dashboard.html" del proyecto de Apps Script.
"""
import json
import os
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(AQUI, "..", "Dashboard_COCO.html")
BURN = os.path.join(AQUI, "..", "burn_runway.json")
SALIDA = os.path.join(AQUI, "Dashboard.html")


def main():
    if not os.path.exists(ORIGEN):
        raise SystemExit(f"No encuentro el tablero: {ORIGEN}")
    html = open(ORIGEN, encoding="utf-8").read()

    # --- incrustar el informe de Burn & Runway ---
    if os.path.exists(BURN):
        datos = json.load(open(BURN, encoding="utf-8"))
        # </script> dentro de un string rompería el bloque: se escapa
        crudo = json.dumps(datos, ensure_ascii=False).replace("</", "<\\/")
        bloque = ("<script>window.BURN_RUNWAY_EMBEBIDO=" + crudo + ";</script>\n")
        marca = f"<!-- generado para Apps Script el {datetime.now():%Y-%m-%d %H:%M} -->\n"
        # se inserta justo antes del primer <script> de la página
        i = html.find("<script")
        html = html[:i] + marca + bloque + html[i:]
        estado_burn = f"incrustado ({len(datos.get('resumen', []))} indicadores)"
    else:
        estado_burn = "NO incrustado — corre antes modelo/extraer_burn_runway.py"

    open(SALIDA, "w", encoding="utf-8").write(html)

    print("Archivo para Apps Script generado:")
    print(f"   {os.path.abspath(SALIDA)}")
    print(f"   Tamaño: {len(html)/1024:.0f} KB")
    print(f"   Informe Burn & Runway: {estado_burn}")
    print("\nSiguiente paso:")
    print("   1. Abre Dashboard.html con el Bloc de notas")
    print("   2. Selecciona todo (Ctrl+E) y copia (Ctrl+C)")
    print("   3. En Apps Script: archivo 'Dashboard.html' -> pega y guarda")
    print("   4. Implementar -> Administrar implementaciones -> editar -> Nueva version")


if __name__ == "__main__":
    main()
