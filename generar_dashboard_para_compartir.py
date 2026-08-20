\
# -*- coding: utf-8 -*-
"""Construye una copia autocontenida de Dashboard_COCO.html para publicar como
Artifact (snapshot preliminar): incrusta la base xlsx (base64) y el informe
Burn & Runway, para que no dependa de fetch a archivos locales."""
import base64
import json
import os
import time

RAIZ = r"C:\Users\andre\OneDrive - CFOcus\01_CFOCUS_CLIENTES\COCO\1.0 Coco Digital\COCO_Dashboard_Cloud"
ORIGEN = os.path.join(RAIZ, "Dashboard_COCO.html")
XLSX = os.path.join(RAIZ, "migracion", "BD_MAESTRA_COCO.xlsx")
BURN = os.path.join(RAIZ, "burn_runway.json")
SALIDA = os.path.join(RAIZ, "Dashboard_COCO_para_compartir.html")


def main():
    sello = time.strftime("%d-%m-%Y %H:%M")
    html = open(ORIGEN, encoding="utf-8").read()
    b64 = base64.b64encode(open(XLSX, "rb").read()).decode("ascii")
    burn = json.load(open(BURN, encoding="utf-8"))
    burn_json = json.dumps(burn, ensure_ascii=False).replace("</", "<\\/")

    bloque = (
        "<script>window.BURN_RUNWAY_EMBEBIDO=" + burn_json + ";\n"
        "window.BD_EMBEBIDA_B64=\"" + b64 + "\";</script>\n"
    )
    marca = "<!-- snapshot preliminar para compartir, incrustado el " + time.strftime("%Y-%m-%d") + " -->\n"
    i = html.find("<script")
    html = html[:i] + marca + bloque + html[i:]

    vieja = "async function arrancar(){"
    nueva = (
        "async function arrancar(){\n"
        "  if(window.BD_EMBEBIDA_B64){\n"
        "    await cargarBurnRunway();\n"
        "    try{\n"
        "      const bin=atob(window.BD_EMBEBIDA_B64);\n"
        "      const bytes=new Uint8Array(bin.length);\n"
        "      for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);\n"
        "      S.fileName=\"BD_MAESTRA_COCO.xlsx . copia del " + sello + "\";\n"
        "      ingest(bytes);\n"
        "    }catch(e){ console.error(e); }\n"
        "    return;\n"
        "  }\n"
    )
    assert html.count(vieja) == 1, html.count(vieja)
    html = html.replace(vieja, nueva, 1)

    open(SALIDA, "w", encoding="utf-8").write(html)
    print("OK ->", SALIDA)
    print("tamano KB:", round(len(html) / 1024, 1))


if __name__ == "__main__":
    main()
