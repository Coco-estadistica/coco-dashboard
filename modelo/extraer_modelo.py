# -*- coding: utf-8 -*-
"""
MOTOR DE EXTRACCIÓN — Modelo Financiero -> Base de datos del dashboard
======================================================================

Qué hace
--------
Lee el modelo (Modelo_COCO.xlsx), toma SOLO las líneas marcadas activo=SI en
mapa_modelo.csv, y las escribe en la base BD_COCO_2026.xlsx como filas normales
de BD_Indicadores.

Por qué por NOMBRE y no por celda
---------------------------------
Si el mapa dice `Inputs!C85` y algún día insertas una fila arriba, esa celda pasa
a ser otra cosa y el dashboard mostraría un dato equivocado sin avisar. Los rangos
con nombre se mueven con la celda. Por eso `tipo=nombre` es la forma recomendada;
`tipo=celda` existe para lo que aún no tiene nombre definido — úsalo con cuidado.

Uso
---
    python extraer_modelo.py              -> VISTA PREVIA (no escribe nada)
    python extraer_modelo.py --escribir   -> escribe en la base

Reglas de seguridad
-------------------
1. Por defecto NO escribe. Hay que pedirlo con --escribir.
2. Antes de escribir hace una COPIA DE SEGURIDAD de la base con fecha y hora.
3. Reemplaza (no duplica): si ya existe una extracción previa del modelo para el
   mismo indicador+periodo+escenario, la actualiza en vez de agregar otra fila.
4. Marca cada fila con segmento='Modelo' para que el dashboard NO la sume junto
   con los datos reales/contables.
"""
import csv
import os
import re
import shutil
import sys
from datetime import datetime

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
MODELO = os.path.join(AQUI, "Modelo_COCO.xlsx")
MAPA = os.path.join(AQUI, "mapa_modelo.csv")
BASE = os.path.join(AQUI, "..", "migracion", "BD_COCO_2026.xlsx")
HOJA_BD = "BD_Indicadores"
TABLA_BD = "tbl_BD_Indicadores"
SEGMENTO = "Modelo"          # separa lo proyectado de lo real
CARPETA_RESPALDO = os.path.join(AQUI, "..", "migracion", "respaldos")


# ---------------------------------------------------------------- lectura mapa
def leer_mapa():
    """Lee el mapa detectando el separador. Excel en español guarda los CSV con ';'
    y en inglés con ',' — si se asume uno solo, al reabrir el archivo desde Excel el
    extractor deja de encontrar las líneas y falla en silencio."""
    if not os.path.exists(MAPA):
        sys.exit(f"No encuentro el mapa: {MAPA}")
    with open(MAPA, encoding="utf-8-sig") as f:
        cabecera = f.readline()
    sep = ";" if cabecera.count(";") > cabecera.count(",") else ","
    filas = []
    with open(MAPA, encoding="utf-8-sig") as f:
        lector = csv.DictReader(f, delimiter=sep)
        if "activo" not in (lector.fieldnames or []):
            sys.exit(f"El mapa no tiene la columna 'activo'. Encontré: {lector.fieldnames}")
        for i, fila in enumerate(lector, start=2):
            if (fila.get("activo") or "").strip().upper() != "SI":
                continue
            fila["_linea"] = i
            filas.append(fila)
    return filas


# ------------------------------------------------------------ lectura de valor
def valor_por_nombre(wb, nombre):
    """Lee un rango con nombre. Devuelve (valor, ubicacion) o (None, motivo)."""
    if nombre not in wb.defined_names:
        return None, f"no existe el rango con nombre '{nombre}'"
    try:
        hoja, ref = list(wb.defined_names[nombre].destinations)[0]
    except Exception:
        return None, f"'{nombre}' apunta a varios rangos (no soportado)"
    celda = wb[hoja][ref.replace("$", "")]
    if isinstance(celda, tuple):  # rango -> se toma la primera celda
        celda = celda[0][0]
    return celda.value, f"{hoja}!{ref}"


def valor_por_celda(wb, origen):
    """Lee 'Hoja'!C12  o  Hoja!C12."""
    m = re.match(r"^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)$", origen.strip())
    if not m:
        return None, f"formato no reconocido: {origen} (usa Hoja!C12)"
    hoja, col, fil = m.group(1), m.group(2), m.group(3)
    if hoja not in wb.sheetnames:
        return None, f"no existe la hoja '{hoja}'"
    return wb[hoja][f"{col}{fil}"].value, f"{hoja}!{col}{fil}"


def obtener_valor(wb, fila):
    tipo = (fila.get("tipo") or "").strip().lower()
    origen = (fila.get("origen") or "").strip()
    if tipo == "nombre":
        return valor_por_nombre(wb, origen)
    if tipo == "celda":
        return valor_por_celda(wb, origen)
    return None, f"tipo desconocido '{tipo}' (usa: nombre | celda)"


# ------------------------------------------------------------------ validación
def revisar(fila, valor, ubic):
    """Devuelve lista de problemas; vacía = fila utilizable."""
    problemas = []
    if valor is None:
        problemas.append(f"sin valor ({ubic})")
    elif not isinstance(valor, (int, float)):
        problemas.append(f"el valor no es numérico: {valor!r}")
    if not (fila.get("codigo_indicador") or "").strip():
        problemas.append("falta codigo_indicador en el mapa")
    per = (fila.get("periodo") or "").strip()
    if not re.match(r"^\d{4}-(0[1-9]|1[0-2])$", per):
        problemas.append(f"periodo inválido '{per}' (formato AAAA-MM)")
    return problemas


# ------------------------------------------------------------------- escritura
def respaldar():
    os.makedirs(CARPETA_RESPALDO, exist_ok=True)
    # con segundos: dos corridas en el mismo minuto no deben pisarse el respaldo
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(CARPETA_RESPALDO, f"BD_COCO_2026_antes_de_modelo_{sello}.xlsx")
    shutil.copy2(BASE, destino)
    return destino


def escribir(listas):
    wb = openpyxl.load_workbook(BASE)
    ws = wb[HOJA_BD]
    enc = [c.value for c in ws[3]]
    idx = {n: i for i, n in enumerate(enc)}

    # índice de filas ya extraídas del modelo -> permite actualizar en vez de duplicar
    previas = {}
    for r in range(4, ws.max_row + 1):
        if ws.cell(row=r, column=idx["segmento"] + 1).value == SEGMENTO:
            clave = (ws.cell(row=r, column=idx["codigo_indicador"] + 1).value,
                     str(ws.cell(row=r, column=idx["periodo"] + 1).value),
                     ws.cell(row=r, column=idx["escenario"] + 1).value)
            previas[clave] = r

    sello = datetime.now().strftime("%Y-%m-%d %H:%M")
    nuevas = actualizadas = 0
    fila_libre = ws.max_row + 1

    for it in listas:
        clave = (it["codigo_indicador"], it["periodo"], it["escenario"])
        destino = previas.get(clave)
        if destino is None:
            destino = fila_libre
            fila_libre += 1
            nuevas += 1
        else:
            actualizadas += 1
        datos = {
            "periodo": it["periodo"], "pais": it["pais"], "compania": "Coco Colombia",
            "moneda": it["unidad"] if it["unidad"] in ("COP", "USD") else "",
            "escenario": it["escenario"], "codigo_indicador": it["codigo_indicador"],
            "segmento": SEGMENTO, "detalle": "", "valor": it["valor"],
            "unidad": it["unidad"], "area": "Finanzas", "area_responsable": "Finanzas",
            "periodicidad": "Mensual",
            "comentario": f"Extraído del modelo ({it['origen']}) el {sello}",
        }
        for col, nombre in enumerate(enc, start=1):
            ws.cell(row=destino, column=col, value=datos.get(nombre, ""))

    # extender el rango de la tabla de Excel
    tabla = ws.tables[TABLA_BD]
    ini, fin = tabla.ref.split(":")
    col_ini = "".join(c for c in ini if c.isalpha())
    fil_ini = "".join(c for c in ini if c.isdigit())
    col_fin = "".join(c for c in fin if c.isalpha())
    tabla.ref = f"{col_ini}{fil_ini}:{col_fin}{max(fila_libre - 1, ws.max_row)}"

    wb.save(BASE)
    return nuevas, actualizadas


# ------------------------------------------------------------------------ main
def main():
    escribir_de_verdad = "--escribir" in sys.argv

    for ruta, que in ((MODELO, "el modelo"), (BASE, "la base")):
        if not os.path.exists(ruta):
            sys.exit(f"No encuentro {que}: {ruta}")

    print("=" * 66)
    print("  MOTOR DE EXTRACCIÓN — Modelo -> Base del dashboard")
    print("=" * 66)
    print(f"  Modelo : {os.path.basename(MODELO)}")
    print(f"  Base   : {os.path.basename(BASE)}")
    print(f"  Modo   : {'ESCRITURA' if escribir_de_verdad else 'VISTA PREVIA (no escribe nada)'}")
    print("-" * 66)

    mapa = leer_mapa()
    if not mapa:
        print("\n  No hay ninguna línea con activo=SI en mapa_modelo.csv.")
        print("  Abre ese archivo en Excel, pon SI en las que quieras traer,")
        print("  completa su 'codigo_indicador' y vuelve a ejecutar.\n")
        return

    wb = openpyxl.load_workbook(MODELO, data_only=True)
    listos, rechazados = [], []

    for fila in mapa:
        valor, ubic = obtener_valor(wb, fila)
        problemas = revisar(fila, valor, ubic)
        if problemas:
            rechazados.append((fila, problemas))
            continue
        listos.append({
            "codigo_indicador": fila["codigo_indicador"].strip(),
            "periodo": fila["periodo"].strip(),
            "pais": (fila.get("pais") or "Colombia").strip(),
            "escenario": (fila.get("escenario") or "Modelo").strip(),
            "unidad": (fila.get("unidad") or "").strip(),
            "valor": valor, "origen": ubic, "nota": fila.get("nota", ""),
        })

    if listos:
        print(f"\n  LISTOS PARA CARGAR ({len(listos)}):\n")
        print(f"    {'indicador':<28}{'periodo':<10}{'escenario':<13}{'valor':>18}")
        print("    " + "-" * 69)
        for it in listos:
            v = f"{it['valor']:,.2f}" if isinstance(it["valor"], float) else f"{it['valor']:,}"
            print(f"    {it['codigo_indicador']:<28}{it['periodo']:<10}{it['escenario']:<13}{v:>18}")

    if rechazados:
        print(f"\n  CON PROBLEMAS ({len(rechazados)}) — no se cargan:\n")
        for fila, problemas in rechazados:
            print(f"    linea {fila['_linea']:>3}  {fila.get('origen','?'):<28} {' · '.join(problemas)}")

    print("\n" + "-" * 66)
    if not escribir_de_verdad:
        print("  Vista previa: no se escribió nada.")
        print("  Si los datos se ven bien, ejecuta:  python extraer_modelo.py --escribir\n")
        return

    if not listos:
        print("  Nada válido para escribir.\n")
        return

    copia = respaldar()
    print(f"  Respaldo de la base -> {os.path.basename(copia)}")
    nuevas, actualizadas = escribir(listos)
    print(f"  Escrito: {nuevas} fila(s) nueva(s), {actualizadas} actualizada(s).")
    print("  Abre el dashboard y vuelve a cargar la base para verlo.\n")


if __name__ == "__main__":
    main()
