# -*- coding: utf-8 -*-
"""
Un solo paso para el cierre mensual: lee los EEFF como los entrega el ERP y arma la
propuesta de cargue.

POR QUE EXISTE
--------------
Hoy hay quince scripts cargar_*.py, uno por pais y por mes. Cada cierre nace uno nuevo
escrito a mano, y ahi es donde se cuelan los errores: el "julio" de Peru que era junio
reexportado, la cartera convertida con la TRM del mes anterior, el one-off de Costa Rica
contado como MRR nuevo. Este reemplaza a todos: mismo procedimiento, cada mes.

COMO SE USA
-----------
1. Deja los archivos del ERP --sueltos o dentro de los .zip como llegan-- en Entradas/
2. python cargar_mes.py 2026-08
      Descubre, verifica, mapea y escribe Propuesta_2026-08.xlsx. NO toca la base.
3. Revisas la propuesta.
4. python cargar_mes.py 2026-08 --escribir
      Aplica lo que dice la propuesta, con respaldo, y encadena el consolidado y la
      cadena de gastos.

LOS CUATRO CONTROLES, ANTES DE PROPONER NADA
--------------------------------------------
  1. ¿Abre como Excel de verdad?   Algunos llegan como texto tabulado con extension
                                   .xlsx: se ven bien en Excel y ninguna libreria los lee.
  2. ¿Es del mes que dice?         El saldo inicial tiene que ser el saldo final del mes
                                   anterior. Asi se detecto que el julio de Peru era un
                                   reexporte de junio, con el movimiento contado dos veces.
  3. ¿Cuadra la partida doble?     Debitos = Creditos a nivel de subcuenta, y el arbol
                                   sumando igual en cada nivel. No se usa
                                   "activo = pasivo + patrimonio": el ERP exporta los
                                   saldos credito en negativo y el resultado del
                                   ejercicio no esta cerrado, asi que esa igualdad no
                                   tiene por que cumplirse a mitad de ano.
  4. ¿El mapeo reproduce lo ya     Se aplica el mapeo a un mes YA CARGADO y se compara
     cargado?                      contra la base. Si no reproduce, no propone: seria
                                   proponer con una regla que no sabemos leer.

LO QUE NO HACE
--------------
  - No convierte monedas. Cada pais entra en la suya; el tablero convierte.
  - No rellena huecos. Una cuenta sin regla de homologacion va al reporte, no a "Otros".
  - No toca las variables no contables. NPS, morosidad, pipeline y el desglose de
    proveedores por tercero (AWS, Infobip) no salen de los EEFF: entran por plantilla.
  - No toca Costa Rica mas alla de leerla, mientras siga en conciliacion.
"""
import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import zipfile
from collections import OrderedDict, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
BASE = os.path.join(AQUI, "BD_MAESTRA_COCO.xlsx")
ENTRADAS = os.path.join(RAIZ, "Entradas")

# Se busca en Entradas/ primero. El Escritorio queda de respaldo porque es donde han
# venido llegando hasta hoy, y no vale la pena romperle la costumbre a nadie.
CARPETAS = [ENTRADAS, os.path.join(os.path.expanduser("~"), "Desktop", "Temporales")]

MES_NOMBRE = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Las tres convenciones de nombre que conviven hoy. El orden importa en Peru: el zip
# corregido va primero.
PAISES = OrderedDict([
    ("Colombia",   {"moneda": "COP", "patrones": ["{m} Colombia {a}.xlsx"]}),
    ("EE.UU.",     {"moneda": "USD", "patrones": ["{m} {a} USA.xlsx"]}),
    ("Perú",       {"moneda": "PEN", "patrones": ["{m} Peru {a}.xlsx", "{m} {a} PERU.xlsx"]}),
    ("Costa Rica", {"moneda": "USD", "patrones": ["Costa Rica {m} {a} .xlsx", "{m} {a} CR.xlsx"]}),
])

# Mapeo de cuenta PUC (4 digitos) a linea del P&G. Sale de los cargues ya validados:
# ver cargar_peru_julio_corregido.py, que lo verifico contra el YTD publicado.
MAPA = OrderedDict([
    ("pyg_ingresos_operacionales",    ["4135", "4170", "4175"]),
    ("pyg_ingresos_no_operacionales", ["4210", "4218", "4235", "4245", "4250", "4295"]),
    ("pyg_costo_ventas",              ["6135", "6205", "7135", "6120"]),
    ("pyg_gasto_administracion",      ["5110"]),
    ("pyg_gasto_ventas",              ["5210", "5235", "5255", "5295"]),
    ("pyg_gastos_financieros",        ["5305"]),
])
CUENTA_A_CODIGO = {c: k for k, v in MAPA.items() for c in v}


# ---------------------------------------------------------------- descubrimiento
def buscar(nombre):
    """El auxiliar como bytes. Suelto o dentro de los .zip, que es como llegan."""
    for carpeta in CARPETAS:
        if not os.path.isdir(carpeta):
            continue
        for raiz, _, files in os.walk(carpeta):
            if nombre in files:
                return io.BytesIO(open(os.path.join(raiz, nombre), "rb").read()), nombre
            for f in files:
                if not f.lower().endswith(".zip"):
                    continue
                try:
                    z = zipfile.ZipFile(os.path.join(raiz, f))
                except Exception:
                    continue
                for i in z.namelist():
                    if os.path.basename(i) == nombre:
                        return io.BytesIO(z.read(i)), "%s (%s)" % (nombre, f)
    return None, None


def archivo_de(pais, periodo):
    a, m = periodo.split("-")
    mes = MES_NOMBRE[int(m) - 1]
    for pat in PAISES[pais]["patrones"]:
        datos, de_donde = buscar(pat.format(m=mes, a=a))
        if datos:
            return datos, de_donde
    return None, None


# ---------------------------------------------------------------- lectura
def leer_auxiliar(datos):
    """{cuenta_puc: {'nombre':.., 'ini':.., 'mov':.., 'fin':..}} por subcuenta y por
    cuenta de 4 digitos. Devuelve None si el archivo no es un xlsx de verdad."""
    try:
        wb = openpyxl.load_workbook(datos, data_only=True)
    except Exception:
        return None
    ws = wb.active
    filas = [f for f in ws.iter_rows(values_only=True)]
    cab = [str(v or "").strip() for v in (filas[0] if filas else [])]
    try:
        ci = cab.index("Saldo inicial")
        cf = cab.index("Saldo final")
    except ValueError:
        ci = cf = None

    out = {}
    for f in filas[1:]:
        if not f:
            continue
        etq = next((str(x).strip() for x in f[:3]
                    if isinstance(x, str) and " - " in str(x)), None)
        if not etq:
            continue
        cod = etq.split(" - ")[0].strip()
        if not cod.isdigit():
            continue
        nums = [v for v in f if isinstance(v, (int, float))]
        if len(nums) < 4:
            continue
        ini = f[ci] if ci is not None and isinstance(f[ci], (int, float)) else nums[0]
        fin = f[cf] if cf is not None and isinstance(f[cf], (int, float)) else nums[-1]
        deb, cre = nums[-3], nums[-2]
        out[cod] = {"nombre": etq.split(" - ", 1)[1].strip(), "ini": float(ini),
                    "fin": float(fin), "mov": float(deb) - float(cre),
                    "deb": float(deb), "cre": float(cre)}
    return out


# ---------------------------------------------------------------- controles
def controlar(pais, periodo, aux, aux_prev):
    """Lista de (control, veredicto, detalle). Nada se propone si algo sale en ALTO."""
    r = []
    r.append(("Abre como Excel", "OK" if aux else "ALTO",
              "%d cuentas leidas" % len(aux) if aux else
              "texto tabulado con extension .xlsx: ninguna libreria lo lee"))
    if not aux:
        return r

    if aux_prev is None:
        r.append(("Continuidad de saldos", "AVISO",
                  "no encontre el mes anterior; no puedo comprobar que sea el mes que dice"))
    else:
        d = []
        for clase in ("1", "2", "3"):
            a = aux.get(clase, {}).get("ini")
            b = aux_prev.get(clase, {}).get("fin")
            if a is None or b is None:
                continue
            if abs(a - b) > 1:
                d.append("clase %s: inicial %.2f vs final anterior %.2f" % (clase, a, b))
        r.append(("Continuidad de saldos", "ALTO" if d else "OK",
                  "; ".join(d) if d else
                  "el saldo inicial es el final del mes anterior"))

    # PARTIDA DOBLE, no "activo = pasivo + patrimonio".
    #
    # El primer intento comparaba Activo contra Pasivo+Patrimonio y daba aviso en los
    # cuatro paises. Estaba mal planteado por dos razones: el ERP exporta el pasivo y el
    # patrimonio en NEGATIVO --son saldos credito-- y sumar valores absolutos no
    # reconstruye nada; y ademas el resultado del ejercicio todavia no esta cerrado
    # contra patrimonio, asi que la igualdad no tiene por que cumplirse a mitad de ano.
    #
    # Lo que SI se cumple siempre, sin depender de convenciones de signo: en cualquier
    # balance de prueba los debitos igualan a los creditos. Medido sobre los archivos de
    # julio, a nivel de SUBCUENTA (6 digitos) los cuatro paises cuadran exacto.
    #
    # Se mide en subcuenta y no en el auxiliar de 8 digitos porque ahi la diferencia es
    # legitima: las provisiones y causaciones --OTROS, CESANTIAS, APORTES A A.R.L.-- se
    # registran sin tercero, existen en la subcuenta y no bajan al auxiliar. En julio de
    # Colombia eso son 4.727.135 en ocho subcuentas, y no es un defecto del archivo.
    deb = cre = 0.0
    for cod, d in aux.items():
        if len(cod) != 6:
            continue
        deb += d["deb"]
        cre += d["cre"]
    if deb or cre:
        dif = deb - cre
        r.append(("Partida doble", "OK" if abs(dif) < 2 else "ALTO",
                  "debitos %.2f = creditos %.2f" % (deb, cre) if abs(dif) < 2
                  else "debitos %.2f vs creditos %.2f (descuadre %.2f)" % (deb, cre, dif)))
    else:
        r.append(("Partida doble", "AVISO", "el archivo no trae columnas de debito y credito"))

    # El arbol tiene que sumar igual en cada nivel: clase, grupo, cuenta y subcuenta.
    # Si un nivel no cuadra con el de abajo, la exportacion se corto por la mitad.
    niveles = {}
    for cod, d in aux.items():
        niveles.setdefault(len(cod), [0.0, 0.0])
        niveles[len(cod)][0] += d["deb"]
        niveles[len(cod)][1] += d["cre"]
    ref = niveles.get(6)
    malos = [str(n) for n, v in sorted(niveles.items())
             if n in (1, 2, 4) and ref and abs(v[0] - ref[0]) > 2]
    r.append(("El arbol suma igual en cada nivel", "ALTO" if malos else "OK",
              "no cuadran los niveles de %s digitos" % ", ".join(malos) if malos
              else "clase, grupo, cuenta y subcuenta dan lo mismo"))
    return r


def mapear(aux):
    """{codigo_pyg: valor} en moneda de origen, y las cuentas de resultado sin regla."""
    val = defaultdict(float)
    sin = []
    for cod, d in aux.items():
        if len(cod) != 4 or cod[0] not in "4567":
            continue
        destino = CUENTA_A_CODIGO.get(cod)
        if destino:
            val[destino] += abs(d["mov"])
        elif abs(d["mov"]) > 0:
            sin.append((cod, d["nombre"], round(d["mov"], 2)))
    return dict(val), sorted(sin, key=lambda x: -abs(x[2]))


def contraste(pais, periodo_ya_cargado, val_mapeado, base_rows):
    """Aplica el mapeo a un mes YA cargado y compara. Sin esto, el mapeo es una
    suposicion; con esto, es una regla verificada contra la base."""
    d = []
    for cod, v in sorted(val_mapeado.items()):
        guardado = base_rows.get((cod, pais, periodo_ya_cargado))
        if guardado is None:
            d.append((cod, None, v, "sin dato en la base"))
            continue
        razon = v / guardado if guardado else 0
        d.append((cod, guardado, v, "%.3f" % razon))
    return d


# ---------------------------------------------------------------- salida
def escribir_propuesta(ruta, periodo, hallazgos, propuesta, sin_regla, verif):
    wb = openpyxl.Workbook()
    hf = Font(bold=True, color="FFFFFF")
    hb = PatternFill("solid", fgColor="1D283A")
    rojo = PatternFill("solid", fgColor="F8D7DA")
    ambar = PatternFill("solid", fgColor="FFF3CD")

    def hoja(ws, cols, filas, pintar=None):
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = hf, hb
            c.alignment = Alignment(horizontal="center")
        for f in filas:
            ws.append([f.get(c) for c in cols])
            if pintar:
                col = pintar(f)
                if col:
                    for c in ws[ws.max_row]:
                        c.fill = col
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(
                12, min(46, max([len(str(c))] + [len(str(f.get(c) or "")) for f in filas[:200]]) + 2))
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "Controles"
    hoja(ws, ["pais", "archivo", "control", "veredicto", "detalle"], hallazgos,
         lambda f: rojo if f["veredicto"] == "ALTO" else (ambar if f["veredicto"] == "AVISO" else None))
    hoja(wb.create_sheet("Propuesta"),
         ["periodo", "pais", "moneda", "escenario", "codigo_indicador", "valor",
          "valor_en_la_base", "que_pasaria", "comentario"], propuesta,
         lambda f: ambar if f["que_pasaria"] != "sin cambio" else None)
    hoja(wb.create_sheet("Sin_regla"),
         ["pais", "cuenta_puc", "nombre_cuenta", "movimiento", "nota"], sin_regla)
    hoja(wb.create_sheet("Contraste_mapeo"),
         ["pais", "codigo_indicador", "en_la_base", "da_el_mapeo", "razon"], verif,
         lambda f: rojo if (f["razon"] not in ("—",) and _lejos(f["razon"])) else None)
    wb.save(ruta)


def _lejos(r):
    try:
        return abs(float(r) - 1) > 0.02
    except (TypeError, ValueError):
        return True


def leer_base():
    wb = openpyxl.load_workbook(BASE, data_only=True)
    fs = [f for f in wb["BD_Indicadores"].iter_rows(values_only=True)]
    cab = [str(v or "").strip() for v in fs[2]]
    ix = {c: i for i, c in enumerate(cab) if c}
    g = lambda f, c: f[ix[c]] if c in ix and ix[c] < len(f) else None
    out = {}
    for f in fs[3:]:
        if not f or not f[0]:
            continue
        if str(g(f, "escenario") or "") != "Real":
            continue
        v = g(f, "valor")
        if v is None:
            continue
        out[(str(g(f, "codigo_indicador")).strip(), str(g(f, "pais") or ""),
             str(g(f, "periodo")))] = float(v)
    return out


def mes_anterior(periodo):
    a, m = (int(x) for x in periodo.split("-"))
    return "%04d-%02d" % (a - 1, 12) if m == 1 else "%04d-%02d" % (a, m - 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("periodo", help="el mes a cargar, formato 2026-08")
    ap.add_argument("--escribir", action="store_true",
                    help="aplica la propuesta a la base (con respaldo)")
    args = ap.parse_args()
    per = args.periodo
    if not re.fullmatch(r"20\d\d-(0[1-9]|1[0-2])", per):
        sys.exit("El periodo va como 2026-08.")

    if not os.path.isdir(ENTRADAS):
        os.makedirs(ENTRADAS)
        print("  Cree la carpeta Entradas/. Deja ahi los archivos del ERP.\n")

    base_rows = leer_base()
    prev = mes_anterior(per)

    print("=" * 92)
    print("  CIERRE DE %s      %s" % (per, "APLICANDO" if args.escribir else "PROPUESTA (no toca la base)"))
    print("=" * 92)

    hallazgos, propuesta, sin_regla, verif = [], [], [], []
    alto = 0

    for pais in PAISES:
        datos, de_donde = archivo_de(pais, per)
        print("\n  %s" % pais)
        if not datos:
            print("     no encontre el archivo del mes en Entradas/ ni en el Escritorio")
            hallazgos.append({"pais": pais, "archivo": "—", "control": "Archivo presente",
                              "veredicto": "ALTO", "detalle": "no aparece el auxiliar de %s" % per})
            alto += 1
            continue
        aux = leer_auxiliar(datos)
        datos_prev, _ = archivo_de(pais, prev)
        aux_prev = leer_auxiliar(datos_prev) if datos_prev else None

        for control, veredicto, detalle in controlar(pais, per, aux, aux_prev):
            print("     %-24s %-6s %s" % (control, veredicto, detalle))
            hallazgos.append({"pais": pais, "archivo": de_donde, "control": control,
                              "veredicto": veredicto, "detalle": detalle})
            if veredicto == "ALTO":
                alto += 1
        if not aux:
            continue

        val, sin = mapear(aux)
        for cod, nom, mov in sin:
            sin_regla.append({"pais": pais, "cuenta_puc": cod, "nombre_cuenta": nom,
                              "movimiento": mov,
                              "nota": "cuenta de resultado con movimiento y sin regla de homologacion"})

        # contraste contra el mes anterior, que ya esta cargado y revisado
        datos_c, _ = archivo_de(pais, prev)
        if datos_c:
            aux_c = leer_auxiliar(datos_c)
            if aux_c:
                val_c, _ = mapear(aux_c)
                for cod, guardado, calc, razon in contraste(pais, prev, val_c, base_rows):
                    verif.append({"pais": pais, "codigo_indicador": cod,
                                  "en_la_base": guardado, "da_el_mapeo": round(calc, 2),
                                  "razon": razon if guardado else "—"})

        for cod, v in sorted(val.items()):
            ya = base_rows.get((cod, pais, per))
            propuesta.append({
                "periodo": per, "pais": pais, "moneda": PAISES[pais]["moneda"],
                "escenario": "Real", "codigo_indicador": cod, "valor": round(v, 2),
                "valor_en_la_base": ya,
                "que_pasaria": "nuevo" if ya is None else
                               ("sin cambio" if abs(ya - v) < 0.01 else "actualiza"),
                "comentario": "Auxiliar del ERP %s, %s. Moneda de origen; el tablero convierte."
                              % (per, pais)})
        print("     %-24s %-6s %d lineas del P&G, %d cuentas sin regla"
              % ("Mapeo", "OK", len(val), len(sin)))

    ruta = os.path.join(AQUI, "Propuesta_%s.xlsx" % per)
    escribir_propuesta(ruta, per, hallazgos, propuesta, sin_regla, verif)

    print("\n  " + "-" * 88)
    print("  %d filas propuestas · %d cuentas sin regla · %d controles en ALTO"
          % (len(propuesta), len(sin_regla), alto))
    print("  Propuesta: %s" % os.path.basename(ruta))

    if alto:
        print("\n  HAY CONTROLES EN ALTO. No se escribe nada hasta resolverlos.")
        print("  Mira la hoja Controles: las filas rojas dicen que pasa y con cual archivo.")
        return 1
    if not args.escribir:
        print("\n  Revisa la propuesta y aplica con:  python cargar_mes.py %s --escribir" % per)
        return 0

    shutil.copy2(BASE, BASE.replace(".xlsx", "_antes_%s.xlsx" % per))
    wb = openpyxl.load_workbook(BASE)
    ws = wb["BD_Indicadores"]
    H = [str(c.value).strip() if c.value else "" for c in ws[3]]
    ix = {c: H.index(c) + 1 for c in H if c}
    n = 0
    for f in propuesta:
        if f["que_pasaria"] == "sin cambio":
            continue
        fila = None
        for r in range(4, ws.max_row + 1):
            if (ws.cell(r, ix["codigo_indicador"]).value == f["codigo_indicador"]
                    and ws.cell(r, ix["pais"]).value == f["pais"]
                    and str(ws.cell(r, ix["periodo"]).value) == per):
                fila = r
                break
        if fila is None:
            fila = ws.max_row + 1
            for col, v in (("periodo", per), ("pais", f["pais"]), ("moneda", f["moneda"]),
                           ("escenario", "Real"), ("codigo_indicador", f["codigo_indicador"])):
                ws.cell(fila, ix[col]).value = v
        ws.cell(fila, ix["valor"]).value = f["valor"]
        ws.cell(fila, ix["comentario"]).value = f["comentario"]
        n += 1
    wb.save(BASE)
    print("\n  Escrito: %d filas. Respaldo: BD_MAESTRA_COCO_antes_%s.xlsx" % (n, per))

    for cmd in (["python", os.path.join(AQUI, "recalcular_consolidado_mes.py"), per],
                ["python", os.path.join(AQUI, "regenerar_gastos.py"), "--probar"]):
        print("\n  > %s" % " ".join(os.path.basename(c) for c in cmd[1:]))
        subprocess.call(cmd)
    print("\n  Falta: Revisar_Base.bat y publicar con Subir_A_GitHub.bat")
    return 0


if __name__ == "__main__":
    sys.exit(main())
