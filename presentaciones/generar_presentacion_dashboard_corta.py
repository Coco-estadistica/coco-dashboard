# -*- coding: utf-8 -*-
"""Version corta (4 laminas) de la presentacion de arquitectura del tablero."""
import os

from estilo_coco import (
    Presentation, PLANTILLA, BASE, Inches, MSO_SHAPE, PP_ALIGN, MSO_ANCHOR,
    CARD, TEXTO, MUTED, VERDE, VERDE2, CYAN, ROJO, AMBAR,
    MX, CW, txt, rect, cabecera, pie, vinetas, kpi, tabla,
)

SALIDA = os.path.join(BASE, "Dashboard COCO - Resumen 4 laminas.pptx")

prs = Presentation(PLANTILLA)
for sld in list(prs.slides._sldIdLst):
    prs.part.drop_rel(sld.rId)
    prs.slides._sldIdLst.remove(sld)
LAYOUT = next(l for l in prs.slide_layouts if l.name == "COCO")


def nueva():
    return prs.slides.add_slide(LAYOUT)


def flecha(slide, x, y, w=0.22, h=0.22, abajo=True):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.DOWN_ARROW if abajo else MSO_SHAPE.RIGHT_ARROW,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CYAN
    sh.line.fill.background()
    sh.shadow.inherit = False


def izq(t):
    return (t, TEXTO, False, PP_ALIGN.LEFT)


# ===================================================== 1 QUE ES
s = nueva()
txt(s, MX, 0.55, CW, 0.55, "Cockpit Financiero COCO", 32, VERDE, True)
txt(s, MX, 1.12, CW, 0.30,
    "Qué es, de qué está hecho y de dónde saca las cifras · corte julio 2026", 11, MUTED)

kw, kg = 2.93, 0.16
kpi(s, MX + 0 * (kw + kg), 1.65, kw, 1.40, "1", "Archivo HTML",
    "abre sin internet · 1,17 MB", VERDE2, 30)
kpi(s, MX + 1 * (kw + kg), 1.65, kw, 1.40, "13", "Pestañas",
    "de Resumen a Catálogo", VERDE2, 30)
kpi(s, MX + 2 * (kw + kg), 1.65, kw, 1.40, "4", "Países",
    "Colombia · EE.UU. · Perú · Costa Rica", CYAN, 30)
kpi(s, MX + 3 * (kw + kg), 1.65, kw, 1.40, "1.852", "Filas de hechos",
    "una sola fuente de verdad", CYAN, 30)

txt(s, MX, 3.28, CW, 0.3, "DE QUÉ ESTÁ HECHO", 9.5, CYAN, True)
anchos = [3.30, 2.60, 6.30]
filas = [
    ["Componente", "Lenguaje", ("Qué hace", CYAN, True, PP_ALIGN.LEFT)],
    ["Dashboard_COCO.html", "HTML · CSS · JavaScript · SVG", izq("El tablero completo en un archivo: interfaz, cálculo y gráficos")],
    ["27 scripts (.py)", "Python 3", izq("Prepara el dato antes: carga, corrige TRM, concilia capas")],
    ["apps_script/Code.gs", "Google Apps Script", izq("Capa de nube: cada área cargaría su porción (listo, sin desplegar)")],
    ["3 botones (.bat)", "Batch / CMD", izq("Uso diario: abrir el tablero, revisar la base, corregir la base")],
    ["BD_MAESTRA_COCO.xlsx", "Excel · 14 hojas", izq("La única base que el tablero consume — abre 10 de sus 14 hojas")],
]
tabla(s, MX, 3.62, anchos, filas, alto_fila=0.42, size=10)

txt(s, MX, 6.20, CW, 0.55,
    "Sin servidor, sin base de datos, sin licencias y sin librerías descargadas: los gráficos son SVG escrito a mano. "
    "Por eso se abre con doble clic y se puede enviar por correo.", 10, TEXTO)
pie(s, "25 de agosto de 2026 · versión del tablero JULIO-2026-R4.")

# =========================================== 2 DE DONDE SALE LA INFORMACION
s = nueva()
cabecera(s, "De dónde extrae la información",
         "Cuatro niveles, de la fuente contable original a la cifra en pantalla")

y0, bh, bgap = 1.25, 1.15, 0.28
niveles = [
    ("NIVEL 1 · FUENTES PRIMARIAS", VERDE2,
     "Estados financieros oficiales de Colombia (PDF)   ·   Auxiliares contables mensuales por país — 22 archivos, 986 registros\n"
     "Modelo Presupuesto Completo COCO V3   ·   Facturación y CRM   ·   Extractos bancarios"),
    ("NIVEL 2 · TRAZABILIDAD — BD_TRAZABILIDAD_COCO.xlsx · 24 hojas", AMBAR,
     "Homologación (90 reglas de mapeo) · Balance por país · Movimientos · Intercompañías · Alertas · Controles\n"
     "Es el archivo de auditoría: guarda de dónde salió cada cifra.  ⚠  El tablero NUNCA lo abre."),
    ("NIVEL 3 · BASE OPERATIVA — BD_MAESTRA_COCO.xlsx · 14 hojas", CYAN,
     "BD_Indicadores — la tabla de hechos: 1.852 filas (periodo · país · moneda · escenario · indicador · valor · fuente)\n"
     "Más: Diccionario · TRM y TRM_Perú · Churn · Pipeline · Catálogos · Log de cargues · cadena de gastos"),
    ("NIVEL 4 · CONSUMO — Dashboard_COCO.html", VERDE2,
     "Abre exactamente 10 hojas de la base maestra   +   burn_runway.json (viene del modelo) para la pestaña Burn & Runway\n"
     "Si un número no está en esas 10 hojas, no existe para el tablero."),
]
for i, (titulo, col, cuerpo) in enumerate(niveles):
    y = y0 + i * (bh + bgap)
    rect(s, MX, y, CW, bh)
    txt(s, MX + 0.20, y + 0.15, CW - 0.40, 0.26, titulo, 11, col, True)
    txt(s, MX + 0.20, y + 0.48, CW - 0.40, 0.60, cuerpo, 9.5, MUTED, space=4)
    if i < 3:
        flecha(s, MX + CW / 2 - 0.11, y + bh + 0.03)

pie(s, "Los niveles 2 y 3 se separaron el 13-ago-2026: la misma cifra vivía en varias hojas y, al actualizar una y olvidar las otras, "
       "el tablero mostraba datos viejos sin avisar. Hoy solo hay una capa que el tablero consume.")

# ============================================ 3 COMO FUNCIONA Y COMO SE CARGA
s = nueva()
cabecera(s, "Cómo funciona y cómo se carga la información",
         "El recorrido de cada apertura, y las cinco vías de cargue")

pasos = [
    ("1 · ABRE", ["El .bat levanta un", "servidor local y abre", "el tablero"]),
    ("2 · CARGA", ["Busca la base en disco.", "3 reintentos, porque", "vive en OneDrive"]),
    ("3 · LEE", ["Abre exactamente", "10 hojas del Excel", "y ninguna más"]),
    ("4 · FILTRA", ["Fecha · escenario · país", "· moneda · base de TRM"]),
    ("5 · DIBUJA", ["Pinta la pestaña activa", "con gráficos SVG"]),
]
pw, pgap = 2.25, 0.24
for i, (num, items) in enumerate(pasos):
    x = MX + i * (pw + pgap)
    rect(s, x, 1.28, pw, 1.30)
    txt(s, x + 0.15, 1.44, pw - 0.30, 0.24, num, 10, VERDE2, True)
    txt(s, x + 0.15, 1.76, pw - 0.30, 0.72, "\n".join(items), 9, MUTED, space=3)
    if i < 4:
        flecha(s, x + pw + 0.02, 1.82, 0.20, 0.22, abajo=False)

txt(s, MX, 2.78, CW, 0.3, "CINCO MANERAS DE CARGAR LA INFORMACIÓN", 9.5, CYAN, True)
anchos = [0.42, 2.80, 1.45, 7.53]
filas = [
    [("", CYAN, True, PP_ALIGN.CENTER), ("Vía", CYAN, True, PP_ALIGN.LEFT),
     ("Estado", CYAN, True, PP_ALIGN.CENTER), ("Cómo funciona", CYAN, True, PP_ALIGN.LEFT)],
    [("A", CYAN, True, PP_ALIGN.CENTER), izq("Automática desde disco"), ("Activa", VERDE2, True, PP_ALIGN.CENTER),
     izq("El tablero busca la base solo. Es la vía normal: doble clic y listo")],
    [("B", CYAN, True, PP_ALIGN.CENTER), izq("Manual «Cargar BD»"), ("Activa", VERDE2, True, PP_ALIGN.CENTER),
     izq("Arrastrar el Excel al tablero, para revisar una versión sin tocar la oficial")],
    [("C", CYAN, True, PP_ALIGN.CENTER), izq("Caché del navegador"), ("Respaldo", AMBAR, True, PP_ALIGN.CENTER),
     izq("Copia de la última carga. Solo si falla la lectura, y avisa en ámbar que puede estar vieja")],
    [("D", CYAN, True, PP_ALIGN.CENTER), izq("Cargue por área en la nube"), ("Preparada", CYAN, True, PP_ALIGN.CENTER),
     izq("Cada área entra con su cuenta y sube su porción, con previsualización antes de confirmar")],
    [("E", CYAN, True, PP_ALIGN.CENTER), izq("Desde el modelo financiero"), ("Activa", VERDE2, True, PP_ALIGN.CENTER),
     izq("Un botón regenera Burn & Runway desde el modelo V3")],
]
tabla(s, MX, 3.12, anchos, filas, alto_fila=0.42, size=9.8)

txt(s, MX, 5.85, CW, 0.3, "LA REGLA QUE SOSTIENE TODO", 9.5, CYAN, True)
rect(s, MX, 6.18, CW, 0.62)
txt(s, MX + 0.20, 6.30, CW - 0.40, 0.42,
    "Un dato que no llegó se marca «s/d», nunca cero. El tablero distingue tres cosas: $0 es un cero reportado, "
    "«s/d» es un dato pendiente de cargar y «n/a» es una cuenta que no aplica a esa entidad.",
    10, TEXTO, anchor=MSO_ANCHOR.MIDDLE)
pie(s, "La vía D está lista y documentada pero todavía sin desplegar: hoy el tablero corre en el computador.")

# ================================================ 4 CIFRAS CON CONTABILIDAD
s = nueva()
cabecera(s, "Cifras que hay que revisar con contabilidad",
         "Lo que impide firmar un balance consolidado hoy")

txt(s, MX, 1.20, 6.00, 0.3, "① ELIMINACIONES INTERCOMPAÑÍA — SIN PRACTICAR", 10.5, ROJO, True)
anchos = [3.55, 1.35, 1.10]
filas = [
    [("Partida (julio 2026)", CYAN, True, PP_ALIGN.LEFT), "USD", ("", CYAN, True, PP_ALIGN.LEFT)],
    [izq("CxC EE.UU. → Coco Inversiones (Colombia)"), "571.892", ""],
    [izq("CxC EE.UU. → Costa Rica"), "37.220", ""],
    [izq("CxC EE.UU. → Perú"), "4.300", ""],
    [izq("Ingresos EE.UU. ← Perú"), "4.010", ""],
    [("TOTAL POR ELIMINAR", TEXTO, True, PP_ALIGN.LEFT), ("617.422", VERDE2, True, PP_ALIGN.RIGHT), ""],
]
tabla(s, MX, 1.55, anchos, filas, alto_fila=0.29, alto_head=0.29, size=9.8)
txt(s, MX, 3.42, 6.00, 0.42,
    "El 99,5% de las «otras cuentas por cobrar» de EE.UU. es intercompañía. Sin eliminar, el activo del grupo "
    "está inflado en más de medio millón de dólares.", 9.5, MUTED)

txt(s, MX, 4.00, 6.00, 0.3, "② TRES TASAS DISTINTAS PARA LO MISMO", 10.5, ROJO, True)
anchos = [3.55, 1.35, 1.10]
filas = [
    [("Dónde se usa", CYAN, True, PP_ALIGN.LEFT), "COP/USD", ""],
    [izq("Modelo V3 — tasa de presupuesto"), "3.690,00", ""],
    [izq("Hoja Consolidado — tasa única"), "3.505,27", ""],
    [izq("Burn & Runway — tasa de cierre"), "3.248,87", ""],
]
tabla(s, MX, 4.35, anchos, filas, alto_fila=0.29, alto_head=0.29, size=9.8)
txt(s, MX, 5.62, 6.00, 0.60,
    "Los mismos ingresos de Colombia de junio valen USD 203.646 o USD 214.378 según la tasa: 5,3% de diferencia. "
    "Falta definir política: promedio del mes para el resultado, cierre para el balance.", 9.5, MUTED)

x2 = MX + 6.25
txt(s, x2, 1.20, 5.95, 0.3, "③ EL RESULTADO NO CUADRA CON EL CIERRE", 10.5, ROJO, True)
anchos2 = [1.70, 1.45, 1.45, 1.35]
filas2 = [
    [("País", CYAN, True, PP_ALIGN.LEFT), "Calculado", "Registrado", "Dif."],
    [izq("EE.UU. (USD)"), "-48.548", "13.043", "61.590"],
    [izq("Costa Rica (USD)"), "109.029", "-22.674", "131.703"],
    [izq("Perú (PEN)"), "25.372", "0", "25.372"],
]
tabla(s, x2, 1.55, anchos2, filas2, alto_fila=0.29, alto_head=0.29, size=9.8)
txt(s, x2, 2.85, 5.95, 0.30,
    "Las cuentas de cierre no reflejan lo abierto en ingresos y gastos.", 9.5, MUTED)

txt(s, x2, 3.28, 5.95, 0.3, "④ COSTA RICA — PRECIOS DE TRANSFERENCIA", 10.5, ROJO, True)
rect(s, x2, 3.62, 5.95, 0.95)
txt(s, x2 + 0.18, 3.76, 5.59, 0.70,
    "Factura con margen aparente del 98%: los costos de ese servicio están en otra entidad.\n"
    "Julio: USD 134.104 facturados, contra Clientes y no contra el banco. Sin Costa Rica, julio da pérdida de USD 886.",
    9.5, MUTED, space=4)

txt(s, x2, 4.72, 5.95, 0.3, "⑤ COLOMBIA — RENTA NO PROVISIONADA", 10.5, AMBAR, True)
rect(s, x2, 5.06, 5.95, 0.80)
txt(s, x2 + 0.18, 5.20, 5.59, 0.58,
    "Siete meses con impuesto de renta en cero. Provisión estimada al 35%: COP 99,63 MM.\n"
    "Utilidad neta COP 246,73 MM  →  con renta, COP 147,10 MM.", 9.5, MUTED, space=4)

txt(s, x2, 6.00, 5.95, 0.3, "⑥ HUECOS DE CARGUE", 10.5, AMBAR, True)
rect(s, x2, 6.34, 5.95, 0.50)
txt(s, x2 + 0.18, 6.44, 5.59, 0.34,
    "78 indicadores incompletos · MRR y Customer Success van a mayo · el balance de Costa Rica a junio.",
    9.5, MUTED)
pie(s, "Hoy sí se puede afirmar el P&G consolidado enero–julio en USD. NO un balance consolidado: faltan las eliminaciones y la política de conversión.")

prs.save(SALIDA)
print("OK ->", SALIDA, "|", len(prs.slides._sldIdLst), "laminas")
