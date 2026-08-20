# -*- coding: utf-8 -*-
"""
Genera "Presentacion Resultados Consolidados a Julio 2026.pptx" con el formato
COCO (tomado de COCO_Junta_Jun2026_8_Diapositivas.pptx: fondo #142433, Arial,
titulo verde #5BC770, acento cian #35C3DE, negativo #E8635E).

Reestructura la presentacion de julio para que el consolidado abra el documento
en vez de cerrarlo. Cifras contrastadas contra BD_MAESTRA_COCO y el modelo V3.
"""
import copy
import os

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                             XL_LABEL_POSITION, XL_TICK_LABEL_POSITION)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

BASE = r"C:\Users\andre\OneDrive - CFOcus\01_CFOCUS_CLIENTES\COCO\1.0 Coco Digital"
PLANTILLA = os.path.join(BASE, "COCO_Junta_Jun2026_8_Diapositivas.pptx")
SALIDA = os.path.join(BASE, "Presentacion Resultados Consolidados Julio 2026.pptx")

# --- paleta COCO ---
BG      = RGBColor(0x14, 0x24, 0x33)
CARD    = RGBColor(0x1C, 0x34, 0x46)
CARD2   = RGBColor(0x20, 0x3C, 0x50)
BORDE   = RGBColor(0x2C, 0x4B, 0x60)
TEXTO   = RGBColor(0xEA, 0xF0, 0xF4)
MUTED   = RGBColor(0x93, 0xA9, 0xB8)
VERDE   = RGBColor(0x5B, 0xC7, 0x70)
VERDE2  = RGBColor(0x4F, 0xCD, 0x87)
CYAN    = RGBColor(0x35, 0xC3, 0xDE)
ROJO    = RGBColor(0xE8, 0x63, 0x5E)
AMBAR   = RGBColor(0xE0, 0xA6, 0x4A)

W = 13.3333
MX = 0.55          # margen izquierdo
CW = 12.20         # ancho util


# ---------- formato de numeros (es-CO) ----------
def n0(v):
    return f"{v:,.0f}".replace(",", ".")


def n1(v):
    return f"{v:,.1f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def pct(v, d=1):
    return f"{v:.{d}f}".replace(".", ",") + "%"


def usd(v):
    return "USD " + n0(v)


# ---------- helpers de dibujo ----------
def txt(slide, x, y, w, h, text, size=11, color=TEXTO, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if space and i:
            p.space_before = Pt(space)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = "Arial"
        r.font.color.rgb = color
    return box


def rect(slide, x, y, w, h, fill=CARD, line=BORDE, lw=0.75):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def cabecera(slide, titulo, sub):
    txt(slide, MX, 0.28, CW, 0.46, titulo, 26, VERDE, True, anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, MX, 0.76, CW, 0.30, sub, 10, MUTED, anchor=MSO_ANCHOR.TOP)


def pie(slide, nota, y=6.92):
    txt(slide, MX, y, CW, 0.42, nota, 8.5, MUTED)


def vinetas(slide, x, y, w, items, size=11.5, gap=9):
    """items: lista de (texto, color_resalte|None)"""
    cur = y
    for t, col in items:
        rect(slide, x, cur + 0.085, 0.055, 0.055, fill=col or CYAN, line=None)
        b = txt(slide, x + 0.20, cur, w - 0.20, 0.30, t, size, TEXTO)
        b.text_frame.word_wrap = True
        # alto estimado por longitud
        n = max(1, int(len(t) / (w * 8.6)) + 1)
        cur += 0.235 * n + gap / 72.0
    return cur


def kpi(slide, x, y, w, h, valor, etiqueta, nota="", color=VERDE2, vsize=30):
    rect(slide, x, y, w, h)
    txt(slide, x + 0.22, y + 0.20, w - 0.44, 0.28, etiqueta.upper(), 9, CYAN, True)
    txt(slide, x + 0.22, y + 0.52, w - 0.44, 0.52, valor, vsize, color, True)
    if nota:
        txt(slide, x + 0.22, y + h - 0.52, w - 0.44, 0.42, nota, 9, MUTED)


def tabla(slide, x, y, anchos, filas, alto_fila=0.30, alto_head=0.32, size=10.5):
    """filas[0] = encabezado. Cada celda: str o (str, color, bold, align)"""
    cur = y
    for i, fila in enumerate(filas):
        es_head = (i == 0)
        h = alto_head if es_head else alto_fila
        cx = x
        fondo = CARD2 if es_head else (CARD if i % 2 else None)
        for j, celda in enumerate(fila):
            w = anchos[j]
            if fondo is not None:
                rect(slide, cx, cur, w, h, fill=fondo, line=None)
            if isinstance(celda, tuple):
                t, col, bold, al = celda
            else:
                t, col, bold, al = celda, (CYAN if es_head else TEXTO), es_head, \
                    (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
            pad = 0.11
            txt(slide, cx + pad, cur, w - 2 * pad, h, t,
                9.5 if es_head else size, col, bold, al, MSO_ANCHOR.MIDDLE)
            cx += w
        # linea inferior
        rect(slide, x, cur + h - 0.01, sum(anchos), 0.01, fill=BORDE, line=None)
        cur += h
    return cur


def estilo_grafico(chart, leyenda=True, tam=9.5):
    chart.font.size = Pt(tam)
    chart.font.name = "Arial"
    chart.font.color.rgb = MUTED
    # PowerPoint pone titulo automatico cuando hay una sola serie: se quita
    # siempre, el titulo de cada grafico va como texto propio de la lamina.
    chart.has_title = False
    if leyenda:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9.5)
        chart.legend.font.color.rgb = TEXTO
    else:
        chart.has_legend = False
    try:
        va = chart.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = BORDE
        va.major_gridlines.format.line.width = Pt(0.5)
        va.tick_labels.font.size = Pt(9)
        va.tick_labels.font.color.rgb = MUTED
        va.format.line.color.rgb = BORDE
    except Exception:
        pass
    try:
        ca = chart.category_axis
        ca.has_major_gridlines = False
        ca.tick_labels.font.size = Pt(9.5)
        ca.tick_labels.font.color.rgb = MUTED
        ca.format.line.color.rgb = BORDE
    except Exception:
        pass


def serie_color(chart, idx, color, transparente=False):
    s = chart.series[idx]
    if transparente:
        s.format.fill.background()
        s.format.line.fill.background()
    else:
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = color
        s.format.line.fill.background()
    # sin esto PowerPoint pinta de BLANCO cualquier barra negativa (invierte el
    # relleno por defecto), y en fondo oscuro la barra desaparece del grafico.
    try:
        s.invert_if_negative = False
    except Exception:
        pass
    return s


def eje_cat_abajo(chart):
    """Con valores negativos PowerPoint dibuja las etiquetas de mes sobre la
    linea del cero, encima de las barras. Se bajan al pie del grafico."""
    try:
        chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    except Exception:
        pass


def punto_color(chart, sidx, pidx, color):
    """Colorea un punto suelto. Ademas apaga invertIfNegative EN EL PUNTO: el
    ajuste a nivel de serie no lo hereda, y una barra negativa se seguia
    pintando blanca encima del color asignado."""
    ser = chart.series[sidx]
    pt = ser.points[pidx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = color
    # OJO: pt._element es el <c:ser>, no el <c:dPt>. El punto se pide asi:
    dPt = ser._element.get_or_add_dPt_for_point(pidx)
    if dPt.find(qn("c:invertIfNegative")) is None:
        # orden del esquema en c:dPt: idx, invertIfNegative, marker, ..., spPr
        dPt.find(qn("c:idx")).addnext(
            parse_xml(f'<c:invertIfNegative {nsdecls("c")} val="0"/>'))


def etiquetas_es(chart, idx, valores, dec=0, pos="ctr", color=TEXTO, size=9.5, bold=False):
    """Etiquetas de datos escritas como texto literal ya formateado en es-CO.

    PowerPoint ignora el prefijo de configuracion regional ([$-240A]) al pintar
    un grafico: usa la del equipo, y en este sale "214,383" cuando el resto de
    la lamina dice "214.383". Escribiendo el texto a mano el formato queda fijo.
    """
    ser = chart.series[idx]._element
    dLbls = ser.get_or_add_dLbls()
    for hijo in list(dLbls):
        dLbls.remove(hijo)
    # color puede ser uno solo o una lista por punto: en un puente, un tramo
    # delgado no alcanza a contener su etiqueta y el texto oscuro se derrama
    # sobre el fondo, tambien oscuro, y desaparece. Esos van en claro.
    # OJO: RGBColor hereda de tuple, asi que un color suelto pasaria por lista
    es_lista = isinstance(color, (list, tuple)) and not isinstance(color, RGBColor)
    for i, v in enumerate(valores):
        if v is None:
            continue
        col = str(color[i] if es_lista else color)
        t = n1(v) if dec else n0(v)
        dLbls.append(parse_xml(
            f'<c:dLbl {nsdecls("c", "a")}><c:idx val="{i}"/>'
            f'<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
            f'<a:rPr lang="es-CO" sz="{int(size * 100)}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:latin typeface="Arial"/></a:rPr><a:t>{t}</a:t>'
            f'</a:r></a:p></c:rich></c:tx>'
            f'<c:dLblPos val="{pos}"/>'
            f'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
            f'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>'
            f'</c:dLbl>'))
    for tag in ("showLegendKey", "showVal", "showCatName", "showSerName",
                "showPercent", "showBubbleSize"):
        dLbls.append(parse_xml(
            f'<c:{tag} {nsdecls("c")} val="{1 if tag == "showVal" else 0}"/>'))


def etiquetas(chart, idx, fmt='#,##0', pos=XL_LABEL_POSITION.OUTSIDE_END,
              color=TEXTO, size=9.5, bold=False):
    s = chart.series[idx]
    s.has_data_labels = True
    dl = s.data_labels
    # has_data_labels solo crea el elemento: sin show_value el XML queda con
    # <c:showVal val="0"/> y PowerPoint no pinta ningun numero.
    dl.show_value = True
    # [$-240A] = espanol (Colombia): sin esto PowerPoint pinta los miles con coma
    # ("109,029") y choca con las tablas de la lamina, que usan punto.
    dl.number_format = "[$-240A]" + fmt
    dl.number_format_is_linked = False
    dl.font.size = Pt(size)
    dl.font.bold = bold
    dl.font.name = "Arial"
    dl.font.color.rgb = color
    try:
        dl.position = pos
    except Exception:
        pass


# ================= construccion =================
prs = Presentation(PLANTILLA)

# quitar las 8 diapositivas de la plantilla, conservando master/layout "COCO"
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    prs.part.drop_rel(sld.rId)
    sldIdLst.remove(sld)

LAYOUT = next(l for l in prs.slide_layouts if l.name == "COCO")


def nueva():
    return prs.slides.add_slide(LAYOUT)


# ---------------------------------------------------------------- 1 PORTADA
s = nueva()
txt(s, MX, 2.05, CW, 0.9, "COCO Tecnologías", 44, VERDE, True)
txt(s, MX, 3.00, CW, 0.6, "Resultados consolidados a julio 2026", 26, TEXTO)
rect(s, MX, 3.80, 2.6, 0.035, fill=CYAN, line=None)
txt(s, MX, 4.10, 9.0, 0.9,
    "Informe para junta directiva · cierre a 31 de julio de 2026\n"
    "Consolidado de Colombia, Costa Rica, Estados Unidos y Perú",
    13, MUTED, space=5)
txt(s, MX, 6.30, CW, 0.5,
    "Cifras en USD para el grupo y COP millones para Colombia y caja. "
    "Fuente: estados financieros oficiales, auxiliares contables de las filiales y modelo presupuestal V3.",
    9, MUTED)

# ------------------------------------------------- 2 EL GRUPO EN UNA PAGINA
s = nueva()
cabecera(s, "El grupo en una página",
         "Consolidado enero–julio 2026 · Colombia, Costa Rica, Estados Unidos y Perú")
kx, kw, kg = MX, 2.93, 0.16
kpi(s, kx + 0 * (kw + kg), 1.25, kw, 1.55, "1.651.411", "Ingresos netos", "USD · acumulado ene–jul", VERDE2, 27)
kpi(s, kx + 1 * (kw + kg), 1.25, kw, 1.55, "127.717", "EBITDA", "USD · margen 7,7%", VERDE2, 27)
kpi(s, kx + 2 * (kw + kg), 1.25, kw, 1.55, "137.766", "Utilidad neta", "USD · margen 8,3%", VERDE2, 27)
kpi(s, kx + 3 * (kw + kg), 1.25, kw, 1.55, "1.162,6", "Caja consolidada", "COP MM · 8,8 meses de runway", CYAN, 27)

txt(s, MX, 3.15, CW, 0.3, "LO QUE HAY QUE SABER", 9.5, CYAN, True)
vinetas(s, MX, 3.55, 12.1, [
    ("El grupo cierra siete meses con utilidad, pero la pone Costa Rica: sin su primera facturación el consolidado del año sería de USD 28.737, no de 137.766.", VERDE),
    ("La utilidad neta (137.766) supera al EBITDA (127.717) porque todavía no se registra la provisión de renta — solo en Colombia son COP 99,6 MM acumulados del año.", AMBAR),
    ("El 89% del ingreso lo genera Colombia, pero el 85% de la caja está en Estados Unidos, que pierde USD 48.548 en el año.", ROJO),
    ("Julio fue el mejor mes del año en ingreso consolidado (USD 350.895), impulsado por el one-off de Costa Rica de USD 134.104 — íntegro por cobrar.", CYAN),
])
pie(s, "Consolidado armado desde los auxiliares contables de las filiales y los estados financieros oficiales de Colombia. "
       "Caja convertida a TRM 3.268,95. El EBITDA del grupo no incluye provisión de renta.")

# ------------------------------------------- 3 SIN COSTA RICA NO HAY UTILIDAD
s = nueva()
cabecera(s, "Sin Costa Rica no hay utilidad",
         "Utilidad neta acumulada enero–julio 2026, USD · las cuatro filiales suman exacto el consolidado")

cd = CategoryChartData()
cd.categories = ["Colombia", "Costa Rica", "Perú", "Estados Unidos", "Consolidado"]
cd.add_series("Base",   (None, 70388, 179417, 137766, None))
cd.add_series("Total",  (70388, None, None, None, 137766))
cd.add_series("Aporta", (None, 109029, 6897, None, None))
cd.add_series("Resta",  (None, None, None, 48548, None))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(MX), Inches(1.30),
                        Inches(7.75), Inches(4.95), cd)
ch = gf.chart
estilo_grafico(ch, leyenda=False)
ch.plots[0].gap_width = 55
ch.plots[0].overlap = 100
serie_color(ch, 0, None, transparente=True)
serie_color(ch, 1, CYAN)
serie_color(ch, 2, VERDE2)
serie_color(ch, 3, ROJO)
etiquetas_es(ch, 1, (70388, None, None, None, 137766), dec=0, pos='ctr', color=BG, size=10, bold=True)
etiquetas_es(ch, 2, (None, 109029, 6897, None, None), dec=0, pos='ctr', color=BG, size=10, bold=True)
etiquetas_es(ch, 3, (None, None, None, 48548, None), dec=0, pos='ctr', color=BG, size=10, bold=True)

rect(s, 8.62, 1.30, 4.13, 2.30, fill=CARD)
txt(s, 8.84, 1.52, 3.7, 0.3, "SI SE QUITA COSTA RICA", 9.5, CYAN, True)
txt(s, 8.84, 1.88, 3.7, 0.62, "USD 28.737", 30, ROJO, True)
txt(s, 8.84, 2.52, 3.7, 0.95,
    "es toda la utilidad que el grupo habría generado en siete meses. "
    "Equivale al 1,7% de los ingresos.", 11, TEXTO)

txt(s, 8.62, 3.85, 4.13, 0.3, "LECTURA", 9.5, CYAN, True)
vinetas(s, 8.62, 4.20, 4.13, [
    ("Costa Rica aporta más utilidad que Colombia con el 8% del ingreso.", VERDE),
    ("Estados Unidos se lleva una tercera parte de la utilidad del grupo.", ROJO),
    ("Esa utilidad de Costa Rica está íntegra por cobrar: es contable, no es caja.", AMBAR),
], size=10.5, gap=7)
pie(s, "Utilidad neta ene–jul 2026 en USD, sin provisión de renta. Colombia 70.388 + Costa Rica 109.029 + Perú 6.897 − Estados Unidos 48.548 = 137.766.")

# ----------------------------------------------- 4 P&G CONSOLIDADO MENSUAL
s = nueva()
cabecera(s, "El consolidado mes a mes",
         "Ingreso y resultado del grupo, USD · la serie mensual del consolidado arranca en abril")

cd = CategoryChartData()
cd.categories = ["Abr", "May", "Jun", "Jul"]
cd.add_series("Ingresos netos", (214383, 233075, 221864, 350895))
cd.add_series("Utilidad neta", (-12139, 8082, -8483, 131564))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MX), Inches(1.30),
                        Inches(7.75), Inches(4.60), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 70
serie_color(ch, 0, CYAN)
serie_color(ch, 1, VERDE2)
eje_cat_abajo(ch)
etiquetas_es(ch, 0, (214383, 233075, 221864, 350895), dec=0, pos='outEnd', color=TEXTO, size=9, bold=False)
etiquetas_es(ch, 1, (-12139, 8082, -8483, 131564), dec=0, pos='outEnd', color=TEXTO, size=9, bold=False)

rect(s, 8.62, 1.30, 4.13, 1.55, fill=CARD)
txt(s, 8.84, 1.50, 3.7, 0.28, "ACUMULADO ENE–JUL", 9.5, CYAN, True)
txt(s, 8.84, 1.82, 3.7, 0.5, "USD 1.651.411", 22, VERDE2, True)
txt(s, 8.84, 2.34, 3.7, 0.4, "Incluye el primer trimestre, que solo existe como acumulado.", 9.5, MUTED)

txt(s, 8.62, 3.10, 4.13, 0.3, "QUÉ MUESTRA", 9.5, CYAN, True)
vinetas(s, 8.62, 3.45, 4.13, [
    ("Abril, mayo y junio se mueven en una banda plana: entre USD 214 mil y 233 mil de ingreso.", CYAN),
    ("El resultado de esos tres meses ronda cero: −12.139, +8.082 y −8.483.", AMBAR),
    ("Julio rompe la serie por un solo cliente nuevo, no por un cambio de tendencia.", VERDE),
], size=10.5, gap=7)
pie(s, "Enero a marzo no aparece en la serie mensual: el consolidado de ese trimestre solo existe como acumulado en el modelo, "
       "y la base todavía no tiene una tasa de conversión validada para esos tres meses. Se reporta dentro del acumulado del año.")

# --------------------------------------------- 5 LA GEOGRAFIA DEL RIESGO
s = nueva()
cabecera(s, "La geografía del riesgo",
         "Dónde se genera el ingreso, dónde queda el resultado y dónde está la caja")

filas = [
    ["País", "Ingresos USD", "% del grupo", "Utilidad USD", "Caja COP MM", "% de la caja"],
    [("Colombia", TEXTO, True, PP_ALIGN.LEFT), "1.469.854", "89,0%",
     ("70.388", VERDE2, False, PP_ALIGN.RIGHT), "174,3", ("15,0%", AMBAR, False, PP_ALIGN.RIGHT)],
    [("Costa Rica", TEXTO, True, PP_ALIGN.LEFT), "134.104", "8,1%",
     ("109.029", VERDE2, False, PP_ALIGN.RIGHT), "—", "—"],
    [("Estados Unidos", TEXTO, True, PP_ALIGN.LEFT), "39.011", "2,4%",
     ("(48.548)", ROJO, False, PP_ALIGN.RIGHT), "988,3", ("85,0%", AMBAR, False, PP_ALIGN.RIGHT)],
    [("Perú", TEXTO, True, PP_ALIGN.LEFT), "8.442", "0,5%",
     ("6.897", VERDE2, False, PP_ALIGN.RIGHT), "—", "—"],
    [("Consolidado", CYAN, True, PP_ALIGN.LEFT), ("1.651.411", CYAN, True, PP_ALIGN.RIGHT),
     ("100%", CYAN, True, PP_ALIGN.RIGHT), ("137.766", CYAN, True, PP_ALIGN.RIGHT),
     ("1.162,6", CYAN, True, PP_ALIGN.RIGHT), ("100%", CYAN, True, PP_ALIGN.RIGHT)],
]
tabla(s, MX, 1.30, [3.05, 1.95, 1.55, 1.95, 1.85, 1.85], filas, alto_fila=0.42, alto_head=0.36, size=11)

txt(s, MX, 4.15, CW, 0.3, "LA ASIMETRÍA", 9.5, CYAN, True)
vinetas(s, MX, 4.52, 6.0, [
    ("Colombia genera 9 de cada 10 pesos de ingreso y se queda con el 15% de la caja del grupo.", ROJO),
    ("Estados Unidos concentra el 85% de la caja, aporta el 2,4% del ingreso y pierde plata.", ROJO),
], size=10.5, gap=8)
vinetas(s, 6.85, 4.52, 5.90, [
    ("Costa Rica es hoy el mejor negocio del grupo por margen, con un solo cliente.", VERDE),
    ("Mover caja de Estados Unidos a Colombia es la palanca más rápida que existe.", VERDE),
], size=10.5, gap=8)
pie(s, "Ingresos y utilidad neta acumulados ene–jul 2026 en USD. Caja al 31 de julio; la de Estados Unidos (USD 302.332) convertida a TRM 3.268,95. "
       "Costa Rica y Perú operan sin caja propia relevante.")

# ------------------------------------------------------- 6 JULIO DEL GRUPO
s = nueva()
cabecera(s, "Julio del grupo — el mejor mes del año, por una sola factura",
         "Consolidado de julio 2026 frente a junio · USD")

kx, kw, kg = MX, 2.93, 0.16
kpi(s, kx + 0 * (kw + kg), 1.25, kw, 1.42, "350.895", "Ingresos julio", "vs 221.864 en junio · +58%", VERDE2, 26)
kpi(s, kx + 1 * (kw + kg), 1.25, kw, 1.42, "132.314", "EBITDA julio", "vs pérdida operativa en junio", VERDE2, 26)
kpi(s, kx + 2 * (kw + kg), 1.25, kw, 1.42, "131.564", "Utilidad neta", "vs −8.483 en junio", VERDE2, 26)
kpi(s, kx + 3 * (kw + kg), 1.25, kw, 1.42, "134.104", "One-off Costa Rica", "el 38% del ingreso del mes", AMBAR, 26)

txt(s, MX, 3.02, 6.0, 0.3, "LO QUE PASÓ", 9.5, CYAN, True)
vinetas(s, MX, 3.38, 6.0, [
    ("Costa Rica facturó por primera vez: USD 134.104, un cliente nuevo sin histórico previo.", VERDE),
    ("Colombia bajó a COP 679,9 MM de ingreso, su mes más bajo del año.", AMBAR),
    ("Sin el one-off, el grupo habría cerrado julio en USD 216.791 — en línea con abril, mayo y junio.", MUTED),
], size=10.5, gap=8)

txt(s, 6.85, 3.02, 5.90, 0.3, "LO QUE HAY QUE VIGILAR", 9.5, CYAN, True)
vinetas(s, 6.85, 3.38, 5.90, [
    ("Los USD 134.104 de Costa Rica están íntegros por cobrar: el resultado mejoró, la caja no.", ROJO),
    ("El margen EBITDA de Colombia quedó en 4,3%, el tercero más bajo de los siete meses.", ROJO),
    ("Ese cliente no tiene histórico: no se puede asumir como MRR estable hasta ver recurrencia.", AMBAR),
], size=10.5, gap=8)
pie(s, "Consolidado en USD; Colombia convertida a TRM 3.268,95. El EBITDA consolidado de julio no incluye provisión de renta.")

# ------------------------------------------------------- 7 COLOMBIA MOTOR
s = nueva()
cabecera(s, "Colombia: el motor, con margen delgado",
         "Estado de resultados de julio 2026 · COP millones · 89% del ingreso del grupo")

filas = [
    ["Concepto", "COP MM", "% ingresos"],
    ["Ingresos netos", ("679,9", TEXTO, True, PP_ALIGN.RIGHT), "100,0%"],
    ["Costo del servicio", "321,9", "47,3%"],
    [("Utilidad bruta", VERDE, True, PP_ALIGN.LEFT), ("358,0", VERDE, True, PP_ALIGN.RIGHT),
     ("52,7%", VERDE, True, PP_ALIGN.RIGHT)],
    ["Gastos de administración", "107,9", "15,9%"],
    ["Gastos de ventas", "117,8", "17,3%"],
    ["Gastos de operación", "102,9", "15,1%"],
    [("EBITDA", VERDE, True, PP_ALIGN.LEFT), ("29,3", VERDE, True, PP_ALIGN.RIGHT),
     ("4,3%", VERDE, True, PP_ALIGN.RIGHT)],
    [("Utilidad neta", CYAN, True, PP_ALIGN.LEFT), ("24,7", CYAN, True, PP_ALIGN.RIGHT),
     ("3,6%", CYAN, True, PP_ALIGN.RIGHT)],
]
tabla(s, MX, 1.30, [3.55, 1.55, 1.45], filas, alto_fila=0.375, alto_head=0.34, size=11)

cd = CategoryChartData()
cd.categories = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]
cd.add_series("EBITDA COP MM", (70.0, 46.1, -22.2, 4.8, 113.0, 35.0, 29.3))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.30), Inches(1.42),
                        Inches(5.45), Inches(2.60), cd)
ch = gf.chart
estilo_grafico(ch, leyenda=False)
ch.plots[0].gap_width = 60
serie_color(ch, 0, VERDE2)
for i, v in enumerate((70.0, 46.1, -22.2, 4.8, 113.0, 35.0, 29.3)):
    if v < 0:
        punto_color(ch, 0, i, ROJO)
eje_cat_abajo(ch)
etiquetas_es(ch, 0, (70.0, 46.1, -22.2, 4.8, 113.0, 35.0, 29.3), dec=1, pos='outEnd', color=TEXTO, size=9, bold=False)
txt(s, 7.30, 1.30, 5.45, 0.26, "EBITDA MENSUAL 2026 · COP MM", 9.5, CYAN, True)

txt(s, MX, 4.92, CW, 0.3, "LECTURA DEL MES", 9.5, CYAN, True)
vinetas(s, MX, 5.28, 6.0, [
    ("El margen bruto de 52,7% aguanta, pero la estructura se come casi todo lo que queda.", AMBAR),
    ("Faltan COP 99,6 MM de provisión de renta acumulada del año que la contabilidad aún no registra.", ROJO),
], size=10.5, gap=8)
vinetas(s, 6.85, 5.28, 5.90, [
    ("EBITDA acumulado ene–jul: COP 276,0 MM, margen de 5,4%.", VERDE),
    ("Dos meses concentran el 66% del EBITDA del año: mayo (113,0) y enero (70,0).", AMBAR),
], size=10.5, gap=8)
pie(s, "Colombia sola, COP millones, estados financieros oficiales de julio. La depreciación va dentro del costo del servicio, así que EBITDA y EBIT coinciden.")

# --------------------------------------------- 8 MARGEN INESTABLE + ESTRUCTURA
s = nueva()
cabecera(s, "El margen no es bajo: es inestable",
         "Colombia · margen EBITDA mensual contra el peso de la estructura sobre el ingreso")

cd = CategoryChartData()
cd.categories = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]
cd.add_series("Costo del servicio", (.3896, .4613, .5789, .5051, .4301, .4637, .4735))
cd.add_series("Administración", (.1612, .1419, .1802, .1522, .1413, .1494, .1588))
cd.add_series("Ventas", (.2149, .2333, .1774, .1733, .1455, .1954, .1732))
cd.add_series("Operación", (.1199, .1015, .0918, .1630, .1464, .1449, .1514))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(MX), Inches(1.32),
                        Inches(7.75), Inches(4.55), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 60
ch.plots[0].overlap = 100
for i, c in enumerate([CYAN, RGBColor(0x3B, 0x82, 0xF6), VERDE2, AMBAR]):
    serie_color(ch, i, c)
ch.value_axis.tick_labels.number_format = '0%'
ch.value_axis.tick_labels.number_format_is_linked = False

rect(s, 8.62, 1.32, 4.13, 1.30, fill=CARD)
txt(s, 8.84, 1.52, 3.7, 0.28, "LA ESTRUCTURA CONSUME", 9.5, CYAN, True)
txt(s, 8.84, 1.82, 3.7, 0.5, "94,6%", 30, AMBAR, True)
txt(s, 8.84, 2.32, 3.7, 0.28, "del ingreso, en promedio del año", 9.5, MUTED)

txt(s, 8.62, 2.90, 4.13, 0.3, "DÓNDE ESTÁ LA TENSIÓN", 9.5, CYAN, True)
vinetas(s, 8.62, 3.26, 4.13, [
    ("En marzo el costo y el gasto superaron el ingreso: 102,8%, pérdida operativa de 22,2 MM.", ROJO),
    ("Abril quedó al filo, en 99,4%.", AMBAR),
    ("Operación pasó de ~10% del ingreso en el primer trimestre a ~15% desde abril: es la línea que más se movió.", AMBAR),
    ("El rango del margen EBITDA va de −2,8% a 13,7%: el problema es la varianza, no el nivel.", CYAN),
], size=10.5, gap=6)
pie(s, "Colombia sola, cada línea como porcentaje del ingreso neto del mes. Cuando las cuatro bandas superan el 100%, el mes cierra en pérdida operativa.")

# ---------------------------------------------------- 9 CUMPLIMIENTO PPTO
s = nueva()
cabecera(s, "Cumplimiento contra presupuesto — enero a julio",
         "Grupo consolidado contra presupuesto combinado · USD · perímetro verificado")

cd = CategoryChartData()
cd.categories = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]
cd.add_series("Presupuesto", (186232, 354321, 396239, 359191, 360136, 259076, 283850))
cd.add_series("Real del grupo", (169420, 206566, 218716, 210306, 234524, 211132, 327163))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MX), Inches(1.32),
                        Inches(7.75), Inches(4.30), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 60
serie_color(ch, 0, BORDE)
serie_color(ch, 1, CYAN)
etiquetas_es(ch, 1, (169420, 206566, 218716, 210306, 234524, 211132, 327163), dec=0, pos='outEnd', color=TEXTO, size=8.5, bold=False)

pcts = ["91%", "58%", "55%", "59%", "65%", "82%", "115%"]
for i, p in enumerate(pcts):
    col = VERDE2 if i == 6 else (AMBAR if i in (0, 5) else ROJO)
    txt(s, MX + 0.62 + i * 1.045, 5.68, 0.95, 0.28, p, 12, col, True, PP_ALIGN.CENTER)
txt(s, MX, 5.98, 7.75, 0.25, "cumplimiento del mes", 9, MUTED, align=PP_ALIGN.CENTER)

rect(s, 8.62, 1.32, 4.13, 1.42, fill=CARD)
txt(s, 8.84, 1.52, 3.7, 0.28, "ACUMULADO DEL AÑO", 9.5, CYAN, True)
txt(s, 8.84, 1.84, 3.7, 0.45, "71,8%", 28, AMBAR, True)
txt(s, 8.84, 2.32, 3.7, 0.32, "USD 1.577.827 contra 2.199.045 · brecha de 621.218", 9.5, MUTED)

txt(s, 8.62, 3.00, 4.13, 0.3, "CÓMO LEERLO", 9.5, CYAN, True)
vinetas(s, 8.62, 3.36, 4.13, [
    ("El negocio recurrente lleva siete meses entre 55% y 91%: la brecha es estructural, no de un mes.", ROJO),
    ("Julio cumple 115% solo por el one-off de Costa Rica; sin él habría cerrado en 68%.", AMBAR),
    ("Enero–abril se mide contra el presupuesto original (USD 4.555.726 anuales) y mayo–julio contra el revisado (4.067.468).", MUTED),
], size=10.5, gap=6)
pie(s, "Corrección de perímetro respecto de la versión anterior: la serie real es del GRUPO, no de Colombia sola. Verificado — julio = "
       "USD 193.059 recurrente del grupo + 134.104 del one-off de Costa Rica = 327.163. La cifra de 1.577.827 mide recurrente más one-off, "
       "y por eso difiere del ingreso neto consolidado de 1.651.411; ambas son correctas y miden cosas distintas.")

# ---------------------------------------------------- 10 LIQUIDEZ Y RUNWAY
s = nueva()
cabecera(s, "Liquidez consolidada y runway",
         "Caja del grupo al 31 de julio de 2026 · COP millones")

kx, kw, kg = MX, 2.93, 0.16
kpi(s, kx + 0 * (kw + kg), 1.25, kw, 1.42, "1.162,6", "Caja consolidada", "COP MM al 31 de julio", CYAN, 27)
kpi(s, kx + 1 * (kw + kg), 1.25, kw, 1.42, "8,8", "Meses de runway", "al burn promedio may–jul", AMBAR, 27)
kpi(s, kx + 2 * (kw + kg), 1.25, kw, 1.42, "132,1", "Quema mensual", "COP MM · promedio 3 meses", ROJO, 27)
kpi(s, kx + 3 * (kw + kg), 1.25, kw, 1.42, "85%", "Caja en EE.UU.", "donde se genera el 2,4% del ingreso", ROJO, 27)

cd = CategoryChartData()
cd.categories = ["Caja jul-26"]
cd.add_series("Colombia", (174.3,))
cd.add_series("Estados Unidos", (988.3,))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(MX), Inches(3.05),
                        Inches(7.75), Inches(1.75), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 90
ch.plots[0].overlap = 100
serie_color(ch, 0, VERDE2)
serie_color(ch, 1, CYAN)
etiquetas_es(ch, 0, (174.3,), dec=1, pos='ctr', color=BG, size=11, bold=True)
etiquetas_es(ch, 1, (988.3,), dec=1, pos='ctr', color=BG, size=11, bold=True)

txt(s, MX, 5.05, CW, 0.3, "EL PROBLEMA NO ES CUÁNTA CAJA HAY, ES DÓNDE ESTÁ", 9.5, CYAN, True)
vinetas(s, MX, 5.42, 6.0, [
    ("Colombia opera con COP 174,3 MM: cerca de un mes de su propio gasto de estructura.", ROJO),
    ("Estados Unidos guarda COP 988,3 MM (USD 302.332) y consume USD 12.128 al mes.", AMBAR),
], size=10.5, gap=8)
vinetas(s, 6.85, 5.42, 5.90, [
    ("El runway de 8,8 meses es del grupo; el de Colombia sola es mucho más corto.", ROJO),
    ("Girar caja desde Estados Unidos es la palanca inmediata para el bache de agosto.", VERDE),
], size=10.5, gap=8)
pie(s, "Caja de Estados Unidos convertida a TRM 3.268,95. El runway usa el burn promedio de mayo a julio (COP 132,1 MM/mes); "
       "julio no aportó quema y mejora el promedio trimestral.")

# ------------------------------------------------------ 11 PUENTE DE CAJA
s = nueva()
cabecera(s, "Puente de caja de julio — Colombia",
         "De COP 115,4 a 174,3 millones · el aumento viene de cobrar, no de operar")

cd = CategoryChartData()
cd.categories = ["Caja inicial", "Utilidad neta", "Cartera", "Otros deudores", "Proveedores",
                 "Impuestos", "Laborales", "Inversiones", "Tarjetas", "Socios", "Caja final"]
cd.add_series("Base",    (None, 115.4, 140.1, 187.5, 139.7, 135.6, 135.6, 153.0, 153.0, 174.1, None))
cd.add_series("Saldo",   (115.4, None, None, None, None, None, None, None, None, None, 174.3))
cd.add_series("Entradas", (None, 24.7, 111.0, None, None, None, 18.1, None, 21.1, 0.2, None))
cd.add_series("Salidas",  (None, None, None, 31.8, 79.6, 4.1, None, 0.7, None, None, None))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(MX), Inches(1.32),
                        Inches(12.20), Inches(3.95), cd)
ch = gf.chart
estilo_grafico(ch, leyenda=False)
ch.plots[0].gap_width = 45
ch.plots[0].overlap = 100
serie_color(ch, 0, None, transparente=True)
serie_color(ch, 1, CYAN)
serie_color(ch, 2, VERDE2)
serie_color(ch, 3, ROJO)
etiquetas_es(ch, 1, (115.4, None, None, None, None, None, None, None, None, None, 174.3), dec=1, pos='ctr', color=BG, size=9, bold=True)
_e = (None, 24.7, 111.0, None, None, None, 18.1, None, 21.1, 0.2, None)
_sal = (None, None, None, 31.8, 79.6, 4.1, None, 0.7, None, None, None)
_cl = lambda vs: [BG if (v or 0) >= 10 else TEXTO for v in vs]
etiquetas_es(ch, 2, _e, dec=1, pos='ctr', color=_cl(_e), size=9, bold=True)
etiquetas_es(ch, 3, _sal, dec=1, pos='ctr', color=_cl(_sal), size=9, bold=True)

txt(s, MX, 5.48, CW, 0.3, "LO QUE DICE EL PUENTE", 9.5, CYAN, True)
vinetas(s, MX, 5.84, 6.0, [
    ("La caja creció COP 58,8 MM: primera reposición desde marzo.", VERDE),
    ("El motor fue el cobro de cartera (+111,0), no la generación operativa.", AMBAR),
], size=10.5, gap=8)
vinetas(s, 6.85, 5.84, 5.90, [
    ("Se pagaron proveedores por 79,6 y se redujeron otros deudores en 31,8.", MUTED),
    ("Es una mejora de capital de trabajo, no un cambio en la rentabilidad del negocio.", ROJO),
], size=10.5, gap=8)
pie(s, "Cifras en COP millones. El puente concilia al peso contra el balance auditado de julio.")

# ------------------------------------------------ 12 CARTERA (LAMINA NUEVA)
s = nueva()
cabecera(s, "Cartera y capital de trabajo",
         "Colombia · junio contra julio 2026 · la cobranza mejoró pero la calidad se deterioró")

filas = [
    ["Indicador", "Junio", "Julio", "Variación"],
    ["Cartera total (COP MM)", "1.799,4", "1.647,0", ("−8,5%", VERDE2, False, PP_ALIGN.RIGHT)],
    ["Cartera sana (COP MM)", "1.503,2", "961,7", ("−36,0%", ROJO, False, PP_ALIGN.RIGHT)],
    ["Cartera vencida (COP MM)", "579,0", "685,3", ("+18,4%", ROJO, False, PP_ALIGN.RIGHT)],
    [("Vencida sobre el total", TEXTO, True, PP_ALIGN.LEFT), "32,2%",
     ("41,6%", ROJO, True, PP_ALIGN.RIGHT), ("+9,4 pp", ROJO, False, PP_ALIGN.RIGHT)],
    ["Días de cartera (DSO)", "70,8", ("58,0", VERDE2, True, PP_ALIGN.RIGHT),
     ("−12,8 días", VERDE2, False, PP_ALIGN.RIGHT)],
    ["Mora mayor a 60 días", "20,3%", ("7,0%", VERDE2, False, PP_ALIGN.RIGHT),
     ("−13,3 pp", VERDE2, False, PP_ALIGN.RIGHT)],
    ["Cartera en cobro jurídico (COP MM)", "61,9", "57,3", ("−7,5%", VERDE2, False, PP_ALIGN.RIGHT)],
    ["Castigos sobre cartera", "0,25%", "0,22%", ("−0,03 pp", VERDE2, False, PP_ALIGN.RIGHT)],
]
tabla(s, MX, 1.30, [3.75, 1.50, 1.50, 1.60], filas, alto_fila=0.365, alto_head=0.34, size=10.5)

rect(s, 9.40, 1.30, 3.35, 1.55, fill=CARD)
txt(s, 9.62, 1.50, 2.95, 0.28, "COBRADO EN JULIO", 9.5, CYAN, True)
txt(s, 9.62, 1.82, 2.95, 0.5, "111,0", 30, VERDE2, True)
txt(s, 9.62, 2.34, 2.95, 0.32, "COP MM · el motor de la caja del mes", 9.5, MUTED)

rect(s, 9.40, 3.00, 3.35, 1.55, fill=CARD)
txt(s, 9.62, 3.20, 2.95, 0.28, "SIGUE POR COBRAR", 9.5, CYAN, True)
txt(s, 9.62, 3.52, 2.95, 0.5, "685,3", 30, ROJO, True)
txt(s, 9.62, 4.04, 2.95, 0.32, "COP MM vencidos · 4 veces la caja de Colombia", 9.5, MUTED)

txt(s, MX, 4.90, 8.85, 0.3, "LA LECTURA INCÓMODA", 9.5, CYAN, True)
vinetas(s, MX, 5.26, 8.85, [
    ("El DSO mejoró 13 días y la mora mayor a 60 días bajó de 20,3% a 7,0%: la gestión de cobro funcionó.", VERDE),
    ("Pero la cartera vencida creció 18% mientras la total bajaba: se cobró lo sano y quedó lo difícil.", ROJO),
    ("La cartera vencida (685,3) es casi cuatro veces la caja disponible de Colombia (174,3).", ROJO),
], size=10.5, gap=8)
pie(s, "Fuente: reporte de cartera y liquidez de julio, convertido a COP con TRM 3.505,27. Dos indicadores del reporte "
       "(índice de morosidad y concentración de cartera) quedaron sin cargar por saltos no explicados frente a junio.")

# -------------------------------------------------------- 13 AGOSTO NO CIERRA
s = nueva()
cabecera(s, "Agosto no cierra",
         "Proyección de caja de Colombia agosto–diciembre 2026 · COP millones")

cd = CategoryChartData()
cd.categories = ["Ago", "Sep", "Oct", "Nov", "Dic"]
cd.add_series("Caja fin de mes", (-867.7, -554.4, -181.5, 200.0, 200.0))
cd.add_series("Caja mínima operativa", (200.0, 200.0, 200.0, 200.0, 200.0))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MX), Inches(1.32),
                        Inches(7.75), Inches(4.35), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 65
serie_color(ch, 0, ROJO)
serie_color(ch, 1, BORDE)
for i in (3, 4):                      # nov y dic cierran EN el minimo, no en deficit
    punto_color(ch, 0, i, VERDE2)
eje_cat_abajo(ch)
etiquetas_es(ch, 0, (-867.7, -554.4, -181.5, 200.0, 200.0), dec=1, pos='outEnd', color=TEXTO, size=9.5, bold=False)

rect(s, 8.62, 1.32, 4.13, 1.45, fill=CARD)
txt(s, 8.84, 1.52, 3.7, 0.28, "FALTANTE MÁXIMO", 9.5, CYAN, True)
txt(s, 8.84, 1.84, 3.7, 0.48, "867,7", 30, ROJO, True)
txt(s, 8.84, 2.34, 3.7, 0.32, "COP MM en agosto, antes de aplicar palancas", 9.5, MUTED)

txt(s, 8.62, 3.05, 4.13, 0.3, "POR QUÉ PASA", 9.5, CYAN, True)
vinetas(s, 8.62, 3.41, 4.13, [
    ("El bache es de capital de trabajo, no una pérdida operativa.", CYAN),
    ("El presupuesto proyecta cartera a 84 días sobre un ingreso de agosto 2,3 veces el de julio.", AMBAR),
    ("Girando toda la capacidad disponible de Estados Unidos (750,4 MM, dejándole 6 meses de reserva), agosto todavía queda 1.067,7 MM corto.", ROJO),
], size=10.5, gap=7)
pie(s, "ADVERTENCIA: el presupuesto de agosto a diciembre tiene el costo de ventas congelado y no fue rebasificado contra los reales. "
       "La utilidad bruta del semestre está sobreestimada en aproximadamente COP 2.375,6 MM, por lo que el faltante de caja mostrado es un piso, no un techo.")

# ----------------------------------------------------- 14 PETICION A LA JUNTA
s = nueva()
cabecera(s, "Lo que pedimos a la junta",
         "Tres decisiones · cualquiera de las dos primeras cubre por sí sola el déficit de agosto")

cd = CategoryChartData()
cd.categories = ["Recuperar 25% de\notros deudores", "Devolución del\nanticipo de impuestos",
                 "Capitalizar el\npasivo con socios", "Déficit máximo\nde agosto"]
cd.add_series("COP MM", (882.4, 866.1, 2641.8, 867.7))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(MX), Inches(1.32),
                        Inches(6.30), Inches(4.54), cd)
ch = gf.chart
estilo_grafico(ch, leyenda=False)
ch.plots[0].gap_width = 55
serie_color(ch, 0, VERDE2)
punto_color(ch, 0, 3, ROJO)
etiquetas_es(ch, 0, (882.4, 866.1, 2641.8, 867.7), dec=1, pos='outEnd', color=TEXTO, size=10, bold=False)

bx = 7.15
for i, (t, d, col) in enumerate([
    ("Cobrar", "Hay COP 3.529,7 MM en otros deudores y 1.688,5 en cartera: 4,5 veces la caja total del grupo. "
               "Recuperar el 25% libera 882 MM y borra el bache de agosto.", VERDE2),
    ("Radicar la devolución", "COP 866,1 MM de anticipo de impuestos están hoy financiando gratis a la DIAN. "
                              "Radicar la devolución es trámite, no negociación.", CYAN),
    ("Definir el pasivo con socios", "COP 2.641,8 MM: ¿se capitaliza o sigue como deuda? Es el 77% del pasivo total "
                                     "y como línea de fondeo ya está agotada.", AMBAR),
]):
    y = 1.32 + i * 1.57
    rect(s, bx, y, 5.60, 1.42, fill=CARD)
    rect(s, bx, y, 0.05, 1.42, fill=col, line=None)
    txt(s, bx + 0.24, y + 0.18, 5.15, 0.30, t.upper(), 12, col, True)
    txt(s, bx + 0.24, y + 0.54, 5.15, 0.78, d, 10, TEXTO)

pie(s, "Cifras en COP millones al 31 de julio de 2026.", y=6.25)

# ------------------------------------------------------------- A1 ANEXO PPTO
s = nueva()
cabecera(s, "Anexo · cómo se construye el presupuesto de julio",
         "Presupuesto revisado de mayo, consolidado · USD")

cd = CategoryChartData()
cd.categories = ["MRR base", "Logos nuevos", "Innovación", "Up/Cross", "One-off",
                 "MRR Intl.", "Subtotal bruto", "Churn", "Contracción", "Ppto julio"]
cd.add_series("Base", (None, 216802, 250046, 286260, 305816, 311416, None, 287583, 283850, None))
cd.add_series("Total", (None, None, None, None, None, None, 315583, None, None, 283850))
cd.add_series("Suma", (216802, 33244, 36214, 19556, 5600, 4167, None, None, None, None))
cd.add_series("Resta", (None, None, None, None, None, None, None, 28000, 3733, None))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(MX), Inches(1.32),
                        Inches(12.20), Inches(4.05), cd)
ch = gf.chart
estilo_grafico(ch, leyenda=False)
ch.plots[0].gap_width = 45
ch.plots[0].overlap = 100
serie_color(ch, 0, None, transparente=True)
serie_color(ch, 1, CYAN)
serie_color(ch, 2, VERDE2)
serie_color(ch, 3, ROJO)
# tramo delgado (<15.000) => etiqueta clara, porque se sale de la barra
CLARO = lambda vs: [BG if (v or 0) >= 15000 else TEXTO for v in vs]
etiquetas_es(ch, 1, (None, None, None, None, None, None, 315583, None, None, 283850), dec=0, pos='ctr', color=BG, size=9, bold=True)
_v2 = (216802, 33244, 36214, 19556, 5600, 4167, None, None, None, None)
_v3 = (None, None, None, None, None, None, None, 28000, 3733, None)
etiquetas_es(ch, 2, _v2, dec=0, pos='ctr', color=CLARO(_v2), size=9, bold=True)
etiquetas_es(ch, 3, _v3, dec=0, pos='ctr', color=CLARO(_v3), size=9, bold=True)

txt(s, MX, 5.58, CW, 0.3, "CÓMO SE LEE", 9.5, CYAN, True)
vinetas(s, MX, 5.94, 6.0, [
    ("El MRR base es el 76,4% del presupuesto; logos nuevos e innovación aportan 24,4% entre los dos.", CYAN),
    ("Contra el real recurrente del grupo (193.059), la brecha del mes es de 90.791.", ROJO),
], size=10.5, gap=8)
vinetas(s, 6.85, 5.94, 5.90, [
    ("Con el one-off de Costa Rica, julio cerró 43.313 por encima del presupuesto.", VERDE),
    ("Pendiente clasificar la brecha entre diferimiento y pérdida permanente.", AMBAR),
], size=10.5, gap=8)
pie(s, "PENDIENTE DE CORRECCIÓN EN EL LIBRO: churn y contracción devuelven cero por un rango mal referenciado; mayo y junio tienen el mismo hueco. "
       "Las cifras mostradas son las del modelo revisado; la cascada cierra exacto contra su total.")

# ------------------------------------------------------- A2 ANEXO GASTOS AREA
s = nueva()
cabecera(s, "Anexo · gastos por área — Colombia",
         "Junio contra julio 2026 · COP millones · reconstruido desde el balance de prueba del ERP")

cd = CategoryChartData()
cd.categories = ["Administración", "Comercial", "Proveedores", "TI", "Financieros"]
cd.add_series("Junio", (112.2, 146.8, 188.5, 268.8, 5.5))
cd.add_series("Julio", (95.2, 120.9, 114.8, 264.2, 5.4))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MX), Inches(1.32),
                        Inches(7.75), Inches(4.30), cd)
ch = gf.chart
estilo_grafico(ch)
ch.plots[0].gap_width = 60
serie_color(ch, 0, BORDE)
serie_color(ch, 1, CYAN)
etiquetas_es(ch, 0, (112.2, 146.8, 188.5, 268.8, 5.5), dec=1, pos='outEnd', color=MUTED, size=9, bold=False)
etiquetas_es(ch, 1, (95.2, 120.9, 114.8, 264.2, 5.4), dec=1, pos='outEnd', color=TEXTO, size=9, bold=False)

rect(s, 8.62, 1.32, 4.13, 1.42, fill=CARD)
txt(s, 8.84, 1.52, 3.7, 0.28, "GASTO TOTAL DE JULIO", 9.5, CYAN, True)
txt(s, 8.84, 1.84, 3.7, 0.48, "600,4", 30, VERDE2, True)
txt(s, 8.84, 2.32, 3.7, 0.32, "COP MM · contra 721,8 en junio", 9.5, MUTED)

txt(s, 8.62, 3.02, 4.13, 0.3, "QUÉ SE MOVIÓ", 9.5, CYAN, True)
vinetas(s, 8.62, 3.38, 4.13, [
    ("Proveedores baja de 188,5 a 114,8: es la caída más grande del mes.", VERDE),
    ("TI se mantiene prácticamente igual (268,8 → 264,2): es el rubro más rígido.", AMBAR),
    ("Comercial cede 26 MM y administración 17 MM.", VERDE),
], size=10.5, gap=7)
pie(s, "Julio calculado por diferencia: acumulado enero–julio del balance de prueba del ERP menos el acumulado enero–junio ya conocido. "
       "Tres partidas quedaron sin asignar por falta de trazabilidad (diversos administrativos, AWS y viajes de operación), "
       "por lo que el gasto de julio mostrado es un piso.")

prs.save(SALIDA)
print("OK ->", SALIDA)
print("Diapositivas:", len(prs.slides.__iter__.__self__._sldIdLst))
