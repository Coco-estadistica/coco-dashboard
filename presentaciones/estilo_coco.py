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


