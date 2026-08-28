# -*- coding: utf-8 -*-
"""
Genera "Presentacion Arquitectura Dashboard COCO.pptx".

Documenta el tablero para quien lo tenga que operar o auditar:
componentes y lenguajes, anatomia del HTML, los scripts, la logica de
ejecucion, el linaje del dato y las cifras pendientes con contabilidad.

Reutiliza el formato COCO de estilo_coco.py (extraido de
generar_presentacion_consolidada.py): fondo #142433, Arial, titulo verde.
"""
import os

from estilo_coco import (  # noqa: F401
    Presentation, PLANTILLA, BASE,
    RGBColor, Inches, Pt, MSO_SHAPE, PP_ALIGN, MSO_ANCHOR,
    BG, CARD, CARD2, BORDE, TEXTO, MUTED, VERDE, VERDE2, CYAN, ROJO, AMBAR,
    MX, CW, n0, txt, rect, cabecera, pie, vinetas, kpi, tabla,
)

SALIDA = os.path.join(BASE, "Presentacion Arquitectura Dashboard COCO.pptx")

prs = Presentation(PLANTILLA)
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    prs.part.drop_rel(sld.rId)
    sldIdLst.remove(sld)
LAYOUT = next(l for l in prs.slide_layouts if l.name == "COCO")


def nueva():
    return prs.slides.add_slide(LAYOUT)


def flecha(slide, x, y, w=0.20, h=0.22, color=CYAN, abajo=False):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.DOWN_ARROW if abajo else MSO_SHAPE.RIGHT_ARROW,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def caja(slide, x, y, w, h, titulo, lineas, tcolor=CYAN, fill=CARD,
         tsize=10, lsize=8.6, borde=BORDE):
    """Caja de diagrama: titulo arriba, lineas debajo."""
    rect(slide, x, y, w, h, fill=fill, line=borde)
    txt(slide, x + 0.14, y + 0.13, w - 0.28, 0.24, titulo, tsize, tcolor, True)
    if lineas:
        txt(slide, x + 0.14, y + 0.42, w - 0.28, h - 0.52,
            "\n".join(lineas), lsize, MUTED, space=3)


# ============================================================ 1 PORTADA
s = nueva()
txt(s, MX, 1.95, CW, 0.9, "Cockpit Financiero COCO", 42, VERDE, True)
txt(s, MX, 2.92, CW, 0.6, "Arquitectura del tablero", 25, TEXTO)
rect(s, MX, 3.72, 2.6, 0.035, fill=CYAN, line=None)
txt(s, MX, 4.02, 10.0, 1.1,
    "Componentes, lenguajes y scripts · Lógica de ejecución · De dónde sale cada cifra\n"
    "Formas de cargar la información · Cifras pendientes de revisar con contabilidad",
    13, MUTED, space=5)
txt(s, MX, 6.30, CW, 0.5,
    "Documento técnico y de control · 25 de agosto de 2026 · "
    "Corte de datos: julio de 2026 · Versión del tablero: JULIO-2026-R4",
    9, MUTED)

# ================================================ 2 EL TABLERO EN UNA PAGINA
s = nueva()
cabecera(s, "El tablero en una página",
         "Qué es, de qué está hecho y qué lo hace distinto de un Excel compartido")
kx, kw, kg = MX, 2.93, 0.16
kpi(s, kx + 0 * (kw + kg), 1.22, kw, 1.45, "1", "Archivo HTML",
    "autocontenido · sin internet · 1,17 MB", VERDE2, 30)
kpi(s, kx + 1 * (kw + kg), 1.22, kw, 1.45, "13", "Pestañas",
    "de Resumen a Catálogo", VERDE2, 30)
kpi(s, kx + 2 * (kw + kg), 1.22, kw, 1.45, "27", "Scripts Python",
    "4.690 líneas de automatización", CYAN, 30)
kpi(s, kx + 3 * (kw + kg), 1.22, kw, 1.45, "1.852", "Filas de hechos",
    "BD_Indicadores · formato largo", CYAN, 30)

txt(s, MX, 3.00, CW, 0.3, "CÓMO ESTÁ PENSADO", 9.5, CYAN, True)
vinetas(s, MX, 3.38, 12.1, [
    ("Un solo archivo HTML que se abre en el navegador y lee un Excel. No hay servidor de aplicación, ni base de datos, ni licencias.", VERDE2),
    ("Una sola fuente de verdad: BD_MAESTRA_COCO.xlsx. El tablero abre exactamente 10 de sus 14 hojas y ninguna otra.", VERDE2),
    ("Multipaís y multimoneda: Colombia, EE.UU., Perú y Costa Rica; COP, USD y PEN, con conversión explícita y trazable.", CYAN),
    ("Un dato que no llegó se marca «s/d», nunca cero. La diferencia entre «no reportado» y «reportó cero» está en el código, no en la interpretación de quien lee.", AMBAR),
    ("Todo lo pesado —homologar, corregir, conciliar— ocurre antes, en Python. El navegador solo lee y dibuja.", CYAN),
], size=11.2, gap=10)
pie(s, "Fuente: Dashboard_COCO.html, README.md y BD_MAESTRA_COCO.xlsx del proyecto COCO_Dashboard_Cloud.")

# ================================================ 3 COMPONENTES Y LENGUAJES
s = nueva()
cabecera(s, "Componentes y lenguajes",
         "Qué hay en el proyecto, en qué está escrito y para qué sirve")
anchos = [3.05, 1.85, 1.15, 6.15]


def izq(t):
    return (t, TEXTO, False, PP_ALIGN.LEFT)


filas = [
    ["Componente", "Lenguaje", "Tamaño", ("Qué hace", CYAN, True, PP_ALIGN.LEFT)],
    ["Dashboard_COCO.html", "HTML · CSS · JS · SVG", "2.776 lín.", izq("El tablero completo: interfaz, cálculo y gráficos en un archivo")],
    ["SheetJS (XLSX) embebido", "JavaScript minificado", "26 lín.", izq("Lee el .xlsx dentro del navegador; va incrustado, no se descarga")],
    ["apps_script/Code.gs", "Google Apps Script", "400 lín.", izq("Capa de nube: entrega los datos y recibe los cargues por área")],
    ["migracion/*.py (22)", "Python 3 · openpyxl", "3.220 lín.", izq("Construye, carga, corrige y concilia la base maestra")],
    ["modelo/*.py (2)", "Python 3 · openpyxl", "493 lín.", izq("Extrae Burn & Runway y el P&G del modelo presupuestal V3")],
    ["generar_*.py (3)", "Python 3 · python-pptx", "977 lín.", izq("Publica el tablero para compartir y arma las presentaciones")],
    ["*.bat (3)", "Batch / CMD", "~40 lín.", izq("Los tres botones de uso diario: abrir, revisar y corregir")],
    ["burn_runway.json", "JSON", "14,7 KB", izq("Sello del modelo financiero que alimenta la pestaña 03")],
    ["BD_MAESTRA_COCO.xlsx", "Excel · 14 hojas", "1.852 filas", izq("La única fuente de verdad que el tablero consume")],
    ["BD_TRAZABILIDAD_COCO.xlsx", "Excel · 24 hojas", "986 reg.", izq("Archivo de auditoría. El tablero nunca lo abre")],
]
tabla(s, MX, 1.22, anchos, filas, alto_fila=0.395, size=9.6)

txt(s, MX, 5.90, CW, 0.3, "OCHO LENGUAJES, CADA UNO EN SU SITIO", 9.5, CYAN, True)
rect(s, MX, 6.24, CW, 0.52)
txt(s, MX + 0.20, 6.36, CW - 0.40, 0.32,
    "HTML   ·   CSS   ·   JavaScript (ES6)   ·   SVG   ·   Google Apps Script   ·   "
    "Python 3   ·   Batch / CMD   ·   JSON", 10, TEXTO, anchor=MSO_ANCHOR.MIDDLE)
pie(s, "Cero dependencias de red en el tablero: no hay etiqueta script con src, ni CDN, ni Chart.js. Los gráficos son SVG escrito a mano "
       "(barras, apiladas, torta, líneas con cortes y puentes). Eso es lo que permite abrirlo sin internet y enviarlo por correo.", y=7.00)

# ============================================= 4 ANATOMIA DEL HTML
s = nueva()
cabecera(s, "Anatomía del archivo del tablero",
         "Dashboard_COCO.html · 2.776 líneas · los cuatro bloques que lo componen")
bx, bw, bg = MX, 2.93, 0.16
bloques = [
    ("Líneas 1–9 · 200–327", "ESTRUCTURA", "HTML", [
        "Barra de filtros",
        "Menú de pestañas",
        "Contenedor de láminas",
        "Modales de cargue"], VERDE2),
    ("Líneas 10–199", "PRESENTACIÓN", "CSS", [
        "Tema oscuro con variables",
        "Tarjetas de KPI y tablas",
        "Estados: s/d, n/a, cero",
        "Diseño adaptable"], CYAN),
    ("Líneas 328–353", "LECTURA DE EXCEL", "SheetJS", [
        "XLSX.read() del libro",
        "sheet_to_json a matrices",
        "Va incrustado y minificado",
        "Sin descarga externa"], AMBAR),
    ("Líneas 355–2.774", "LA APLICACIÓN", "JavaScript", [
        "~300 funciones propias",
        "parseWorkbook, render, filtros",
        "13 renderizadores de pestaña",
        "Gráficos SVG a mano"], VERDE2),
]
for i, (rango, etiqueta, leng, items, col) in enumerate(bloques):
    x = bx + i * (bw + bg)
    rect(s, x, 1.25, bw, 2.60)
    txt(s, x + 0.18, 1.42, bw - 0.36, 0.22, rango, 8.2, MUTED)
    txt(s, x + 0.18, 1.68, bw - 0.36, 0.28, etiqueta, 10.5, col, True)
    txt(s, x + 0.18, 2.00, bw - 0.36, 0.30, leng, 15, TEXTO, True)
    txt(s, x + 0.18, 2.45, bw - 0.36, 1.30, "\n".join("· " + t for t in items),
        9, MUTED, space=4)

txt(s, MX, 4.15, CW, 0.3, "LAS 13 PESTAÑAS", 9.5, CYAN, True)
tab_txt = ("01 Resumen   ·   02 Consolidado   ·   03 Burn & Runway   ·   04 Finanzas & P&L   ·   "
           "05 P&G Colombia   ·   06 Ingresos & MRR   ·   07 Cartera & Liquidez\n"
           "08 Gastos Colombia   ·   09 Gastos Consolidado   ·   10 Ventas & Comercial   ·   "
           "11 Marketing   ·   12 Customer Success & Churn   ·   13 Catálogo")
rect(s, MX, 4.50, CW, 0.86)
txt(s, MX + 0.20, 4.66, CW - 0.40, 0.60, tab_txt, 10, TEXTO, space=6)

txt(s, MX, 5.55, CW, 0.3, "DECISIONES DE DISEÑO QUE HAY QUE CONOCER", 9.5, CYAN, True)
vinetas(s, MX, 5.90, 12.1, [
    ("Las tres pestañas comerciales van al final a propósito: son las de menor cobertura, así el tablero abre por lo que está al día.", CYAN),
    ("Gastos Colombia y Gastos Consolidado son pestañas fijas: cada una fuerza su país y el filtro superior no las puede sacar de su alcance.", CYAN),
], size=10, gap=6)
pie(s, "El corte por líneas es exacto: los bloques de estilo y de script del archivo están en 10–199, 328–353 y 355–2.774.")

# ==================================================== 5 SCRIPTS PYTHON
s = nueva()
cabecera(s, "Los 27 scripts de Python",
         "Toda la preparación del dato ocurre aquí, antes de que el tablero abra · 4.690 líneas")
anchos = [2.55, 0.75, 8.90]
filas = [
    ["Familia", "N.º", ("Qué hace y cuáles son", CYAN, True, PP_ALIGN.LEFT)],
    ["Construir la base", "3", izq("construir_bd_maestra · separar_base (partió la base en dos el 13-ago) · preparar_google_sheets")],
    ["Cargue mensual", "7", izq("cargar_pyg_filiales_ene_jul · cargar_gastos_colombia_junio/julio · cargar_cartera_julio · cargar_mrr_julio · cargar_gasto_area_julio · cargar_proveedor_ventas_pct_junio")],
    ["Corrección de TRM", "3", izq("corregir_trm_ene_jun (quitó la TRM fija de 4.170) · corregir_trm_julio (reemplazó la implícita 3.505,27) · corregir_trm_2024_2025_y_peru_julio (tasas de relleno)")],
    ["Cierres y ajustes", "4", izq("corregir_junio_2026 · agregar_h1_2026 · agregar_presupuesto · actualizar_mrr_serie_completa")],
    ["Control de integridad", "3", izq("conciliar_capas (el motor de Revisar_Base) · sincronizar_subsidiarias (el de Corregir_Base) · borrar_ebitda_legacy")],
    ["Del modelo financiero", "2", izq("extraer_burn_runway (produce burn_runway.json) · extraer_modelo (P&G del modelo al tablero)")],
    ["Exportar", "2", izq("exportar_faltantes_contadora (la matriz de huecos) · exportar_pyg_para_modelo")],
    ["Publicar", "3", izq("generar_dashboard_para_compartir · generar_Dashboard_html (versión Apps Script) · generar_presentacion_consolidada")],
]
tabla(s, MX, 1.22, anchos, filas, alto_fila=0.60, alto_head=0.32, size=9.3)
pie(s, "Regla de operación del README: los .py son las instrucciones que ejecutan los tres botones .bat. No se abren ni se corren a mano. "
       "Cada script de cargue documenta en su encabezado la fuente, la tasa y el alcance de lo que escribe.")

# ============================================ 6 CAPA NUBE Y BOTONES
s = nueva()
cabecera(s, "La capa de nube y los tres botones",
         "Google Apps Script (preparado, sin desplegar) y el uso diario en el computador")

txt(s, MX, 1.20, 5.95, 0.3, "APPS SCRIPT · apps_script/Code.gs · 400 líneas", 9.5, CYAN, True)
caja(s, MX, 1.55, 5.95, 1.62, "doGet — lectura", [
    "Entrega las hojas del Sheet como matrices (AOA),",
    "el mismo formato que consume parseWorkbook().",
    "Acceso: público para quien tenga la URL."], VERDE2)
caja(s, MX, 3.34, 5.95, 1.62, "doPost — escritura", [
    "registro_individual — una fila en BD_Indicadores",
    "carga_masiva — plantilla de área, con modo",
    "     «previsualizar» (no escribe) o «confirmar»",
    "pipeline_masivo — «reemplazar» o «agregar»",
    "Todo cambio queda en la pestaña LOG."], VERDE2)
caja(s, MX, 5.13, 5.95, 1.62, "Seguridad", [
    "Token compartido en Script Properties + dominio",
    "autorizado: cocotech.ai y cfocus.co. Sin claves",
    "que repartir: la identidad la verifica Google.",
    "",
    "Session.getActiveUser() llega vacío en llamadas",
    "desde HTML externo: es un límite de CORS, documentado."], AMBAR)

x2 = MX + 6.25
txt(s, x2, 1.20, 5.95, 0.3, "USO DIARIO · tres archivos .bat", 9.5, CYAN, True)
caja(s, x2, 1.55, 5.95, 1.35, "Abrir_Dashboard_COCO.bat", [
    "Levanta un servidor local en el puerto 8740",
    "y abre el tablero en el navegador.",
    "No modifica nada."], VERDE2)
caja(s, x2, 3.05, 5.95, 1.35, "Revisar_Base.bat", [
    "Comprueba que las capas de la base digan lo mismo",
    "y que burn_runway.json corresponda al modelo vigente.",
    "No modifica nada. Correr antes de publicar un cierre."], VERDE2)
caja(s, x2, 4.55, 5.95, 1.35, "Corregir_Base.bat", [
    "Alinea las capas desfasadas.",
    "Muestra qué cambiaría y pide confirmación.",
    "Sí modifica, con respaldo previo."], AMBAR)
caja(s, x2, 6.05, 5.95, 0.72, "Por qué importa", [
    "Sin esta revisión, una capa vieja hacía que el tablero mostrara datos "
    "desactualizados sin avisar. Pasó con Perú julio y con Costa Rica junio."], ROJO)
pie(s, "El despliegue en Google está documentado y listo en DESPLIEGUE.md, pero todavía no se ha ejecutado: hoy el tablero corre local.")

# ================================================= 7 LOGICA DEL TABLERO
s = nueva()
cabecera(s, "Esquema de la lógica del tablero",
         "Qué ocurre desde que se abre el archivo hasta que se dibuja una pestaña")

pasos = [
    ("1 · ARRANQUE", "initUI()", ["Prepara la interfaz", "y pide burn_runway.json", "para la pestaña 03"]),
    ("2 · CARGA", "fetch / FileReader", ["Busca la base en disco.", "3 reintentos por OneDrive", "(700 ms × intento)"]),
    ("3 · LECTURA", "parseWorkbook()", ["XLSX.read() del libro.", "Abre exactamente", "10 hojas, ninguna más"]),
    ("4 · ESTADO", "objeto S en memoria", ["catalog · facts · trm ·", "trmPeru · pipeline ·", "gastosResumen · churn"]),
    ("5 · PINTADO", "render() → buildTabs()", ["Aplica los 5 filtros y", "llama al renderizador", "de la pestaña activa"]),
]
pw, pgap = 2.25, 0.24
for i, (num, fn, items) in enumerate(pasos):
    x = MX + i * (pw + pgap)
    rect(s, x, 1.28, pw, 1.70)
    txt(s, x + 0.15, 1.44, pw - 0.30, 0.22, num, 8.5, CYAN, True)
    txt(s, x + 0.15, 1.72, pw - 0.30, 0.26, fn, 10.5, VERDE2, True)
    txt(s, x + 0.15, 2.08, pw - 0.30, 0.82, "\n".join(items), 8.6, MUTED, space=3)
    if i < 4:
        flecha(s, x + pw + 0.02, 2.02, 0.20, 0.22)

txt(s, MX, 3.20, CW, 0.3, "LOS CINCO FILTROS", 9.5, CYAN, True)
rect(s, MX, 3.54, CW, 0.52)
txt(s, MX + 0.20, 3.66, CW - 0.40, 0.32,
    "Rango de fechas   ·   Escenario (Real / Presupuesto)   ·   País   ·   Moneda (COP / USD)   ·   "
    "Base de TRM (implícita de EEFF o promedio de mercado)", 10, TEXTO, anchor=MSO_ANCHOR.MIDDLE)

caja(s, MX, 4.25, 3.93, 2.05, "Si no hay servidor o falla la lectura", [
    "El tablero cae a la copia guardada en el navegador",
    "(localStorage, clave coco_bd_julio_2026_r4) y avisa",
    "en ámbar: «de memoria, puede estar desactualizado».",
    "",
    "Nunca muestra caché en silencio."], AMBAR)
caja(s, MX + 4.13, 4.25, 3.94, 2.05, "Cómo se leen los vacíos", [
    "$0     cero reportado — es un dato, no falta nada",
    "s/d    dato no recibido — perseguirlo antes del cierre",
    "n/a    no aplica a esa entidad",
    "",
    "Las líneas de tendencia se cortan en los meses sin dato:",
    "unir el hueco dibujaría una pendiente que nadie reportó."], VERDE2)
caja(s, MX + 8.27, 4.25, 3.93, 2.05, "Burn & Runway va aparte", [
    "La pestaña 03 no sale de la base: sale de",
    "burn_runway.json, sellado desde el modelo V3.",
    "",
    "Hoy va cortada a junio mientras las otras once",
    "van a julio, y convierte la caja de EE.UU. a 3.248,87",
    "y no a 3.505. No comparar esas cifras de frente."], ROJO)
pie(s, "Detalle verificable en Dashboard_COCO.html: initUI() y el bucle de reintentos al final del bloque de script; parseWorkbook() y findSheet() para las 10 hojas.")

# ============================================ 8 DE DONDE SALE LA INFORMACION
s = nueva()
cabecera(s, "Esquema de dónde extrae la información",
         "Cuatro niveles, de la fuente contable original a la cifra que se ve en pantalla")

y0, bh, bgap = 1.22, 1.13, 0.26
niveles = [
    ("NIVEL 1 · FUENTES PRIMARIAS", VERDE2,
     "Estados financieros oficiales de Colombia (PDF)   ·   Auxiliares contables mensuales por país — 22 archivos, 986 registros\n"
     "Modelo Presupuesto Completo COCO V3 (presupuesto y caja)   ·   Facturación y CRM   ·   Extractos bancarios"),
    ("NIVEL 2 · TRAZABILIDAD — BD_TRAZABILIDAD_COCO.xlsx · 24 hojas", AMBAR,
     "EEFF_Fuentes · EEFF_Homologacion (90 reglas de mapeo) · EEFF_Base_Homologada · EEFF_Balance · EEFF_Movimientos\n"
     "EEFF_Intercompanias · EEFF_Alertas · EEFF_Checks          ⚠  Archivo de auditoría: el tablero NUNCA lo abre"),
    ("NIVEL 3 · BASE OPERATIVA — BD_MAESTRA_COCO.xlsx · 14 hojas", CYAN,
     "BD_Indicadores — la tabla de hechos: 1.852 filas en formato largo (periodo · país · moneda · escenario · código · valor · fuente)\n"
     "Diccionario · TRM · TRM_Peru · Calc_LTV_Fin · Analisis_Churn · Analisis_Reconciliacion · Pipeline_Comercial · CATALOGOS · LOG · cadena de gastos"),
    ("NIVEL 4 · CONSUMO — Dashboard_COCO.html", VERDE2,
     "Abre exactamente 10 hojas de la base maestra y ninguna otra   +   burn_runway.json para la pestaña 03\n"
     "Todo lo que se ve en las 13 pestañas sale de ahí. Si un número no está en esas 10 hojas, no existe para el tablero."),
]
for i, (titulo, col, cuerpo) in enumerate(niveles):
    y = y0 + i * (bh + bgap)
    rect(s, MX, y, CW, bh)
    txt(s, MX + 0.20, y + 0.14, CW - 0.40, 0.26, titulo, 10.5, col, True)
    txt(s, MX + 0.20, y + 0.46, CW - 0.40, 0.60, cuerpo, 9, MUTED, space=4)
    if i < 3:
        flecha(s, MX + CW / 2 - 0.11, y + bh + 0.02, 0.22, 0.22, abajo=True)

pie(s, "La separación de niveles 2 y 3 se hizo el 13-ago-2026: la misma cifra vivía en tres o cuatro hojas y, al actualizar una y olvidar las otras, "
       "el tablero mostraba datos viejos sin avisar. Hoy solo hay una capa que el tablero consume.")

# ================================================ 9 COMO CARGAR INFORMACION
s = nueva()
cabecera(s, "De qué manera se puede cargar la información",
         "Cinco vías, tres disponibles hoy y dos listas para activar")
anchos = [0.45, 2.55, 1.45, 3.60, 4.15]


def cen(t, col=TEXTO, bold=False):
    return (t, col, bold, PP_ALIGN.CENTER)


filas = [
    [cen(""), ("Vía", CYAN, True, PP_ALIGN.LEFT), cen("Estado", CYAN, True),
     ("Cómo funciona", CYAN, True, PP_ALIGN.LEFT), ("Cuándo usarla", CYAN, True, PP_ALIGN.LEFT)],
    [cen("A", CYAN, True), izq("Automática desde disco"), cen("Activa", VERDE2, True),
     izq("El .bat levanta un servidor local y el tablero busca migracion/BD_MAESTRA_COCO.xlsx"),
     izq("Uso diario. Es la vía normal: doble clic y listo")],
    [cen("B", CYAN, True), izq("Manual «Cargar BD»"), cen("Activa", VERDE2, True),
     izq("Arrastrar el .xlsx al tablero; lo lee el navegador con FileReader"),
     izq("Para revisar una versión distinta sin tocar la oficial")],
    [cen("C", CYAN, True), izq("Caché del navegador"), cen("Respaldo", AMBAR, True),
     izq("Copia en localStorage de la última carga exitosa"),
     izq("Solo si falla la lectura. Avisa en ámbar que puede estar vieja")],
    [cen("D", CYAN, True), izq("Cargue por área en la nube"), cen("Preparada", CYAN, True),
     izq("Apps Script: cada área entra con su cuenta y sube su porción, con previsualización antes de confirmar"),
     izq("Al desplegarse: quita el cuello de botella de una sola persona cargando todo")],
    [cen("E", CYAN, True), izq("Desde el modelo financiero"), cen("Activa", VERDE2, True),
     izq("Actualizar_Burn_Runway.bat regenera burn_runway.json desde el modelo V3"),
     izq("Cada vez que cambie el modelo. Alimenta solo la pestaña 03")],
]
tabla(s, MX, 1.22, anchos, filas, alto_fila=0.60, alto_head=0.32, size=9.2)

txt(s, MX, 4.70, CW, 0.3, "TRES CAUTELAS AL CARGAR", 9.5, CYAN, True)
vinetas(s, MX, 5.05, 12.1, [
    ("Acumulado ≠ movimiento del mes. Para un mes se usa el movimiento del auxiliar (débitos y créditos), no el saldo.", AMBAR),
    ("Al cambiar cifras de un país hay que recalcular el consolidado, y dejar fuente, tasa y alcance en la columna «comentario».", AMBAR),
    ("Si se reemplaza el modelo y no se regenera burn_runway.json, Revisar_Base.bat avisa DESACTUALIZADO. El campo «corte» va escrito a mano.", ROJO),
], size=10.2, gap=8)
pie(s, "Vía D documentada paso a paso en DESPLIEGUE.md. Lectura pública con la URL; escritura con token y dominio autorizado, y todo registrado en LOG.")

# ================================= 10 CIFRAS CON CONTABILIDAD - ALTA
s = nueva()
cabecera(s, "Cifras que hay que revisar con contabilidad — 1 de 2",
         "Alertas de severidad alta: bloquean un consolidado firmable")

txt(s, MX, 1.18, CW, 0.3, "① CONVERSIÓN DE MONEDA — SIN POLÍTICA DEFINIDA", 10, ROJO, True)
anchos = [4.55, 2.60, 5.05]
filas = [
    ["Tasa en uso", "COP/USD", "Dónde y con qué efecto"],
    ["Modelo V3 · TrmLargoPlazo", "3.690,00", "Tasa de presupuesto usada sobre cifras reales"],
    ["Hoja Consolidado · tasa única del acumulado", "3.505,27", "La propia nota del modelo la llama «una simplificación»"],
    ["Burn & Runway · tasa de cierre de un saldo", "3.248,87", "Solo la caja de EE.UU.: no comparable con el resto"],
]
tabla(s, MX, 1.52, anchos, filas, alto_fila=0.30, alto_head=0.30, size=9.4)
txt(s, MX, 2.80, CW, 0.42,
    "Efecto: los ingresos de Colombia de junio valen USD 203.646 a 3.690 y USD 214.378 a 3.505 — 5,3% de diferencia sobre la misma cifra contable. "
    "Lo correcto es tasa promedio del mes para el resultado y tasa de cierre para el balance, con la diferencia a ajuste por conversión en patrimonio.",
    9, MUTED)

txt(s, MX, 3.32, CW, 0.3, "② ELIMINACIONES INTERCOMPAÑÍA — IDENTIFICADAS, NO PRACTICADAS", 10, ROJO, True)
anchos = [3.30, 3.05, 1.95, 3.90]
filas = [
    ["Partida (corte julio 2026)", "Contraparte", "USD", "Estado"],
    ["Otras cuentas por cobrar — EE.UU.", "Coco Inversiones Tecnológicas SAS", "571.891,71", "Pendiente de conciliar"],
    ["Otras cuentas por cobrar — EE.UU.", "Coco Tech AI Costa Rica", "37.219,96", "Pendiente de conciliar"],
    ["Otras cuentas por cobrar — EE.UU.", "Coco Tech AI Perú", "4.300,00", "Pendiente de conciliar"],
    ["Ingresos por servicios — EE.UU.", "Coco Tech AI Perú", "4.010,00", "Pendiente de conciliar"],
    [("TOTAL POR ELIMINAR", TEXTO, True, PP_ALIGN.LEFT), ("", TEXTO, True, PP_ALIGN.LEFT),
     ("617.421,67", VERDE2, True, PP_ALIGN.RIGHT), ("", TEXTO, True, PP_ALIGN.LEFT)],
]
tabla(s, MX, 3.66, anchos, filas, alto_fila=0.30, alto_head=0.30, size=9.4)
txt(s, MX, 5.62, CW, 0.75,
    "Dimensión del problema: EE.UU. reporta USD 616.411,67 en «Otras cuentas por cobrar» y USD 613.411,67 de eso es intercompañía — el 99,5%. "
    "Sin la eliminación, el activo del grupo está inflado en más de medio millón de dólares.\n"
    "Pendiente adicional: conciliar contra los COP 2.641.566.481 de «Cuentas por pagar a socios» de Colombia. El modelo los llama «CxP socios» en una hoja "
    "y «el pasivo con Coco Delaware» en otra. Si son con la matriz se eliminan; si son con socios personas naturales, se quedan.",
    9, MUTED, space=5)
pie(s, "Fuente: BD_TRAZABILIDAD_COCO.xlsx, hojas EEFF_Alertas (severidad Alta, ambas abiertas), EEFF_Intercompanias y EEFF_Balance.")

# ================================= 11 CIFRAS CON CONTABILIDAD - MEDIA
s = nueva()
cabecera(s, "Cifras que hay que revisar con contabilidad — 2 de 2",
         "Cierre contable, precios de transferencia, clasificación y huecos de cargue")

txt(s, MX, 1.18, 6.00, 0.3, "③ EL RESULTADO NO CUADRA CON LAS CUENTAS 36", 10, ROJO, True)
anchos = [1.55, 1.55, 1.55, 1.35]
filas = [
    ["País", "Calculado", "Registrado", "Diferencia"],
    ["EE.UU. (USD)", "-48.547,93", "13.042,50", "61.590,43"],
    ["Costa Rica (USD)", "109.028,68", "-22.674,27", "131.702,95"],
    ["Perú (PEN)", "25.372,37", "0,00", "25.372,37"],
]
tabla(s, MX, 1.52, anchos, filas, alto_fila=0.30, alto_head=0.30, size=9.2)
txt(s, MX, 2.80, 6.00, 0.42,
    "Las cuentas 36 no representan el resultado abierto en las cuentas 4 y 5. "
    "Hay que revisar el cierre y el traslado de resultados en los tres países.", 8.8, MUTED)

txt(s, MX, 3.35, 6.00, 0.3, "④ PRECIOS DE TRANSFERENCIA — COSTA RICA", 10, ROJO, True)
rect(s, MX, 3.70, 6.00, 1.42)
txt(s, MX + 0.18, 3.86, 5.64, 1.12,
    "Costa Rica factura con margen aparente del 98%: los costos de entregar ese servicio están en otra entidad.\n"
    "Julio: USD 134.104 facturados, registrados contra Clientes y no contra el banco. Sin Costa Rica, julio da pérdida de USD 886.\n"
    "No es un error de captura. Es una decisión de precios de transferencia pendiente.",
    9, MUTED, space=5)

txt(s, MX, 5.30, 6.00, 0.3, "⑤ IMPUESTO DE RENTA NO REGISTRADO — COLOMBIA", 10, AMBAR, True)
rect(s, MX, 5.65, 6.00, 1.10)
txt(s, MX + 0.18, 5.81, 5.64, 0.82,
    "Colombia registra impuesto de renta en cero los siete meses. La provisión estimada al 35% suma COP 99,63 MM.\n"
    "Utilidad neta reportada COP 246,73 MM  →  con renta provisionada, COP 147,10 MM.",
    9, MUTED, space=5)

x2 = MX + 6.25
txt(s, x2, 1.18, 5.95, 0.3, "⑥ SIGNOS Y CLASIFICACIÓN — REVISAR SOPORTE", 10, AMBAR, True)
anchos2 = [3.35, 1.35, 1.25]
filas2 = [
    ["Partida", "Monto", "Moneda"],
    ["24080201 IGV compras — signo contrario", "-984,16", "PEN"],
    ["51055101 Dotación — signo contrario", "-540,92", "USD"],
    ["13050501 Nacionales — signo contrario", "-217,00", "USD"],
    ["Capote & Capote PA en auxilios", "7.600,00", "USD"],
    ["AT&T en «otros servicios»", "773,99", "USD"],
]
tabla(s, x2, 1.52, anchos2, filas2, alto_fila=0.30, alto_head=0.30, size=9.2)

txt(s, x2, 3.42, 5.95, 0.3, "⑦ SALDOS INICIALES REEXPRESADOS SIN DOCUMENTAR", 10, AMBAR, True)
rect(s, x2, 3.77, 5.95, 1.10)
txt(s, x2 + 0.18, 3.93, 5.59, 0.82,
    "Perú: 16 cuentas feb→mar, 15 mar→abr y 21 jun→jul (máx. PEN 7.305,25).   Costa Rica: 7 cuentas ene→feb y 12 feb→mar (máx. USD 2.300).\n"
    "El total sigue cuadrando, así que son reclasificaciones. Hay que documentarlas antes de comparar meses.",
    9, MUTED, space=5)

txt(s, x2, 5.02, 5.95, 0.3, "⑧ HUECOS DE CARGUE — 78 INDICADORES", 10, AMBAR, True)
rect(s, x2, 5.37, 5.95, 1.38)
txt(s, x2 + 0.18, 5.53, 5.59, 1.10,
    "Nunca cargados: gastos operativos totales · margen bruto y margen bruto % · margen EBITDA · ARPA · DPO · CCC · liquidez total · churn MRR mensual · cobros del periodo.\n"
    "Rezagos: MRR y Customer Success a mayo · ARR, marketing y cartera a junio · detalle contable colombiano a junio.\n"
    "El balance de Costa Rica va a junio mientras los otros tres van a julio. La TRM de Perú de julio es provisional (promedio hasta el día 9).",
    8.6, MUTED, space=4)
pie(s, "Fuente: EEFF_Alertas, migracion/matriz_faltantes.json (78 registros), Indicadores_faltantes_para_contadora.xlsx y la hoja Pack Junta Jul-26 del modelo V3.")

# ==================================================== 12 QUE HACER AHORA
s = nueva()
cabecera(s, "Qué hacer ahora",
         "Lo que está listo, lo que falta decidir y en qué orden")

txt(s, MX, 1.20, 3.93, 0.3, "EL TABLERO ESTÁ SANO", 10, VERDE2, True)
rect(s, MX, 1.55, 3.93, 2.55)
txt(s, MX + 0.18, 1.72, 3.57, 2.25,
    "· Una sola fuente de verdad, con control de integridad automático\n\n"
    "· Los vacíos se distinguen de los ceros, por código\n\n"
    "· Los cuatro países llegan a julio en el P&G\n\n"
    "· Cada cifra es rastreable hasta su archivo de origen\n\n"
    "· Cero dependencias externas: abre sin internet",
    9.4, MUTED, space=3)

txt(s, MX + 4.13, 1.20, 3.94, 0.3, "LO QUE FALTA DECIDIR", 10, AMBAR, True)
rect(s, MX + 4.13, 1.55, 3.94, 2.55)
txt(s, MX + 4.31, 1.72, 3.58, 2.25,
    "· Política de conversión: qué tasa, para qué estado\n\n"
    "· Si los COP 2.641 MM de «CxP socios» son con la matriz\n\n"
    "· Precios de transferencia de Costa Rica\n\n"
    "· Si se provisiona la renta de Colombia\n\n"
    "· Si se despliega el cargue por área en Google",
    9.4, MUTED, space=3)

txt(s, MX + 8.27, 1.20, 3.93, 0.3, "EN QUÉ ORDEN", 10, CYAN, True)
rect(s, MX + 8.27, 1.55, 3.93, 2.55)
txt(s, MX + 8.45, 1.72, 3.57, 2.25,
    "1.  Definir tasas y política de conversión\n\n"
    "2.  Conciliar y registrar las eliminaciones\n\n"
    "3.  Cerrar el traslado de resultados (cuentas 36)\n\n"
    "4.  Completar julio en Costa Rica y los huecos\n\n"
    "5.  Recién ahí, balance consolidado firmable",
    9.4, MUTED, space=3)

txt(s, MX, 4.35, CW, 0.3, "LO QUE HOY SE PUEDE Y NO SE PUEDE AFIRMAR", 9.5, CYAN, True)
rect(s, MX, 4.70, CW, 1.85)
txt(s, MX + 0.22, 4.90, CW - 0.44, 1.50,
    "SÍ:  el P&G consolidado enero–julio de los cuatro países, en USD, con las tasas declaradas y sus notas metodológicas.\n\n"
    "NO:  un balance general consolidado. Faltan las eliminaciones intercompañía —USD 617.422 identificados y sin practicar— y la política de conversión. "
    "Las dos alertas de severidad alta de EEFF_Alertas siguen abiertas.\n\n"
    "OJO:  el bloque mensual de la hoja Consolidado no concilia con su propio acumulado: USD 217.562 de EBITDA sumando los períodos contra USD 127.717 del "
    "acumulado. Son USD 89.845 de diferencia, ya detectados en la auditoría del modelo y todavía sin corregir.",
    10, TEXTO, space=6)
pie(s, "Documento preparado el 25 de agosto de 2026 sobre el proyecto COCO_Dashboard_Cloud y el Modelo Presupuesto Completo COCO V3 (13-ago-2026).")

prs.save(SALIDA)
print("OK ->", SALIDA)
print("Diapositivas:", len(prs.slides._sldIdLst))
