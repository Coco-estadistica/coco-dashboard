// Informe de cosechas para junta directiva — Districolmotos
// Genera Cosechas_Districolmotos_Junta.pptx a partir de deck_data.json
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const D = JSON.parse(fs.readFileSync(__dirname + '/deck_data.json', 'utf8'));

// ---------------------------------------------------------------- paleta y tipos
const NAVY = '12263A', INK = '18242F', MUTED = '62717F', SOFT = '8B98A5';
const WHITE = 'FFFFFF', SURF = 'F2F5F8', LINE = 'DDE3EA';
const BLUE = '2A78D6', ORANGE = 'EB6834', RED = 'C4342F', GREEN = '0F7B3F';
const ORD = ['9EC5F4', '5598E7', '2A78D6', '1C5CAB', '0D366B'];
const HEAD = 'Cambria', BODY = 'Calibri';
const MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const cosNice = c => MES[+c.slice(4) - 1] + '-' + c.slice(0, 4);
const nf = (v, d = 1) => v == null ? 's/d'
  : v.toLocaleString('es-CO', { minimumFractionDigits: d, maximumFractionDigits: d });

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';            // 13.333 x 7.5
pres.author = 'Districolmotos';
pres.title = 'Cosechas de originación — Informe a junta';
const W = 13.333, H = 7.5, MX = 0.62;

// ---------------------------------------------------------------- helpers
function txt(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, margin: 0 }, o)); }

function header(s, kicker, title, sub) {
  txt(s, kicker.toUpperCase(), { x: MX, y: 0.34, w: 11, h: 0.22, fontSize: 10.5,
    color: BLUE, bold: true, charSpacing: 1.6, fontFace: BODY });
  txt(s, title, { x: MX, y: 0.58, w: W - 2 * MX, h: 0.80, fontSize: 23,
    bold: true, color: INK, fontFace: HEAD, valign: 'top', lineSpacing: 28 });
  if (sub) txt(s, sub, { x: MX, y: 1.42, w: W - 2 * MX, h: 0.50, fontSize: 11.5,
    color: MUTED, fontFace: BODY, valign: 'top', lineSpacing: 15 });
}

function footer(s, t) {
  txt(s, t, { x: MX, y: H - 0.78, w: W - 2 * MX, h: 0.56, fontSize: 9,
    color: SOFT, fontFace: BODY, valign: 'bottom', lineSpacing: 12 });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || SURF }, line: { color: fill ? fill : LINE, width: 0.75 } });
}

function stat(s, x, y, w, value, label, note, color) {
  card(s, x, y, w, 1.62);
  txt(s, value, { x: x + 0.24, y: y + 0.18, w: w - 0.4, h: 0.62, fontSize: 32,
    bold: true, color: color || INK, fontFace: HEAD });
  txt(s, label, { x: x + 0.24, y: y + 0.84, w: w - 0.4, h: 0.3, fontSize: 11.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, note, { x: x + 0.24, y: y + 1.13, w: w - 0.4, h: 0.4, fontSize: 9.8,
    color: MUTED, fontFace: BODY, valign: 'top' });
}

function numbered(s, x, y, w, n, title, bodyTxt) {
  s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.02, w: 0.34, h: 0.34,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  txt(s, String(n), { x, y: y + 0.02, w: 0.34, h: 0.34, fontSize: 13, bold: true,
    color: WHITE, align: 'center', valign: 'middle', fontFace: BODY });
  txt(s, title, { x: x + 0.5, y: y - 0.02, w: w - 0.5, h: 0.52, fontSize: 13, bold: true,
    color: INK, fontFace: BODY, valign: 'top', lineSpacing: 17 });
  txt(s, bodyTxt, { x: x + 0.5, y: y + 0.54, w: w - 0.5, h: 1.1, fontSize: 10.8,
    color: MUTED, fontFace: BODY, valign: 'top', lineSpacing: 14 });
}

// eje de MOB común a todos los gráficos de curvas
const LBL = Array.from({ length: 37 }, (_, i) => String(i));
const cada = k => LBL.map((v, i) => (i % k === 0 ? v : ' '));   // eje legible: solo cada k meses
// 56 cosechas rotadas son ilegibles proyectadas: se rotula solo el enero de cada año
const ejeCosechas = arr => arr.map((c, i) => (c.slice(4) === '01' || i === 0 ? c.slice(0, 4) : ' '));
function serieMob(mobs, pcts) {
  const m = new Map(mobs.map((v, i) => [v, pcts[i]]));
  return LBL.map((_, i) => (m.has(i) ? m.get(i) : null));
}
const baseChart = {
  showLegend: false, showTitle: false,
  valAxisMinVal: 0, valAxisLabelFormatCode: '0"%"',
  valGridLine: { color: 'EDF0F3', size: 0.75 }, catGridLine: { style: 'none' },
  catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
  catAxisLineShow: false, valAxisLineShow: false,
  lineSmooth: false, lineDataSymbol: 'none', dataBorder: { pt: 1, color: 'FFFFFF' },
  chartColors: [BLUE], lineSize: 2,
};

// ================================================================= 1. PORTADA
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, 'INFORME A JUNTA DIRECTIVA · SEPTIEMBRE DE 2026', { x: MX, y: 1.5, w: 11, h: 0.3,
    fontSize: 11, color: '7FA9DC', bold: true, charSpacing: 2, fontFace: BODY });
  txt(s, 'Desde 2024 la originación mejoró:\nal mes 18, la mitad de la mora de 2022 y 2023',
    { x: MX, y: 2.0, w: 11.2, h: 1.7, fontSize: 34, bold: true, color: WHITE,
      fontFace: HEAD, lineSpacing: 45, valign: 'top' });
  txt(s, 'Análisis de cosechas de la originación de crédito de motocicletas',
    { x: MX, y: 3.85, w: 10, h: 0.36, fontSize: 15, color: 'CADCFC', fontFace: BODY });

  const items = [
    ['56 cosechas', 'enero 2022 – agosto 2026'],
    ['COP 16.448 M', 'colocación analizada'],
    ['31-ago-2026', 'fecha de corte'],
  ];
  items.forEach(([a, b], i) => {
    const x = MX + i * 3.5;
    txt(s, a, { x, y: 4.75, w: 3.2, h: 0.4, fontSize: 19, bold: true, color: WHITE, fontFace: HEAD });
    txt(s, b, { x, y: 5.16, w: 3.2, h: 0.3, fontSize: 10.5, color: '9FB6D4', fontFace: BODY });
  });
  txt(s, 'Fuente: Cosechas_2026.xlsx — hojas Consolidado y Cosechas, y los siete cortes mensuales de cartera de feb a ago de 2026.',
    { x: MX, y: 6.6, w: 11.5, h: 0.4, fontSize: 9.5, color: '7189A5', fontFace: BODY });
  s.addNotes('Informe de cosechas de la cartera de crédito de motos. El mensaje central: la originación mejoró desde 2024. Pero hay una salvedad contable de julio que hay que explicar antes de mostrar resultados — lámina 3.');
}

// ================================================================= 2. RESUMEN EJECUTIVO
{
  const s = pres.addSlide();
  header(s, 'Resumen ejecutivo', 'Tres mensajes y tres decisiones',
    'Si solo se lee una lámina de este informe, que sea esta.');

  stat(s, MX, 2.0, 3.85, '4,4%', 'Mora de las cosechas 2023 al mes 12',
    'El peor año. 2022 iba en 2,5% en el mismo punto de la vida del crédito.', RED);
  stat(s, MX + 4.05, 2.0, 3.85, '2,9%', 'Mora de las cosechas 2024 al mes 18',
    'Contra 5,6% de 2023 y 4,7% de 2022. La corrección funcionó.', GREEN);
  stat(s, MX + 8.10, 2.0, 3.85, 'COP 549 M', 'Salieron del reporte en julio de 2026',
    'Castigo o depuración, no recuperación. Invalida el dato del último mes.', ORANGE);

  const dec = [
    ['Intervenir los dos puntos que pierden', 'Cartagena-Suzuki y Puerto Berrío AKT llegan al mes 6 con 12,4% y 12,0% de mora, contra 5,8% del mejor punto. Cerrar esa brecha vale unos COP 71 millones al año.'],
    ['Cuadrar el castigo con contabilidad', 'Confirmar qué se castigó, contra qué provisión y con qué efecto en el P&G, antes de reportar cualquier mejora de la mora.'],
    ['Pedir tres campos que hoy no existen', 'Cuota inicial, score de buró y asesor comercial. Sin ellos no se puede fijar el punto de corte de aprobación.'],
  ];
  txt(s, 'Lo que se pide aprobar', { x: MX, y: 4.1, w: 6, h: 0.3, fontSize: 13,
    bold: true, color: INK, fontFace: BODY });
  dec.forEach(([t, b], i) => numbered(s, MX + i * 4.05, 4.6, 3.85, i + 1, t, b));
  footer(s, 'Indicador: capital vencido de los créditos con más de 30 días de mora, sobre la colocación original de cada cosecha. Series cortadas en jun-2026.');
  s.addNotes('Arrancar por aquí. Los tres números de arriba son el informe entero. Las tres decisiones de abajo son lo que se somete a aprobación.');
}

// ================================================================= 3. LA SALVEDAD (dark)
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, 'LEER ANTES DE INTERPRETAR CUALQUIER CIFRA', { x: MX, y: 0.62, w: 11, h: 0.26,
    fontSize: 11, color: 'F0A17A', bold: true, charSpacing: 1.8, fontFace: BODY });
  txt(s, 'La mora cayó 64% en un mes, y no fue por recuperación de cartera',
    { x: MX, y: 0.95, w: 11.6, h: 0.9, fontSize: 26, bold: true, color: WHITE,
      fontFace: HEAD, lineSpacing: 32 });

  const facts = [
    ['COP 549 M', 'de mora salieron del reporte entre junio y julio de 2026'],
    ['242', 'créditos desaparecieron del archivo de un mes a otro'],
    ['22', 'cosechas bajaron de golpe; doce quedaron en 0,00% exacto'],
    ['jun-2016 → oct-2021', 'saltó la fecha de desembolso más antigua del reporte'],
  ];
  facts.forEach(([a, b], i) => {
    const x = MX + (i % 2) * 6.1, y = 2.15 + Math.floor(i / 2) * 1.35;
    txt(s, a, { x, y, w: 5.7, h: 0.46, fontSize: 22, bold: true, color: 'F3B08C', fontFace: HEAD });
    txt(s, b, { x, y: y + 0.48, w: 5.7, h: 0.6, fontSize: 12.5, color: 'CADCFC',
      fontFace: BODY, valign: 'top', lineSpacing: 17 });
  });

  card(s, MX, 5.25, W - 2 * MX, 1.25, '1B3450');
  txt(s, 'Qué hicimos con eso', { x: MX + 0.3, y: 5.42, w: 11, h: 0.28, fontSize: 12.5,
    bold: true, color: WHITE, fontFace: BODY });
  txt(s, 'Todas las series de este informe se cortan en junio de 2026. El dato de julio y agosto existe en el archivo, pero no es comparable hacia atrás: quien lo lea sin esta advertencia va a concluir que la cartera mejoró un 64% en treinta días.',
    { x: MX + 0.3, y: 5.72, w: 11.5, h: 0.72, fontSize: 11.5, color: 'A8C0DC',
      fontFace: BODY, valign: 'top', lineSpacing: 16 });
  s.addNotes('Esta lámina va antes de los resultados a propósito. Si la junta se entera del castigo después de ver las mejoras, se pierde la credibilidad de todo el informe.');
}

// ================================================================= 4. QUÉ ES UNA COSECHA
{
  const s = pres.addSlide();
  header(s, 'Cómo se lee', 'Una cosecha es lo prestado en un mes, seguido a lo largo de su vida',
    'Comparar cosechas solo tiene sentido si se miden a la misma edad, no en la misma fecha.');

  const pasos = [
    ['Se agrupa por mes de desembolso', 'Todo lo desembolsado en, por ejemplo, mayo de 2023 es una sola cosecha: COP 312 millones.'],
    ['Se sigue mes a mes de su vida', 'El mes 0 es el desembolso, el mes 12 es un año después. A eso lo llamamos MOB (meses en libros).'],
    ['Se mide siempre sobre lo colocado', 'La mora se divide por la colocación original de esa cosecha, que nunca cambia. Por eso son comparables.'],
  ];
  pasos.forEach(([t, b], i) => {
    const x = MX + i * 4.05;
    card(s, x, 2.05, 3.85, 1.78);
    numbered(s, x + 0.28, 2.24, 3.25, i + 1, t, b);
  });

  card(s, MX, 4.15, W - 2 * MX, 2.3, 'FBF2EC');
  txt(s, 'El error que hay que evitar', { x: MX + 0.35, y: 4.4, w: 11, h: 0.3,
    fontSize: 14, bold: true, color: '9A4520', fontFace: BODY });
  txt(s, 'Comparar la cosecha de hace tres meses con la de hace dos años y concluir que la nueva es mejor. La nueva SIEMPRE se ve mejor: todavía no ha tenido tiempo de deteriorarse. Toda comparación de este informe se hace a la misma edad — la mora al mes 6 contra la mora al mes 6, la del mes 12 contra la del mes 12.',
    { x: MX + 0.35, y: 4.75, w: 11.35, h: 1.05, fontSize: 12.5, color: '6E4130',
      fontFace: BODY, valign: 'top', lineSpacing: 18 });
  txt(s, 'Por la misma razón, las cosechas jóvenes aparecen con la curva más corta en las láminas siguientes: no es que falte el dato, es que todavía no ha pasado el tiempo.',
    { x: MX + 0.35, y: 5.85, w: 11.35, h: 0.5, fontSize: 11, color: '8A6552',
      fontFace: BODY, italic: true, valign: 'top' });
  s.addNotes('Treinta segundos. Si la junta ya maneja el concepto, saltar directo a la lámina 5.');
}

// ================================================================= 5. CURVAS POR AÑO
{
  const s = pres.addSlide();
  header(s, 'El mensaje central', '2024 y 2025 maduran a la mitad del ritmo de 2022 y 2023',
    'Capital vencido con mora 30+, como % de la colocación original de cada cosecha. Agrupadas por año y ponderadas por monto colocado. Observaciones hasta jun-2026.');

  const years = Object.keys(D.curvas_anio).sort();
  const data = years.map(y => ({ name: 'Cosechas ' + y, labels: cada(6),
    values: serieMob(D.curvas_anio[y].mob, D.curvas_anio[y].pct) }));
  s.addChart(pres.ChartType.line, data, Object.assign({}, baseChart, {
    x: MX, y: 2.02, w: 8.5, h: 4.55, chartColors: ORD, lineSize: 2.5,
    showLegend: true, legendPos: 'b', legendFontSize: 10.5, legendColor: MUTED,
    valAxisMaxVal: 9, valAxisMajorUnit: 3, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
    showCatAxisTitle: true, catAxisTitle: 'Meses en libros desde el desembolso (MOB)',
    catAxisTitleFontSize: 10, catAxisTitleColor: SOFT,
  }));

  const notas = [
    ['Al mes 12', '2023 llega a 4,4% y 2024 a 2,3%. Casi la mitad, medido en el mismo punto de la vida.'],
    ['Al mes 18', '2023 va en 5,6%, 2022 en 4,7% y 2024 en 2,9%.'],
    ['Ojo con el final', '2022 y 2023 convergen cerca del 8% hacia los meses 24–30. 2024 apenas va en el mes 18: la ventaja es de velocidad, todavía no está probada al final de la vida.'],
    ['Por qué las líneas ceden', 'La curva mide saldo vencido vivo, no acumulado. Baja cuando el crédito se pone al día, se paga o se castiga.'],
  ];
  let y = 2.05;
  notas.forEach(([t, b]) => {
    txt(s, t, { x: 9.5, y, w: 3.3, h: 0.26, fontSize: 12, bold: true, color: INK, fontFace: BODY });
    txt(s, b, { x: 9.5, y: y + 0.27, w: 3.3, h: 0.98, fontSize: 10.2, color: MUTED,
      fontFace: BODY, valign: 'top', lineSpacing: 14 });
    y += 1.30;
  });
  footer(s, 'Composición constante: cada línea llega solo hasta el mes en que todas las cosechas de ese año siguen observadas. 2022 arranca en el mes 12 porque el archivo empieza a observar en ene-2023.');
  s.addNotes('El gráfico del informe. Si hay una sola lámina que discutir, es esta.');
}

// ================================================================= 6. MOB 12 POR COSECHA
{
  const s = pres.addSlide();
  header(s, 'Dónde está el daño', 'Mayo, abril y septiembre de 2023 concentran el deterioro',
    'Mora 30+ al mes 12 de vida, cosecha por cosecha. % sobre la colocación de cada una.');

  const dat = D.mob_fijo.filter(d => d.m12 != null);
  const labels = ejeCosechas(dat.map(d => d.cosecha));
  const is23 = c => c.slice(0, 4) === '2023';
  s.addChart(pres.ChartType.bar, [
    { name: 'Cosechas 2023', labels, values: dat.map(d => is23(d.cosecha) ? d.m12 : null) },
    { name: 'Resto de cosechas', labels, values: dat.map(d => is23(d.cosecha) ? null : d.m12) },
  ], Object.assign({}, baseChart, {
    x: MX, y: 2.02, w: W - 2 * MX, h: 4.10, barDir: 'col', barGrouping: 'stacked',
    chartColors: [RED, 'C8D2DC'], valAxisMaxVal: 9, valAxisMajorUnit: 3,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    showLegend: true, legendPos: 'b', legendFontSize: 10.5, legendColor: MUTED,
  }));
  txt(s, 'Las tres peores cosechas — may-2023 (8,5%), abr-2023 (7,4%) y sep-2023 (6,2%) — colocaron COP 889 millones entre las tres. Las cosechas desde jul-2025 todavía no cumplen 12 meses: su ausencia en el gráfico no es un cero.',
    { x: MX, y: 6.35, w: W - 2 * MX, h: 0.6, fontSize: 11, color: MUTED,
      fontFace: BODY, valign: 'top', lineSpacing: 15 });
  s.addNotes('Comparación a edad constante: todas las cosechas medidas en el mes 12. Es la única comparación limpia entre cosechas.');
}

// ============================================ 7-11. SMALL MULTIPLES, UNA LÁMINA POR AÑO
const REF = D.curvas_anio['2024'];
const refVals = serieMob(REF.mob, REF.pct);
const ANIOS = [
  ['2022', 'De las cosechas de 2022 solo se observa a partir del mes 12'],
  ['2023', 'Nueve de las doce cosechas de 2023 superan el patrón de 2024'],
  ['2024', 'Ninguna cosecha de 2024 pasa del 4% al mes 12; en 2023, tres pasaron del 6%'],
  ['2025', 'Enero llega al 6,6%; ninguna otra cosecha de 2025 pasa del 4%'],
  ['2026', 'Las cosechas de 2026 todavía no tienen edad para decir nada'],
];
ANIOS.forEach(([anio, titulo]) => {
  const cos = Object.keys(D.cosechas).filter(c => c.startsWith(anio)).sort();
  if (!cos.length) return;
  const s = pres.addSlide();
  header(s, 'Cosecha por cosecha · ' + anio, titulo, null);

  const cols = 4, gx = 0.14, gy = 0.30;
  const cw = (W - 2 * MX - (cols - 1) * gx) / cols;
  const filas = Math.ceil(cos.length / cols);
  const ch = filas >= 3 ? 1.40 : 2.00;   // una sola fila: cuadro más alto, y el resto lo explica la tarjeta
  cos.forEach((c, i) => {
    const x = MX + (i % cols) * (cw + gx);
    const yy = 1.75 + Math.floor(i / cols) * (ch + gy);
    const v = D.cosechas[c];
    txt(s, cosNice(c), { x, y: yy - 0.24, w: cw * 0.55, h: 0.22, fontSize: 11,
      bold: true, color: INK, fontFace: BODY });
    txt(s, 'COP ' + Math.round(v.colocacion / 1e6) + ' M', { x: x + cw * 0.5, y: yy - 0.24,
      w: cw * 0.5, h: 0.22, fontSize: 9, color: SOFT, fontFace: BODY, align: 'right' });
    s.addChart(pres.ChartType.line, [
      { name: 'Patrón 2024', labels: cada(12), values: refVals },
      { name: cosNice(c), labels: cada(12), values: serieMob(v.mob, v.pct) },
    ], Object.assign({}, baseChart, {
      x, y: yy, w: cw, h: ch, chartColors: ['C8D2DC', BLUE], lineSize: 1.6,
      valAxisMaxVal: 16, valAxisMajorUnit: 8,
      catAxisLabelFontSize: 7.5, valAxisLabelFontSize: 7.5,
      valGridLine: { color: 'F0F3F6', size: 0.75 },
    }));
    // la anotación va DENTRO del cuadro, en la zona vacía de la izquierda
    if (c === '202201') txt(s, 'pico 22,1%,\nfuera de escala', { x: x + 0.42, y: yy + 0.10,
      w: 1.5, h: 0.42, fontSize: 7.5, color: ORANGE, italic: true, fontFace: BODY,
      valign: 'top', lineSpacing: 9 });
  });
  if (cos.length <= 4) {
    card(s, MX, 4.35, W - 2 * MX, 1.95, 'FBF2EC');
    txt(s, 'Por qué esta lámina tiene cuatro cuadros y no doce, y por qué están casi vacíos',
      { x: MX + 0.35, y: 4.6, w: 11.4, h: 0.32, fontSize: 14, bold: true,
        color: '9A4520', fontFace: BODY, isTextBox: true, margin: 0 });
    txt(s, 'De mayo de 2026 en adelante, el archivo no trae el porcentaje calculado en los bloques de observación anteriores al corte de junio, así que esas cuatro cosechas quedan sin curva. No es que no tengan mora: es que el dato no está.',
      { x: MX + 0.35, y: 4.98, w: 11.35, h: 0.52, fontSize: 12, color: '6E4130',
        fontFace: BODY, valign: 'top', lineSpacing: 16, isTextBox: true, margin: 0 });
    txt(s, 'Las cuatro que sí se ven llevan tres meses de vida y van por debajo del 0,6%. Se dibujan en la misma escala que las demás láminas a propósito: así se ve que a esta edad ninguna cosecha ha mostrado todavía lo que va a ser. Las cosechas de 2023, que terminaron siendo las peores del periodo, al tercer mes iban en 0,7%.',
      { x: MX + 0.35, y: 5.55, w: 11.35, h: 0.62, fontSize: 12, color: '6E4130',
        fontFace: BODY, valign: 'top', lineSpacing: 16, isTextBox: true, margin: 0 });
  }
  footer(s, anio === '2026'
    ? 'Misma escala que en las demás láminas del bloque: 0 a 36 meses en el eje horizontal, 0 a 16% en el vertical.'
    : 'Misma escala en los doce cuadros: eje horizontal 0 a 36 meses en libros, eje vertical 0 a 16% de la colocación en mora 30+. La línea gris, idéntica en todos, es el patrón promedio de las cosechas de 2024. La cifra junto a cada título es lo colocado en esa cosecha.');
  s.addNotes('Small multiples: misma escala en los ' + cos.length + ' cuadros, así que las formas se comparan directamente. La línea gris es idéntica en todos: el patrón de 2024.');
});

// ================================================================= 12. HEATMAP
{
  const s = pres.addSlide();
  header(s, 'Todo junto', '52 de las 56 cosechas, en una sola vista',
    'Cada fila es un mes de desembolso; cada columna, un mes de vida. Cuanto más oscuro, más mora. Faltan las cuatro últimas cosechas de 2026: el archivo no trae su porcentaje calculado antes del corte de junio.');
  s.addImage({ path: __dirname + '/heatmap.png', x: MX, y: 2.05, w: 10.8, h: 4.35 });
  const notas = [
    ['Dos bandas oscuras', 'Las cosechas de 2022 y las de abril a septiembre de 2023.'],
    ['El aclaramiento', 'De 2024 en adelante, sostenido. Es la corrección.'],
    ['El blanco de abajo', 'Es futuro, no es cero.'],
    ['El blanco de arriba', 'Tampoco: el archivo empieza a observar en ene-2023.'],
  ];
  let y = 2.1;
  notas.forEach(([t, b]) => {
    txt(s, t, { x: 11.55, y, w: 1.6, h: 0.36, fontSize: 10.5, bold: true, color: INK,
      fontFace: BODY, valign: 'top', lineSpacing: 13 });
    txt(s, b, { x: 11.55, y: y + 0.38, w: 1.6, h: 0.85, fontSize: 9.2, color: MUTED,
      fontFace: BODY, valign: 'top', lineSpacing: 12 });
    y += 1.14;
  });
  footer(s, 'Escala común: 0% a 12% o más de la colocación en mora 30+. Observaciones hasta jun-2026.');
}

// ================================================================= 13. ROLL RATES
{
  const s = pres.addSlide();
  header(s, 'Comportamiento de la mora', 'A los 60 días de mora ya casi no hay vuelta atrás',
    'Hacia dónde se mueve el saldo de un mes al siguiente, según el balde de mora en que esté. Promedio de las transiciones de feb a jun de 2026, ponderado por saldo.');

  const rr = D.roll.filter(r => ['1-30', '31-60', '61-90'].includes(r.desde));
  const sale = r => Math.max(0, Math.round((100 - r.cura - r.igual - r.empeora) * 10) / 10);
  s.addChart(pres.ChartType.bar, [
    { name: 'Se pone al día', labels: rr.map(r => r.desde + ' días'), values: rr.map(r => r.cura) },
    { name: 'Sigue igual', labels: rr.map(r => r.desde + ' días'), values: rr.map(r => r.igual) },
    { name: 'Empeora', labels: rr.map(r => r.desde + ' días'), values: rr.map(r => r.empeora) },
    { name: 'Se paga o sale', labels: rr.map(r => r.desde + ' días'), values: rr.map(sale) },
  ], Object.assign({}, baseChart, {
    x: MX, y: 2.15, w: 7.6, h: 3.7, barDir: 'bar', barGrouping: 'stacked',
    chartColors: [GREEN, '6E7D8C', RED, 'B9C3CD'], valAxisMaxVal: 100, valAxisMajorUnit: 25,
    showValue: true, dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF',
    dataLabelFontSize: 10, dataLabelFontBold: true, dataLabelFormatCode: '0"%"',
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
    showLegend: true, legendPos: 'b', legendFontSize: 10.5, legendColor: MUTED,
  }));

  stat(s, 8.5, 2.15, 4.2, '8,8%', 'del saldo al día cae en mora cada mes',
    'Es el caudal de entrada. Sobre COP 3.094 millones al día, son unos COP 272 millones al mes.', ORANGE);
  card(s, 8.5, 3.97, 4.2, 1.88);
  txt(s, 'Dónde está la palanca', { x: 8.74, y: 4.17, w: 3.8, h: 0.28, fontSize: 12.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, 'Del balde de 1 a 30 días, el 41% se recupera solo. Del de 31 a 60, apenas el 9%. Y del de 61 a 90, más de la mitad rueda al balde siguiente. Pasados los 90 días, el 70% se queda donde está y solo el 1% se pone al día.\n\nLa cobranza que sirve es la de los primeros treinta días. Después de los sesenta, lo que se hace es recuperar garantía, no cartera.',
    { x: 8.74, y: 4.47, w: 3.75, h: 1.35, fontSize: 10.5, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });
  footer(s, 'Se excluye la transición de junio a julio de 2026 porque el castigo la distorsiona. "Empeora" incluye el saldo que pasa a un balde de mayor altura. No se grafica el balde de 91 a 180 días porque su franja de cura es demasiado delgada para etiquetarla; su dato está en el texto.');
  s.addNotes('Esta lámina justifica presupuesto de cobranza temprana. El 41% de cura en el primer balde contra 9% en el segundo es el argumento.');
}

// ================================================================= 14. CARTERA HOY
{
  const s = pres.addSlide();
  header(s, 'La cartera hoy', 'Lo vencido son COP 304 M; la exposición real es 2,4 veces mayor',
    'Cartera viva a 31 de agosto de 2026: ' + D.cartera.n.toLocaleString('es-CO') + ' créditos, COP ' + D.cartera.saldo_M.toLocaleString('es-CO') + ' millones.');

  stat(s, MX, 2.0, 3.85, 'COP ' + D.cartera.vencido_M + ' M', 'Capital vencido',
    'Es el numerador del indicador de cosechas: solo las cuotas ya vencidas y no pagadas.', INK);
  stat(s, MX + 4.05, 2.0, 3.85, 'COP ' + D.cartera.exposicion_M + ' M', 'Exposición total',
    'El saldo completo de esos mismos ' + D.cartera.n_mora + ' créditos con mora de más de 30 días.', RED);
  stat(s, MX + 8.1, 2.0, 3.85, '16,4%', 'De la cartera en mora 30+',
    'Medido como exposición sobre saldo vigente. Es el número que mira el fondeador.', ORANGE);

  const alt = D.altura;
  s.addChart(pres.ChartType.bar, [{ name: 'Saldo', labels: alt.map(a => a.balde),
    values: alt.map(a => a.saldo_M) }], Object.assign({}, baseChart, {
    x: MX, y: 4.15, w: 7.4, h: 2.55, barDir: 'col', chartColors: [BLUE],
    valAxisLabelFormatCode: '#,##0', showValue: true, dataLabelPosition: 'outEnd',
    dataLabelColor: MUTED, dataLabelFontSize: 9.5, dataLabelFormatCode: '#,##0',
    catAxisLabelFontSize: 9.5, valAxisLabelFontSize: 9.5, barGapWidthPct: 45,
  }));
  txt(s, 'Saldo por altura de mora, en millones de COP', { x: MX, y: 3.92, w: 6, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });

  txt(s, 'Lo que hay que tener claro', { x: 8.35, y: 3.92, w: 4.4, h: 0.26, fontSize: 12.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, 'Las dos cifras son correctas y miden cosas distintas. El indicador de cosechas (COP 304 millones) sirve para comparar cosechas entre sí, porque solo cuenta lo vencido. La exposición (COP 728 millones) es lo que de verdad está en riesgo si esos créditos no pagan.\n\nHay COP 607 millones más en el balde de 1 a 30 días, que todavía no cuenta como mora 30+ y del que un 20% rueda hacia adelante cada mes.',
    { x: 8.35, y: 4.24, w: 4.4, h: 2.4, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  s.addNotes('Si alguien pregunta por qué el número de este informe no coincide con el de contabilidad, esta es la lámina: son denominadores y numeradores distintos, y ambos son correctos.');
}

// ================================================================= 15. PUNTO DE VENTA
{
  const s = pres.addSlide();
  header(s, 'Dónde se origina el problema', 'Al mes 6, el peor punto pierde el doble que el mejor',
    'Mora 30+ sobre el capital originado, midiendo todas las cosechas a la misma edad: seis meses de vida. Cosechas de sep-2025 a feb-2026.');

  const NOMBRE = {
    'Cartagena  Suzuki': 'Cartagena · Suzuki', 'Puerto Berrio Akt #2': 'Puerto Berrío · AKT',
    'Aguachica Hero': 'Aguachica · Hero', 'Suzuki Valledupar': 'Valledupar · Suzuki',
    'Pto Berrio': 'Puerto Berrío · Hero', 'Pto Boyaca': 'Puerto Boyacá',
    'Suzuki Pto Berrio': 'Puerto Berrío · Suzuki',
  };
  const p = [...D.punto_mob6].map(d => ({ ...d, punto: NOMBRE[d.punto] || d.punto }))
    .sort((a, b) => b.pct - a.pct);
  s.addChart(pres.ChartType.bar, [{ name: 'Mora 30+ al mes 6', labels: p.map(d => d.punto),
    values: p.map(d => d.pct) }], Object.assign({}, baseChart, {
    x: MX, y: 2.12, w: 7.9, h: 3.85, barDir: 'bar', chartColors: [BLUE],
    valAxisMaxVal: 14, valAxisMajorUnit: 3.5, showValue: true, dataLabelPosition: 'outEnd',
    dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
    dataLabelFormatCode: '0.0"%"', catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10,
    barGapWidthPct: 40,
  }));

  txt(s, 'No es la plaza, es la operación', { x: 8.75, y: 2.12, w: 4.0, h: 0.28,
    fontSize: 13, bold: true, color: INK, fontFace: BODY });
  txt(s, 'En Puerto Berrío operan tres líneas distintas y rinden distinto entre sí: Suzuki va en 5,8%, Hero en 8,1% y AKT en 12,0%, con la misma ciudad y la misma clientela. La diferencia está en cómo origina cada punto, no en el mercado que atiende.',
    { x: 8.75, y: 2.44, w: 4.0, h: 1.45, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });

  card(s, 8.75, 3.97, 4.0, 2.0, 'FBF2EC');
  txt(s, 'COP 71 M al año', { x: 8.99, y: 4.15, w: 3.6, h: 0.5, fontSize: 24, bold: true,
    color: '9A4520', fontFace: HEAD });
  txt(s, 'menos en mora si Cartagena-Suzuki y Puerto Berrío AKT rindieran como el promedio de los demás',
    { x: 8.99, y: 4.68, w: 3.55, h: 0.7, fontSize: 11, color: '6E4130', fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  txt(s, 'Esos dos puntos ponen el 33% del capital y el 42% de la mora. La brecha son COP 22 millones sobre las cinco cosechas medidas; anualizada al ritmo de colocación de 2026, COP 71 millones.',
    { x: 8.99, y: 5.38, w: 3.55, h: 0.72, fontSize: 9.5, color: '8A6552', fontFace: BODY,
      valign: 'top', lineSpacing: 12 });
  footer(s, 'Se excluye Administración por tener menos de 20 créditos. Denominador: capital inicial de los créditos vigentes, que a los 6 meses cubre el 81% de lo colocado.');
  s.addNotes('La comparación es a edad constante, que es lo que la hace defendible. Si preguntan por qué no se ven todas las cosechas: a los 6 meses la cobertura es del 81%, más allá el sesgo de supervivencia la arruina.');
}

// ================================================================= 16. CONCENTRACIÓN
{
  const s = pres.addSlide();
  header(s, 'Cobranza', 'La décima parte de la cartera tiene tres cuartas partes de la mora',
    'Cartera viva a ago-2026, ordenada por saldo en mora de cada crédito.');

  stat(s, MX, 1.95, 3.85, D.concentracion.creditos + ' créditos', 'concentran el ' + nf(D.concentracion.pct_mora) + '% de la mora',
    'Son el 10% de la cartera y suman COP ' + D.concentracion.monto_M + ' millones de exposición.', RED);
  stat(s, MX + 4.05, 1.95, 3.85, D.cartera.n_mora + ' créditos', 'están en mora de más de 30 días',
    'El 19% de los ' + D.cartera.n.toLocaleString('es-CO') + ' créditos vivos.', INK);
  stat(s, MX + 8.1, 1.95, 3.85, '65%', 'de la cartera es cosecha 2026',
    'Y va en 5,4% de ICV. El grueso del riesgo vivo es originación reciente y todavía joven.', INK);

  const e = D.exposicion.filter(x => +x.anio >= 2023);
  s.addChart(pres.ChartType.bar, [{ name: 'ICV', labels: e.map(x => 'Cosechas ' + x.anio),
    values: e.map(x => x.icv) }], Object.assign({}, baseChart, {
    x: MX, y: 4.15, w: 6.0, h: 2.15, barDir: 'col', chartColors: [ORANGE],
    valAxisMaxVal: 100, valAxisMajorUnit: 25, showValue: true, dataLabelPosition: 'outEnd',
    dataLabelColor: INK, dataLabelFontSize: 10, dataLabelFontBold: true,
    dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10, valAxisLabelFontSize: 9.5,
    barGapWidthPct: 55,
  }));
  txt(s, 'Mora sobre saldo vivo, por año de cosecha', { x: MX, y: 3.92, w: 6, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });

  txt(s, 'Por qué las cosechas viejas muestran un ICV tan alto', { x: 7.0, y: 3.92, w: 5.7,
    h: 0.26, fontSize: 12.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'No es que se hayan deteriorado más de lo que dice el análisis de cosechas. Es que de esas cosechas ya se pagó casi todo lo bueno: lo que queda vivo es, por definición, lo que no ha pagado. De 2023 solo quedan COP 85 millones vivos, y son mora casi en su totalidad.\n\nPor eso el ICV sirve para dimensionar lo que hay que cobrar hoy, y el análisis de cosechas para juzgar cómo se originó.',
    { x: 7.0, y: 4.24, w: 5.7, h: 2.2, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  s.addNotes('La concentración es el argumento para focalizar cobranza: 122 créditos, no 1.213.');
}

// ================================================================= 17. COLOCACIÓN
{
  const s = pres.addSlide();
  header(s, 'Crecimiento', 'Se creció 2,3 veces sin perder calidad',
    'Monto colocado por cosecha mensual, en millones de COP.');
  s.addChart(pres.ChartType.bar, [{ name: 'Colocación',
    labels: ejeCosechas(D.colocacion.map(d => d.cosecha)), values: D.colocacion.map(d => d.monto_M) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.0, w: W - 2 * MX, h: 3.45, barDir: 'col', chartColors: [BLUE],
      valAxisLabelFormatCode: '#,##0', catAxisLabelFontSize: 11,
      valAxisLabelFontSize: 10, barGapWidthPct: 30,
    }));
  const kk = [['COP 197 M', 'promedio mensual en 2022'], ['COP 450 M', 'promedio mensual en 2026'],
              ['2,3x', 'más volumen, con mejor mora']];
  kk.forEach(([a, b], i) => {
    const x = MX + i * 4.05;
    txt(s, a, { x, y: 5.75, w: 3.8, h: 0.45, fontSize: 22, bold: true, color: INK, fontFace: HEAD });
    txt(s, b, { x, y: 6.2, w: 3.8, h: 0.3, fontSize: 11, color: MUTED, fontFace: BODY });
  });
  footer(s, 'Crecer y mejorar la calidad al mismo tiempo es lo difícil, y es lo que muestran los datos. El límite hoy no lo pone el riesgo agregado: lo ponen dos o tres puntos de venta.');
}

// ================================================================= 18. LO QUE NO SABEMOS
{
  const s = pres.addSlide();
  header(s, 'Límites del análisis', 'Lo que este archivo no permite responder',
    'Tres campos ausentes, y una consecuencia concreta de cada uno.');

  const gaps = [
    ['Cuota inicial y precio real de la moto', 'La columna VALORGARAN viene casi igual al capital prestado, así que no permite calcular el LTV — el driver de riesgo número uno en financiación de motos.',
     'No se puede decir cuánta cuota inicial hay que exigir.'],
    ['Score de buró en la originación', 'El archivo es el extracto de cartera del core, no un reporte de buró. No trae score ni endeudamiento externo del deudor.',
     'No se puede fijar el punto de corte de aprobación.'],
    ['Asesor comercial', 'La columna NOMBREASES viene vacía en los siete cortes mensuales.',
     'Se puede señalar el punto de venta, pero no a quién dentro de él.'],
  ];
  gaps.forEach(([t, b, c], i) => {
    const y = 2.05 + i * 1.44;
    card(s, MX, y, W - 2 * MX, 1.26);
    txt(s, t, { x: MX + 0.3, y: y + 0.17, w: 4.5, h: 0.3, fontSize: 13, bold: true,
      color: INK, fontFace: BODY });
    txt(s, b, { x: MX + 0.3, y: y + 0.48, w: 4.6, h: 0.72, fontSize: 10.2, color: MUTED,
      fontFace: BODY, valign: 'top', lineSpacing: 13 });
    txt(s, 'Consecuencia', { x: MX + 5.3, y: y + 0.17, w: 6.4, h: 0.26, fontSize: 10,
      bold: true, color: ORANGE, fontFace: BODY, charSpacing: 0.8 });
    txt(s, c, { x: MX + 5.3, y: y + 0.48, w: 6.4, h: 0.7, fontSize: 13, color: INK,
      fontFace: BODY, valign: 'top', lineSpacing: 18 });
  });
  txt(s, 'Con esos tres campos, este mismo análisis pasa de diagnosticar a fijar política de originación.',
    { x: MX, y: 6.5, w: 11.5, h: 0.35, fontSize: 12, color: INK, italic: true, fontFace: BODY });
  s.addNotes('Nunca cerrar un informe sin decir qué no se pudo responder. Es lo que evita que la junta asuma que lo que no se mencionó está bien.');
}

// ================================================================= 19. DECISIONES (dark)
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, 'LO QUE SE SOMETE A APROBACIÓN', { x: MX, y: 0.62, w: 11, h: 0.26, fontSize: 11,
    color: '7FA9DC', bold: true, charSpacing: 1.8, fontFace: BODY });
  txt(s, 'Tres decisiones', { x: MX, y: 0.95, w: 11, h: 0.6, fontSize: 30, bold: true,
    color: WHITE, fontFace: HEAD });

  const dec = [
    ['Intervenir Cartagena-Suzuki y Puerto Berrío AKT', 'Auditoría de originación de los dos puntos con peor mora al mes 6: expediente, avalúo, verificación de ingresos y cuota inicial efectiva de las cosechas 2025–2026.', 'Gerencia comercial', '31 de octubre', 'COP 71 M al año'],
    ['Cuadrar el castigo de julio con contabilidad', 'Confirmar qué se castigó, contra qué provisión, con qué autorización y con qué efecto en el P&G del año. Sin esto no se reporta mejora de mora a la junta ni al fondeador.', 'Contabilidad', '15 de octubre', 'COP 549 M'],
    ['Incorporar tres campos al reporte mensual', 'Cuota inicial y precio de factura, score de buró en la originación, y asesor comercial. Que el corte de octubre ya salga con ellos.', 'Sistemas y crédito', '30 de noviembre', 'Habilita fijar política'],
  ];
  dec.forEach(([t, b, quien, cuando, monto], i) => {
    const y = 1.9 + i * 1.55;
    card(s, MX, y, W - 2 * MX, 1.35, '1B3450');
    s.addShape(pres.ShapeType.ellipse, { x: MX + 0.28, y: y + 0.24, w: 0.42, h: 0.42,
      fill: { color: '2F5580' }, line: { color: '2F5580', width: 0 } });
    txt(s, String(i + 1), { x: MX + 0.28, y: y + 0.24, w: 0.42, h: 0.42, fontSize: 15,
      bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: BODY });
    txt(s, t, { x: MX + 0.92, y: y + 0.19, w: 7.1, h: 0.32, fontSize: 14, bold: true,
      color: WHITE, fontFace: BODY });
    txt(s, b, { x: MX + 0.92, y: y + 0.54, w: 7.2, h: 0.8, fontSize: 10.5, color: 'A8C0DC',
      fontFace: BODY, valign: 'top', lineSpacing: 14 });
    const meta = [['Responsable', quien], ['Fecha', cuando], ['En juego', monto]];
    meta.forEach(([k, v], j) => {
      const mx = MX + 8.25 + j * 1.38;
      txt(s, k.toUpperCase(), { x: mx, y: y + 0.28, w: 1.32, h: 0.2, fontSize: 7.5,
        color: '6E8CAE', bold: true, charSpacing: 0.6, fontFace: BODY });
      txt(s, v, { x: mx, y: y + 0.52, w: 1.32, h: 0.7, fontSize: 10, bold: true,
        color: WHITE, fontFace: BODY, valign: 'top', lineSpacing: 13 });
    });
  });
  s.addNotes('Cada decisión tiene dueño, fecha y monto. Si la junta quiere aprobar solo una, la número 2 es la urgente: sin ella no se puede reportar el cierre.');
}

// ================================================================= 20. ANEXO METODOLOGÍA
{
  const s = pres.addSlide();
  header(s, 'Anexo', 'Metodología y controles',
    'Todo el análisis es reproducible desde el archivo original con un solo comando.');

  card(s, MX, 1.9, 6.0, 2.35);
  txt(s, 'Cómo se calcula', { x: MX + 0.28, y: 2.1, w: 5.5, h: 0.28, fontSize: 13,
    bold: true, color: INK, fontFace: BODY });
  txt(s, 'Cosecha: mes de desembolso. MOB: meses desde el desembolso.\n\nIndicador: capital vencido (columna SALDO EN MORA) de los créditos con más de 30 días de mora, sobre la colocación original de la cosecha. El denominador es fijo, y por eso las cosechas son comparables entre sí.\n\nLas curvas por año ponderan cada cosecha por su monto colocado; no promedian porcentajes.',
    { x: MX + 0.28, y: 2.44, w: 5.5, h: 1.7, fontSize: 10.2, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });

  card(s, MX + 6.3, 1.9, 6.0, 2.35);
  txt(s, 'Controles que se corrieron', { x: MX + 6.58, y: 2.1, w: 5.5, h: 0.28,
    fontSize: 13, bold: true, color: INK, fontFace: BODY });
  txt(s, 'Numerador: la suma del capital vencido de los créditos con más de 30 días de mora en ago-2026 da COP 301.160.233, idéntico peso a peso al total de la hoja Consolidado.\n\nDenominador: las 56 cosechas mantienen la misma colocación en los 13 bloques de observación.\n\nTriángulo contra Consolidado: cuadran en 622 de 638 celdas.',
    { x: MX + 6.58, y: 2.44, w: 5.5, h: 1.7, fontSize: 10.2, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });

  txt(s, 'Salvedades', { x: MX, y: 4.45, w: 6, h: 0.28, fontSize: 13, bold: true,
    color: INK, fontFace: BODY });
  const sal = [
    'Las 16 celdas que no cuadran están todas en la columna julio-2025 de la hoja Cosechas, con valores redondeados a tres decimales: esa columna se digitó a mano. Se usó el Consolidado como fuente.',
    'La hoja ENERO 26 está vacía y el Consolidado no tiene bloque de enero de 2026. Ese mes es s/d, no cero.',
    'Las hojas de febrero y abril traen otro formato (99 y 100 columnas contra 105), sin CAPITALINI y con SALDO EN MORA escrito con un espacio al final. Se homologaron.',
    'Los cortes por punto de venta se miden a los 6 meses de vida, donde el capital de los créditos vigentes todavía cubre el 81% de lo colocado. Más allá, el sesgo de supervivencia los vuelve inservibles.',
  ];
  sal.forEach((t, i) => {
    txt(s, t, { x: MX, y: 4.78 + i * 0.47, w: 12.1, h: 0.44, fontSize: 10, color: MUTED,
      fontFace: BODY, valign: 'top', bullet: { code: '2013' }, lineSpacing: 13 });
  });
  footer(s, 'Reproducible con: python3 analizar_cosechas.py Cosechas_2026.xlsx — imprime los controles, reescribe las tablas y actualiza el tablero.');
}

// ================================================================= 21. ANEXO TABLA
{
  const s = pres.addSlide();
  header(s, 'Anexo', 'Mora por cosecha al mes 6 y al mes 12',
    'Las 56 cosechas, con su colocación en millones de COP. s/d = la cosecha todavía no alcanzó esa edad; no es un cero.');
  const HD = ['Cosecha', 'Colocado', 'Mes 6', 'Mes 12'];
  const rows = [[...HD, ...HD, ...HD].map(h => ({ text: h, options: { bold: true, color: MUTED } }))];
  const d = D.mob_fijo, per = Math.ceil(d.length / 3);
  for (let i = 0; i < per; i++) {
    const cell = o => o ? [cosNice(o.cosecha), Math.round(o.colocacion / 1e6).toLocaleString('es-CO'),
      o.m6 == null ? 's/d' : nf(o.m6) + '%', o.m12 == null ? 's/d' : nf(o.m12) + '%'] : ['', '', '', ''];
    rows.push([...cell(d[i]), ...cell(d[i + per]), ...cell(d[i + 2 * per])]);
  }
  s.addTable(rows, { x: MX, y: 2.0, w: W - 2 * MX,
    colW: [1.0, 1.15, 0.87, 0.95, 1.0, 1.15, 0.87, 0.95, 1.0, 1.15, 0.87, 0.95],
    fontSize: 8.2, fontFace: BODY, color: INK, border: { type: 'solid', color: 'E8ECF1', pt: 0.5 },
    align: 'right', valign: 'middle', autoPage: false, fill: { color: 'FFFFFF' },
  });
  footer(s, 'Tablas completas en cosechas_districolmotos/datos/ — un CSV por gráfico de este informe.');
}

const OUT = __dirname + '/Cosechas_Districolmotos_Junta.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('listo:', OUT));
