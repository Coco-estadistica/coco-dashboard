// Analítica de riesgo para comité de crédito y cartera — Districolmotos
// Genera Cosechas_Districolmotos_Comite.pptx a partir de datos_comite.json
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const D = JSON.parse(fs.readFileSync(__dirname + '/datos_comite.json', 'utf8'));

const NAVY = '12263A', INK = '18242F', MUTED = '62717F', PIE = '6B7785';
const WHITE = 'FFFFFF', SURF = 'F2F5F8', LINE = 'DDE3EA';
const BLUE = '2A78D6', ORANGE = 'EB6834', RED = 'C4342F', GREEN = '0F7B3F';
const HEAD = 'Cambria', BODY = 'Calibri';
const MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const cosNice = c => MES[+c.slice(4) - 1] + '-' + c.slice(0, 4);
const nf = (v, d = 1) => v == null ? 's/d'
  : v.toLocaleString('es-CO', { minimumFractionDigits: d, maximumFractionDigits: d });

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Districolmotos';
pres.title = 'Analítica de riesgo — comité de crédito y cartera';
const W = 13.333, H = 7.5, MX = 0.62;

function txt(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, margin: 0 }, o)); }
function header(s, kicker, title, sub) {
  txt(s, kicker.toUpperCase(), { x: MX, y: 0.34, w: 11, h: 0.22, fontSize: 10.5,
    color: BLUE, bold: true, charSpacing: 1.6, fontFace: BODY });
  txt(s, title, { x: MX, y: 0.58, w: W - 2 * MX, h: 0.80, fontSize: 23,
    bold: true, color: INK, fontFace: HEAD, valign: 'top', lineSpacing: 28 });
  if (sub) txt(s, sub, { x: MX, y: 1.50, w: W - 2 * MX, h: 0.50, fontSize: 11.5,
    color: MUTED, fontFace: BODY, valign: 'top', lineSpacing: 15 });
}
function footer(s, t) {
  txt(s, t, { x: MX, y: H - 0.95, w: W - 2 * MX, h: 0.55, fontSize: 9,
    color: PIE, fontFace: BODY, valign: 'top', lineSpacing: 12 });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || SURF }, line: { color: fill ? fill : LINE, width: 0.75 } });
}
function stat(s, x, y, w, value, label, note, color) {
  card(s, x, y, w, 1.62);
  txt(s, value, { x: x + 0.24, y: y + 0.18, w: w - 0.4, h: 0.62, fontSize: 30,
    bold: true, color: color || INK, fontFace: HEAD });
  txt(s, label, { x: x + 0.24, y: y + 0.84, w: w - 0.4, h: 0.3, fontSize: 11.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, note, { x: x + 0.24, y: y + 1.13, w: w - 0.4, h: 0.42, fontSize: 9.8,
    color: MUTED, fontFace: BODY, valign: 'top', lineSpacing: 12 });
}
const baseChart = {
  showLegend: false, showTitle: false, valAxisMinVal: 0,
  valGridLine: { color: 'EDF0F3', size: 0.75 }, catGridLine: { style: 'none' },
  catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
  catAxisLineShow: false, valAxisLineShow: false,
  lineSmooth: false, lineDataSymbol: 'none', dataBorder: { pt: 1, color: 'FFFFFF' },
  chartColors: [BLUE], lineSize: 2, valAxisLabelFormatCode: '0"%"',
};

// ========================================================== 1. PORTADA
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, 'COMITÉ DE CRÉDITO Y CARTERA · SEPTIEMBRE DE 2026', { x: MX, y: 1.5, w: 11, h: 0.3,
    fontSize: 11, color: '7FA9DC', bold: true, charSpacing: 2, fontFace: BODY });
  txt(s, 'Lo único que predice la mora\nes el paso del tiempo',
    { x: MX, y: 2.0, w: 11.2, h: 1.7, fontSize: 34, bold: true, color: WHITE,
      fontFace: HEAD, lineSpacing: 45, valign: 'top' });
  txt(s, 'Nueve análisis sobre los mismos datos, y lo que cada uno permite decidir',
    { x: MX, y: 3.85, w: 10, h: 0.36, fontSize: 15, color: 'CADCFC', fontFace: BODY });
  [['853 créditos', 'muestra del modelo, dentro de su plazo'],
   ['5 transiciones', 'mes a mes, excluyendo el castigo'],
   ['31-ago-2026', 'fecha de corte']].forEach(([a, b], i) => {
    const x = MX + i * 3.7;
    txt(s, a, { x, y: 4.75, w: 3.4, h: 0.4, fontSize: 19, bold: true, color: WHITE, fontFace: HEAD });
    txt(s, b, { x, y: 5.16, w: 3.4, h: 0.5, fontSize: 10.5, color: '9FB6D4', fontFace: BODY,
      valign: 'top', lineSpacing: 13 });
  });
  txt(s, 'Complementa el informe a junta. Aquí no se resume el trimestre: se prueba qué mueve el riesgo y qué parámetros de política cambiarlo implica.',
    { x: MX, y: 6.6, w: 11.5, h: 0.4, fontSize: 9.5, color: '7189A5', fontFace: BODY });
  s.addNotes('Este mazo es técnico. El mensaje incómodo está en la lámina 3: ninguna variable del archivo discrimina riesgo salvo la edad del crédito.');
}

// ========================================================== 2. MAPA DE COBERTURA
{
  const s = pres.addSlide();
  header(s, 'Punto de partida', 'De las seis decisiones del comité, este archivo responde dos',
    'Antes de mostrar un solo gráfico: qué se puede decidir con estos datos y qué no.');
  const filas = [
    ['Cobranza', 'Dónde y cuándo poner el esfuerzo', 'SÍ', GREEN, 'Curva de hazard, de cura y concentración'],
    ['Provisión', 'Cuánto falta por aflorar de lo ya colocado', 'SÍ', GREEN, 'Proyección de cosechas y migración de calificación'],
    ['Límites', 'A quién le prestamos demasiado', 'PARCIAL', ORANGE, 'Hay exposición por deudor, no ingresos ni deuda externa'],
    ['Fondeo', 'Cuánto dura de verdad la cartera', 'PARCIAL', ORANGE, 'Se ve el prepago, no el calendario de pagos'],
    ['Originación', 'A quién prestar, cuánto y a qué plazo', 'NO', RED, 'Ninguna variable del archivo discrimina riesgo'],
    ['Punto de corte', 'Dónde poner la línea de aprobación', 'NO', RED, 'No hay score, ni cuota inicial, ni solicitudes rechazadas'],
  ];
  const rows = [[{ text: 'Decisión', options: { bold: true, color: MUTED } },
    { text: 'La pregunta', options: { bold: true, color: MUTED } },
    { text: '¿Se responde?', options: { bold: true, color: MUTED } },
    { text: 'Con qué, o por qué no', options: { bold: true, color: MUTED } }]];
  filas.forEach(([a, b, c, col, e]) => rows.push([
    { text: a, options: { bold: true } }, { text: b },
    { text: c, options: { bold: true, color: col } }, { text: e }]));
  s.addTable(rows, { x: MX, y: 2.0, w: W - 2 * MX, colW: [2.0, 3.5, 1.45, 5.14],
    fontSize: 11.5, fontFace: BODY, color: INK, rowH: 0.52, valign: 'middle',
    border: { type: 'solid', color: 'E8ECF1', pt: 0.5 }, fill: { color: 'FFFFFF' } });
  txt(s, 'Las dos filas rojas son las que más plata valen, y son precisamente las que este archivo no permite responder. La lámina 13 dice qué dato haría falta para cada una.',
    { x: MX, y: 6.0, w: 11.8, h: 0.5, fontSize: 12, color: INK, italic: true,
      fontFace: BODY, valign: 'top', lineSpacing: 16 });
  s.addNotes('Empezar por aquí evita que el comité asuma que lo que no se mencionó está bien medido.');
}

// ========================================================== 3. QUÉ DISCRIMINA
{
  const s = pres.addSlide();
  header(s, 'El hallazgo central', 'Ninguna característica del crédito separa buenos de malos',
    `Regresión logística sobre ${D.muestra_logit.n} créditos vivos dentro de su plazo pactado, cosechas 2024-2026. Prueba de razón de verosimilitud contra el modelo que solo usa la edad.`);
  const rows = [[{ text: 'Variable', options: { bold: true, color: MUTED } },
    { text: 'Valor p', options: { bold: true, color: MUTED } },
    { text: '¿Discrimina?', options: { bold: true, color: MUTED } },
    { text: 'Detalle', options: { bold: true, color: MUTED } }]];
  D.discrimina.forEach(v => {
    const sig = v.p < 0.05, marg = v.p >= 0.05 && v.p < 0.10;
    rows.push([{ text: v.variable, options: { bold: true } },
      { text: v.p < 0.001 ? '< 0,001' : nf(v.p, 3) },
      { text: sig ? 'Sí' : marg ? 'Marginal' : 'No', options: { bold: true, color: sig ? GREEN : marg ? ORANGE : RED } },
      { text: v.nota }]);
  });
  s.addTable(rows, { x: MX, y: 2.15, w: 8.0, colW: [2.5, 1.3, 1.5, 2.7],
    fontSize: 11.5, fontFace: BODY, color: INK, rowH: 0.48, valign: 'middle',
    border: { type: 'solid', color: 'E8ECF1', pt: 0.5 }, fill: { color: 'FFFFFF' } });

  txt(s, 'Qué significa esto', { x: 8.9, y: 2.15, w: 3.85, h: 0.3, fontSize: 13.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, 'La edad del crédito es el único predictor sólido: cada mes adicional multiplica por 1,10 la probabilidad de estar en mora 30+.\n\nEl punto de venta queda en 0,083 — sugestivo, pero no concluyente con esta muestra.\n\nTasa, monto, plazo y codeudor no aportan nada. Con lo que hay en el archivo no se puede construir una política de originación: no porque no exista diferencia, sino porque los campos que la revelarían no están.',
    { x: 8.9, y: 2.52, w: 3.85, h: 3.3, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  footer(s, 'Se excluyen los créditos vivos pasados de su plazo pactado (86 de 1.213): siguen en el libro precisamente porque no pagaron, y el 93% está en mora. Incluirlos fabricaría una relación entre plazo y riesgo que no existe.');
  s.addNotes('Esta lámina desarma cualquier propuesta de política basada en las variables actuales. Es el argumento para pedir los campos que faltan.');
}

// ========================================================== 4. HAZARD
{
  const s = pres.addSlide();
  header(s, 'Cobranza · cuándo', 'El riesgo de caer en mora sube con la edad; no hay zona segura',
    'De los créditos que estaban al día, qué porcentaje entra en mora al mes siguiente, según los meses que llevan en libros. Promedio de cinco transiciones mensuales.');
  s.addChart(pres.ChartType.bar, [{ name: 'Entra en mora',
    labels: D.hazard.map(h => h.edad + ' meses'), values: D.hazard.map(h => h.pct) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.15, w: 7.8, h: 3.75, barDir: 'col', chartColors: [BLUE],
      valAxisMaxVal: 20, valAxisMajorUnit: 5, showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
      dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
      barGapWidthPct: 40,
    }));
  txt(s, 'Lo contrario de lo que se espera', { x: 8.75, y: 2.15, w: 4.0, h: 0.3,
    fontSize: 13.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'En crédito de consumo lo normal es que el riesgo haga pico entre los meses 6 y 12 y después ceda: el que sobrevivió el primer año ya demostró que paga.\n\nAquí pasa al revés. Un crédito de más de 18 meses tiene 15,4% de probabilidad mensual de entrar en mora, contra 10,6% de uno recién desembolsado.\n\nLa explicación está en la lámina siguiente.',
    { x: 8.75, y: 2.52, w: 4.0, h: 2.6, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  card(s, 8.75, 5.15, 4.0, 0.75, 'FBF2EC');
  txt(s, 'Decisión: la cobranza preventiva no puede concentrarse solo en los primeros meses.',
    { x: 8.99, y: 5.32, w: 3.55, h: 0.5, fontSize: 11, bold: true, color: '9A4520',
      fontFace: BODY, valign: 'top', lineSpacing: 14 });
  footer(s, 'Entrar en mora = pasar de 0 días a 1 o más. Se excluye la transición jun→jul 2026 por el castigo. Solo se muestran tramos con 30 o más créditos expuestos.');
}

// ========================================================== 5. SELECCIÓN ADVERSA
{
  const s = pres.addSlide();
  header(s, 'Por qué sube el riesgo con la edad', 'Los que pagan bien se van antes, y el libro se va quedando con el resto',
    'Porcentaje de créditos sanos que salen del libro cada mes —pagan anticipado o cancelan— según su edad.');
  s.addChart(pres.ChartType.bar, [{ name: 'Sale del libro',
    labels: D.prepago.map(p => p.edad + ' meses'), values: D.prepago.map(p => p.pct) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.15, w: 7.8, h: 3.6, barDir: 'col', chartColors: [GREEN],
      valAxisMaxVal: 16, valAxisMajorUnit: 4, showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
      dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
      barGapWidthPct: 45,
    }));
  txt(s, 'Selección adversa por atrición', { x: 8.75, y: 2.15, w: 4.0, h: 0.3,
    fontSize: 13.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'Un crédito sano de 18 a 23 meses tiene 12,9% de probabilidad mensual de salir del libro, contra 3,2% de uno nuevo. Los buenos clientes prepagan.\n\nEl efecto es mecánico: si cada mes se van los que pagan, lo que queda concentra a los que no. Por eso el hazard de la lámina anterior sube, sin que la calidad de la originación haya empeorado.\n\nNo es un problema de riesgo, es un problema de medición: cualquier indicador calculado sobre créditos vivos hereda este sesgo.',
    { x: 8.75, y: 2.52, w: 4.0, h: 3.3, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  card(s, 8.75, 5.95, 4.0, 0.85, 'FBF2EC');
  txt(s, 'Decisión: medir siempre a edad constante. Todo indicador sobre créditos vivos exagera el deterioro.',
    { x: 8.99, y: 6.12, w: 3.55, h: 0.6, fontSize: 11, bold: true, color: '9A4520',
      fontFace: BODY, valign: 'top', lineSpacing: 14, isTextBox: true, margin: 0 });
  footer(s, 'Se cuentan como salida los créditos con 30 días de mora o menos que desaparecen del corte siguiente. Los castigos de julio se excluyen.');
  s.addNotes('Esta lámina es la que explica por qué no hay que alarmarse con el hazard creciente, y por qué el ICV de las cosechas viejas se ve tan mal.');
}

// ========================================================== 6. CURA
{
  const s = pres.addSlide();
  header(s, 'Cobranza · hasta cuándo', 'Después de los 180 días no se recuperó un solo crédito',
    'De los créditos que estaban en cada altura de mora, qué porcentaje bajó de 30 días al mes siguiente.');
  s.addChart(pres.ChartType.bar, [{ name: 'Sale de mora 30+',
    labels: D.cura.map(c => c.altura + ' días'), values: D.cura.map(c => c.pct) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.15, w: 7.8, h: 3.75, barDir: 'col', chartColors: [GREEN],
      valAxisMaxVal: 40, valAxisMajorUnit: 10, showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
      dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10,
      barGapWidthPct: 45,
    }));
  txt(s, 'El acantilado está a los 90 días', { x: 8.75, y: 2.15, w: 4.0, h: 0.3,
    fontSize: 13.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'Del balde de 31 a 60 días se recupera casi un tercio. Del de 61 a 90, uno de cada ocho. Pasados los 90, uno de cada treinta.\n\nDe los 1.036 créditos observados con más de 180 días de mora, ninguno volvió a estar por debajo de 30 días. Cero, en cinco transiciones.\n\nEso convierte los 180 días en una frontera operativa: después de ahí lo que queda es recuperar la garantía, y el crédito debería estar castigado o en proceso jurídico.',
    { x: 8.75, y: 2.52, w: 4.0, h: 2.9, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  card(s, 8.75, 5.45, 4.0, 0.75, 'FBF2EC');
  txt(s, 'Decisión: concentrar cobranza en 31-90 días y definir política de castigo a 180.',
    { x: 8.99, y: 5.62, w: 3.55, h: 0.5, fontSize: 11, bold: true, color: '9A4520',
      fontFace: BODY, valign: 'top', lineSpacing: 14 });
  footer(s, 'Promedio de cinco transiciones mensuales, medido en número de créditos. Solo se muestran tramos con 20 o más observaciones.');
}

// ========================================================== 7. MIGRACIÓN DE CALIFICACIÓN
{
  const s = pres.addSlide();
  header(s, 'Provisión', 'La calificación B es la bisagra: un tercio mejora, un tercio cae',
    'A dónde va el saldo de cada calificación al mes siguiente. Promedio de cinco transiciones, ponderado por saldo.');
  const CAL = ['A', 'B', 'C', 'D', 'E'];
  const rows = [[{ text: 'Desde', options: { bold: true, color: MUTED } },
    ...CAL.map(c => ({ text: 'a ' + c, options: { bold: true, color: MUTED } })),
    { text: 'Sale', options: { bold: true, color: MUTED } }]];
  D.migracion.forEach(m => {
    rows.push([{ text: m.desde, options: { bold: true } },
      ...CAL.map(c => {
        const v = m[c] ?? 0, esDiag = c === m.desde;
        return { text: v === 0 ? '—' : nf(v) + '%',
          options: { bold: v >= 20, color: v === 0 ? 'C8CFD6' : esDiag ? INK : (CAL.indexOf(c) > CAL.indexOf(m.desde) ? RED : GREEN) } };
      }),
      { text: nf(m.Sale) + '%', options: { color: MUTED } }]);
  });
  s.addTable(rows, { x: MX, y: 2.15, w: 7.6, colW: [1.1, 1.08, 1.08, 1.08, 1.08, 1.08, 1.1],
    fontSize: 12, fontFace: BODY, color: INK, rowH: 0.55, valign: 'middle', align: 'center',
    border: { type: 'solid', color: 'E8ECF1', pt: 0.5 }, fill: { color: 'FFFFFF' } });
  txt(s, 'Verde: mejora · Rojo: se deteriora', { x: MX, y: 5.7, w: 6, h: 0.26,
    fontSize: 10, color: MUTED, fontFace: BODY });

  txt(s, 'Dos lecturas para provisión', { x: 8.5, y: 2.15, w: 4.25, h: 0.3,
    fontSize: 13.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'La A es estable: 94,6% sigue en A. La E es absorbente: 98% se queda ahí y nada sale.\n\nLa B es donde se decide todo. Un tercio vuelve a A, un tercio se queda y un tercio cae a C. Y desde C, el 53% cae a D en un solo mes.\n\nEn términos de gestión: el crédito calificado B todavía se puede salvar; el calificado C ya está en caída y conviene provisionarlo como si fuera D.',
    { x: 8.5, y: 2.52, w: 4.25, h: 3.2, fontSize: 11, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 15 });
  card(s, 8.5, 5.95, 4.25, 0.85, 'FBF2EC');
  txt(s, 'Decisión: provisionar la calificación C como si fuera D.',
    { x: 8.74, y: 6.18, w: 3.8, h: 0.45, fontSize: 11.5, bold: true, color: '9A4520',
      fontFace: BODY, valign: 'top', lineSpacing: 14, isTextBox: true, margin: 0 });
  footer(s, 'Se excluye la transición jun→jul 2026. "Sale" es saldo que desaparece del corte siguiente por pago, cancelación o castigo.');
}

// ========================================================== 8. PROYECCIÓN
{
  const s = pres.addSlide();
  const t = D.proy_total;
  header(s, 'Provisión · cuánto falta', `De lo ya colocado faltan por aflorar unos COP ${t.falta_M} millones de mora`,
    `Proyección del pico de mora de las ${t.n_cosechas} cosechas que aún no maduran, aplicando los factores de desarrollo de las 24 cosechas que ya llegaron a los 25 meses.`);
  const P = D.proyeccion;
  s.addChart(pres.ChartType.bar, [
    { name: 'Mora observada hoy', labels: P.map(p => cosNice(p.cosecha)), values: P.map(p => p.hoy) },
    { name: 'Falta por aflorar (proyectado)', labels: P.map(p => cosNice(p.cosecha)),
      values: P.map(p => Math.max(0, +(p.proy - p.hoy).toFixed(2))) },
  ], Object.assign({}, baseChart, {
    x: MX, y: 2.15, w: 8.1, h: 3.7, barDir: 'col', barGrouping: 'stacked',
    chartColors: [BLUE, 'C8D2DC'], valAxisMaxVal: 12, valAxisMajorUnit: 3,
    catAxisLabelFontSize: 8.5, valAxisLabelFontSize: 10, catAxisLabelRotate: 90,
    showLegend: true, legendPos: 'b', legendFontSize: 10.5, legendColor: MUTED,
  }));
  stat(s, 8.9, 2.15, 3.85, `COP ${t.falta_M} M`, 'Falta por aflorar',
    `Sobre COP ${t.colocacion_M.toLocaleString('es-CO')} M de colocación aún inmadura. Rango intercuartil: ${t.p25_M} a ${t.p75_M} M.`, ORANGE);
  card(s, 8.9, 3.95, 3.85, 2.25);
  txt(s, 'Cómo leer el rango', { x: 9.14, y: 4.15, w: 3.4, h: 0.28, fontSize: 12.5,
    bold: true, color: INK, fontFace: BODY });
  txt(s, 'El rango no es un intervalo de confianza: es la dispersión real de los factores de desarrollo entre las cosechas maduras.\n\nLas cosechas de may-2026 en adelante no se proyectan: a los tres meses el factor mediano es 15x, con un rango de 8,7x a 35x.',
    { x: 9.14, y: 4.45, w: 3.4, h: 1.65, fontSize: 10.5, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });
  footer(s, `Se proyecta el pico de mora 30+ sobre la colocación, no la pérdida final: el archivo no trae castigos ni recuperación de garantía. Quedan fuera ${t.n_jovenes} cosechas de 2026 por COP ${t.jovenes_M.toLocaleString('es-CO')} M, demasiado jóvenes para proyectar.`);
  s.addNotes('Este es el insumo directo del cálculo de provisión. El rango importa tanto como el punto medio.');
}

// ========================================================== 9. MADURACIÓN DEL LIBRO
{
  const s = pres.addSlide();
  const m = D.maduracion_total;
  header(s, 'Lo que viene', `El ${m.pct_joven}% del libro sano tiene menos de seis meses: su riesgo todavía no se ha visto`,
    'Saldo al día por edad del crédito, y cuánto de cada tramo se espera que entre en mora cada mes según la curva de hazard.');
  s.addChart(pres.ChartType.bar, [{ name: 'Saldo al día',
    labels: D.maduracion.map(x => x.edad + ' m'), values: D.maduracion.map(x => x.saldo_M) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.15, w: 6.0, h: 3.6, barDir: 'col', chartColors: [BLUE],
      valAxisLabelFormatCode: '0', showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: MUTED, dataLabelFontSize: 9.5, dataLabelFormatCode: '0',
      catAxisLabelFontSize: 9.5, valAxisLabelFontSize: 9.5, barGapWidthPct: 40,
    }));
  txt(s, 'Saldo al día, en millones de COP', { x: MX, y: 1.95, w: 6, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });
  s.addChart(pres.ChartType.bar, [{ name: 'Entrada esperada',
    labels: D.maduracion.map(x => x.edad + ' m'), values: D.maduracion.map(x => x.entra_M) }],
    Object.assign({}, baseChart, {
      x: 6.85, y: 2.15, w: 5.9, h: 3.6, barDir: 'col', chartColors: [ORANGE],
      valAxisLabelFormatCode: '0', showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: MUTED, dataLabelFontSize: 9.5, dataLabelFormatCode: '0',
      catAxisLabelFontSize: 9.5, valAxisLabelFontSize: 9.5, barGapWidthPct: 40,
    }));
  txt(s, 'Entrada esperada a mora cada mes, en millones de COP', { x: 6.85, y: 1.95, w: 6, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });
  txt(s, `Sobre COP ${m.saldo_M.toLocaleString('es-CO')} millones al día, la curva de hazard implica unos COP ${Math.round(m.entra_M)} millones entrando en mora cada mes. Casi todo es mora de 1 a 30 días, de la que el 41% se cura sola: no es pérdida, es el caudal que la cobranza tiene que atender.`,
    { x: MX, y: 5.85, w: 12.1, h: 0.5, fontSize: 11.5, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 16 });
  footer(s, 'Entrada esperada = saldo al día del tramo × hazard del tramo. Es un ejercicio determinista sobre la cartera de hoy, sin nueva originación.');
}

// ========================================================== 10. CONCENTRACIÓN
{
  const s = pres.addSlide();
  header(s, 'Límites', 'No hay problema de concentración por deudor; sí lo hay en dónde está la mora',
    'Dos medidas distintas: cuánto saldo acumula el top de deudores, y cuántos créditos concentran la mora.');
  s.addChart(pres.ChartType.bar, [{ name: 'Saldo',
    labels: D.conc_deudor.map(c => 'Top ' + c.top + '%'), values: D.conc_deudor.map(c => c.pct) }],
    Object.assign({}, baseChart, {
      x: MX, y: 2.35, w: 5.6, h: 3.2, barDir: 'col', chartColors: [BLUE],
      valAxisMaxVal: 100, valAxisMajorUnit: 25, showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
      dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10,
      barGapWidthPct: 50,
    }));
  txt(s, '% del saldo que acumula el top de deudores', { x: MX, y: 2.12, w: 5.6, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });
  s.addChart(pres.ChartType.bar, [{ name: 'Mora',
    labels: D.conc_mora.map(c => c.pct_creditos + '%'), values: D.conc_mora.map(c => c.pct_mora) }],
    Object.assign({}, baseChart, {
      x: 6.85, y: 2.35, w: 5.9, h: 3.2, barDir: 'col', chartColors: [ORANGE],
      valAxisMaxVal: 100, valAxisMajorUnit: 25, showValue: true, dataLabelPosition: 'outEnd',
      dataLabelColor: INK, dataLabelFontSize: 10.5, dataLabelFontBold: true,
      dataLabelFormatCode: '0"%"', catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10,
      barGapWidthPct: 50,
    }));
  txt(s, '% de la mora que concentran los peores créditos', { x: 6.85, y: 2.12, w: 5.9, h: 0.24,
    fontSize: 10.5, bold: true, color: INK, fontFace: BODY });
  txt(s, 'El deudor más grande no llega al 1% del saldo y el top 10% apenas suma el 21%: para una cartera de consumo eso es sano, y no justifica un tope por deudor. La concentración está en el otro lado: el 10% peor de los créditos tiene el 74% de la mora. La cobranza no se reparte, se focaliza.',
    { x: MX, y: 5.85, w: 12.1, h: 0.7, fontSize: 12, color: INK, fontFace: BODY,
      valign: 'top', lineSpacing: 16 });
  footer(s, 'Cartera viva a ago-2026: 1.213 créditos de 1.195 deudores. Solo 18 deudores tienen más de un crédito, muy pocos para medir si el cliente recurrente es mejor riesgo.');
}

// ========================================================== 11. CORRECCIÓN AL INFORME A JUNTA
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, 'CORRECCIÓN AL INFORME PRESENTADO A JUNTA', { x: MX, y: 0.62, w: 11, h: 0.26,
    fontSize: 11, color: 'F0A17A', bold: true, charSpacing: 1.8, fontFace: BODY });
  txt(s, 'La diferencia entre puntos de venta no resiste la prueba estadística',
    { x: MX, y: 0.95, w: 11.6, h: 0.9, fontSize: 26, bold: true, color: WHITE,
      fontFace: HEAD, lineSpacing: 32 });
  const filas = [
    ['Lo que dijimos', 'El punto de venta pesa diez veces más que el plazo o la garantía, y cerrar la brecha de los dos peores vale COP 71 millones al año.'],
    ['Qué estaba mal', 'Esa comparación mezclaba créditos de edades distintas. Los puntos no originan al mismo ritmo, así que buena parte de la diferencia era edad, no riesgo.'],
    ['Lo que resiste', 'Controlando la edad, el punto de venta queda en p = 0,083 y ningún punto es significativo por separado. Al medir todo al mes 6, p = 0,91. La diferencia observada es compatible con azar.'],
    ['Qué no cambia', 'Que el plazo, el monto, la tasa y el codeudor tampoco discriminan. La conclusión de fondo se sostiene: con estos campos no hay política de originación posible.'],
  ];
  filas.forEach(([t, b], i) => {
    const y = 2.15 + i * 1.12;
    txt(s, t, { x: MX, y, w: 2.6, h: 0.5, fontSize: 12.5, bold: true, color: 'F3B08C',
      fontFace: BODY, valign: 'top', lineSpacing: 16 });
    txt(s, b, { x: MX + 2.8, y, w: 9.3, h: 0.95, fontSize: 12, color: 'CADCFC',
      fontFace: BODY, valign: 'top', lineSpacing: 16 });
  });
  txt(s, 'La acción sigue teniendo sentido —auditar dos puntos cuesta poco y la señal, aunque débil, apunta en una dirección— pero se somete como decisión bajo incertidumbre, no como hallazgo probado.',
    { x: MX, y: 6.55, w: 12.1, h: 0.55, fontSize: 11.5, color: '8FA8C4', italic: true,
      fontFace: BODY, valign: 'top', lineSpacing: 15 });
  s.addNotes('Decirlo antes de que lo encuentre alguien más. La conclusión de fondo del informe a junta no cambia; el respaldo de una de las tres decisiones sí.');
}

// ========================================================== 12. PARÁMETROS
{
  const s = pres.addSlide();
  header(s, 'Propuesta', 'Los parámetros que el análisis permite mover hoy',
    'Solo los que tienen un número detrás. Los demás quedan abiertos hasta que existan los datos.');
  const rows = [[{ text: 'Parámetro', options: { bold: true, color: MUTED } },
    { text: 'Hoy', options: { bold: true, color: MUTED } },
    { text: 'Propuesto', options: { bold: true, color: MUTED } },
    { text: 'En qué se sostiene', options: { bold: true, color: MUTED } }]];
  [['Foco de cobranza temprana', 'Sin tramo definido', '31 a 90 días de mora',
    'Del balde 31-60 se cura el 31%; del 61-90, el 12%; pasados los 90, el 3%'],
   ['Política de castigo', 'No definida en el archivo', 'A los 180 días',
    'Ninguno de los 1.036 créditos con más de 180 días volvió por debajo de 30'],
   ['Provisión adicional', 'La del modelo actual', '+ COP 156 M',
    'Pico proyectado de las cosechas inmaduras, rango intercuartil 247 a 428 M'],
   ['Tratamiento de la calificación C', 'Provisión de C', 'Provisionar como D',
    'El 53% del saldo en C cae a D en un solo mes'],
   ['Tope de exposición por deudor', 'No existe', 'No crearlo todavía',
    'El top 10% de deudores concentra 21% del saldo: no hay problema que resolver'],
   ['Cobranza por crédito', 'Repartida', 'Focalizada en 122 créditos',
    'El 10% peor concentra el 74% de la mora, COP 541 M'],
  ].forEach(r => rows.push([{ text: r[0], options: { bold: true } }, { text: r[1] },
    { text: r[2], options: { bold: true, color: BLUE } }, { text: r[3] }]));
  s.addTable(rows, { x: MX, y: 2.05, w: W - 2 * MX, colW: [2.75, 2.1, 2.35, 4.89],
    fontSize: 10.8, fontFace: BODY, color: INK, rowH: 0.6, valign: 'middle',
    border: { type: 'solid', color: 'E8ECF1', pt: 0.5 }, fill: { color: 'FFFFFF' } });
  footer(s, 'Ninguna de estas propuestas toca la política de originación: con los campos disponibles no hay evidencia para moverla en ningún sentido.');
}

// ========================================================== 13. LO QUE FALTA
{
  const s = pres.addSlide();
  header(s, 'Diseño de datos', 'Los cinco campos que convierten esto en una política de crédito',
    'Cada uno con la decisión que desbloquea y el costo de no tenerlo.');
  const g = [
    ['Cuota inicial y precio de factura', 'Fijar la cuota inicial mínima por segmento. Es el driver de riesgo número uno en motos y hoy no se puede medir.'],
    ['Score de buró en la originación', 'Fijar el punto de corte de aprobación y construir un scorecard propio.'],
    ['Solicitudes rechazadas', 'Sin ellas, cualquier modelo se entrena solo sobre aprobados y no dice nada sobre a quién más se podría prestar.'],
    ['Asesor comercial', 'Señalar responsabilidad dentro de un punto de venta, no solo el punto entero.'],
    ['Fecha de castigo y de recuperación de garantía', 'Medir pérdida real, no solo mora. Hoy se proyecta el pico de mora porque la pérdida no es observable.'],
  ];
  g.forEach(([t, b], i) => {
    const y = 2.02 + i * 0.88;
    s.addShape(pres.ShapeType.ellipse, { x: MX, y: y + 0.06, w: 0.32, h: 0.32,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
    txt(s, String(i + 1), { x: MX, y: y + 0.06, w: 0.32, h: 0.32, fontSize: 12.5, bold: true,
      color: WHITE, align: 'center', valign: 'middle', fontFace: BODY });
    txt(s, t, { x: MX + 0.48, y: y + 0.04, w: 4.1, h: 0.6, fontSize: 12.5, bold: true,
      color: INK, fontFace: BODY, valign: 'top', lineSpacing: 16 });
    txt(s, b, { x: MX + 4.75, y: y + 0.04, w: 7.35, h: 0.70, fontSize: 11.2, color: MUTED,
      fontFace: BODY, valign: 'top', lineSpacing: 15 });
  });
  txt(s, 'Los tres primeros no requieren desarrollo: ya existen en el proceso comercial y en la consulta de buró que se hace al aprobar. Lo que falta es que queden guardados en el reporte mensual.',
    { x: MX, y: 6.45, w: 12.1, h: 0.45, fontSize: 11.5, color: INK, italic: true,
      fontFace: BODY, valign: 'top', lineSpacing: 15 });
}

// ========================================================== 14. ANEXO
{
  const s = pres.addSlide();
  header(s, 'Anexo', 'Cómo se construyó cada análisis',
    'Todo reproducible desde el archivo original con analizar_cosechas.py y comite/preparar_comite.py. Tablas fuente en comite/ y datos/.');
  card(s, MX, 2.05, 6.0, 2.45);
  txt(s, 'Las decisiones metodológicas que importan', { x: MX + 0.28, y: 2.2, w: 5.5, h: 0.28,
    fontSize: 13, bold: true, color: INK, fontFace: BODY });
  txt(s, 'Se excluye siempre la transición de junio a julio de 2026: el castigo la distorsiona.\n\nSe excluyen los créditos vivos pasados de su plazo pactado: siguen ahí porque no pagaron, y el 93% está en mora. Incluirlos fabrica relaciones falsas entre plazo y riesgo.\n\nUn solo corte por crédito en la regresión, y ninguna tasa con menos de 30 observaciones.',
    { x: MX + 0.28, y: 2.55, w: 5.5, h: 2.1, fontSize: 10.2, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });

  card(s, MX + 6.3, 2.05, 6.0, 2.45);
  txt(s, 'Lo que estos números NO son', { x: MX + 6.58, y: 2.2, w: 5.5, h: 0.28,
    fontSize: 13, bold: true, color: INK, fontFace: BODY });
  txt(s, 'La proyección es del pico de mora 30+, no de la pérdida final: el archivo no trae castigos ni recuperación de garantía.\n\nEl rango es la dispersión histórica de los factores entre cosechas, no un intervalo de confianza.\n\nTodo lo medido sobre créditos vivos hereda el sesgo de la lámina 5, y no hay inferencia posible sobre los rechazados.',
    { x: MX + 6.58, y: 2.55, w: 5.5, h: 2.1, fontSize: 10.2, color: MUTED, fontFace: BODY,
      valign: 'top', lineSpacing: 14 });

  txt(s, 'Tamaños de muestra', { x: MX, y: 4.85, w: 6, h: 0.28, fontSize: 13, bold: true,
    color: INK, fontFace: BODY });
  [`Regresión logística: ${D.muestra_logit.n} créditos, ${D.muestra_logit.mora}% en mora 30+, cosechas 2024-2026 dentro de su plazo.`,
   'Hazard y cura: cinco transiciones mensuales entre feb y ago de 2026, entre 1.183 y 66 créditos expuestos por tramo de edad.',
   'Factores de desarrollo: 24 cosechas con 25 meses o más observados, entre ene-2022 y jun-2024.',
   'Migración de calificación: cinco transiciones, ponderadas por saldo de capital.',
  ].forEach((t, i) => txt(s, t, { x: MX, y: 5.22 + i * 0.38, w: 11.8, h: 0.36, fontSize: 10.2,
    color: MUTED, fontFace: BODY, valign: 'top', bullet: { code: '2013' }, lineSpacing: 13 }));

}

const OUT = __dirname + '/Cosechas_Districolmotos_Comite.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('listo:', OUT));
