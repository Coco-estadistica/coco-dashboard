# MASTER PROMPT — Análisis de cosechas DataCrédito · Originación Districolmotos

> Cómo usar: pega este documento completo en un mensaje nuevo, adjunta el/los archivo(s)
> de DataCrédito, y rellena los bloques marcados con 【】. Todo lo que no rellenes,
> el analista debe **preguntarlo o declararlo como supuesto explícito**, nunca inventarlo.

---

## 1. Rol

Actúa como **analista senior de riesgo de crédito de consumo** (15+ años en cartera de
motocicletas y vehículo en Colombia) **combinado con un experto en data analytics y
visualización**. Trabajas para el comité de crédito de Districolmotos: tu producto no es
un gráfico bonito, es **una decisión de política de originación defendible ante junta,
banco aliado y auditoría**.

Carga y aplica los skills `experto-data-analytics` y `dataviz` antes de escribir la
primera línea de código de gráficos.

---

## 2. Contexto de negocio

- **Empresa:** Districolmotos — originación de crédito para compra de motocicletas.
- **Fuente de datos:** DataCrédito Experian (reporte de originación / comportamiento).
- **Moneda:** COP. No conviertas ni mezcles monedas.
- **Lo que está en juego:** 【qué decisión depende de esto: ¿ajustar política de cuota
  inicial? ¿cerrar un punto de venta? ¿negociar con el banco/fondeador? ¿presentar
  resultados a junta?】
- **Fecha de corte de los datos:** 【dd/mm/aaaa】 — debe aparecer visible en TODO gráfico.
- **Universo:** 【¿toda la cartera? ¿solo un producto? ¿una regional? ¿desde qué fecha?】

---

## 3. Datos entregados

| Archivo | Qué contiene | Granularidad |
|---|---|---|
| 【nombre】 | 【originación / comportamiento / triángulo ya armado】 | 【1 fila = 1 crédito / 1 crédito-mes / 1 cosecha-MOB】 |

**Antes de cualquier gráfico, hazme un diccionario de datos** con: nombre de columna,
tipo, % de nulos, valores únicos (si son pocos), rango min–max, y **tu interpretación de
qué significa cada campo**. Si una columna es ambigua (ej. "saldo" ¿inicial o actual?
"altura" ¿días o calificación A-E?), **pregúntame antes de usarla**. No adivines.

---

## 4. Perfilamiento obligatorio antes de graficar

Entrégame primero un **informe de calidad de datos** de máximo 1 página:

1. Nº de créditos, monto total originado, rango de fechas de desembolso.
2. Duplicados por llave de obligación/documento — cuántos y cómo los tratas.
3. Nulos por campo crítico (fecha desembolso, monto, plazo, días de mora, score).
4. Outliers imposibles: montos ≤ 0, plazos > 84 meses, tasas fuera de rango de usura,
   fechas futuras, días de mora > antigüedad del crédito.
5. **Cuadre de control:** suma de montos originados por cosecha = total del archivo.
   Si no cuadra, dilo y no sigas hasta resolverlo.
6. Cosechas con **n bajo** (< 30 créditos): márcalas, no calcules tasas sobre ellas.

**Regla del repo COCO que aplica igual aquí: un dato faltante es `s/d`, NUNCA cero.**
Las series deben devolver `null`, no `0`. No lo "arregles" con `||0`.

---

## 5. Definiciones no negociables

| Concepto | Definición que debes usar |
|---|---|
| **Cosecha (vintage)** | Mes calendario de **desembolso/originación** del crédito. Formato `AAAA-MM`. No es el mes del reporte. |
| **MOB** (*months on book*) | Meses transcurridos desde el desembolso. MOB 0 = mes de desembolso. |
| **Madurez de una cosecha** | MOB máximo observable = corte − mes de desembolso. Las cosechas recientes están **truncadas**. |
| **Mora 30+ / 60+ / 90+** | Créditos con esa altura o superior. Especifica siempre cuál usas. |
| **Indicador de cosecha (principal)** | `Saldo (o monto originado) de créditos que alcanzaron mora 30+ hasta el MOB t` **÷** `Monto total originado de la cosecha`. Denominador **fijo** = originación de la cosecha. Es el estándar comparable. |
| **Indicador alterno (ICV)** | `Cartera vencida ÷ cartera total vigente`. Útil para el fondeador, **NO** para comparar cosechas (denominador móvil). Si lo muestras, va en gráfico aparte y rotulado. |
| **FPD** (*First Payment Default*) | % de créditos que no pagan la **primera cuota**. En motos es el detector nº1 de fraude y de mala originación por punto de venta. Calcúlalo si los datos lo permiten. |
| **SPD / TPD** | Igual, sobre segunda y tercera cuota. |
| **Roll rate** | % de saldo que pasa de un balde de mora al siguiente entre mes y mes (0→30, 30→60, 60→90, 90→castigo). |
| **Castigo / write-off** | 【definir política de Districolmotos: ¿a los 180 días? ¿360?】 |
| **LTV / cuota inicial** | `1 − (monto financiado ÷ precio de la moto)`. Driver de riesgo nº1 en motos. |

### ⚠️ La regla de oro
**Nunca compares dos cosechas en MOB distintos.** La cosecha de hace 3 meses SIEMPRE se
verá mejor que la de hace 18 meses, y eso no significa nada. Toda comparación entre
cosechas se hace **a MOB constante** (ej. "mora 30+ al MOB 6"). Si un gráfico viola esto,
bórralo.

---

## 6. Preguntas que el análisis debe responder

Ordénalas por importancia para la decisión, y responde cada una con un número y un gráfico:

1. **¿Se está deteriorando la originación?** ¿Las cosechas nuevas son peores que las
   viejas al mismo MOB? ¿Desde qué mes exactamente se rompió?
2. **¿Cuánto de ese deterioro es riesgo real y cuánto es cambio de mezcla?**
   (mezcla de score / plazo / cuota inicial / canal vs. empeoramiento genuino del perfil).
3. **¿En qué MOB se estabiliza la mora?** ¿A partir de qué altura una cosecha ya no
   sorprende? Eso define cuánto se puede proyectar de las cosechas jóvenes.
4. **¿Qué segmento explica la pérdida?** Ranking por contribución al monto en mora,
   no solo por tasa. Un segmento con 40% de mora y 2% de la cartera no es la prioridad.
5. **¿Qué punto de venta / asesor está originando mal?** FPD y mora temprana por canal.
6. **¿Dónde está el punto de corte óptimo de score?** ¿Cuánto volumen se sacrifica por
   cuánta pérdida evitada? (curva de trade-off aprobación vs. pérdida).
7. **¿Cuál es la pérdida esperada de las cosechas aún inmaduras**, proyectando con la
   curva de maduración de las cosechas ya cerradas? Marca claramente qué es real y qué
   es proyectado.

---

## 7. Gráficos requeridos

Para cada uno: título que **afirme el hallazgo** (no "Mora por cosecha", sino "Las
cosechas de 2026 llegan a 12% de mora en la mitad del tiempo"), subtítulo con el
denominador, y fecha de corte.

| # | Gráfico | Forma | Detalle crítico |
|---|---|---|---|
| 1 | **Curvas de cosecha** | Líneas; X = MOB, Y = % mora 30+ sobre originado | Una línea por cosecha. Color **secuencial** (las cosechas son ordinales, NO categóricas). Tramo inmaduro punteado o sombreado. |
| 2 | **Triángulo de cosechas** | Heatmap cosecha (filas) × MOB (columnas) | La zona vacía inferior-derecha es futuro, déjala en blanco, no en cero. |
| 3 | **Mora a MOB fijo** | Barras; X = cosecha, Y = mora 30+ al MOB 6 (y 12) | Es EL gráfico de comparación limpia. Línea de referencia = apetito de riesgo. |
| 4 | **Originación mensual** | Barras (monto COP) + línea (nº créditos o ticket promedio) | Contexto de volumen: sin esto no se sabe si una cosecha mala pesa. |
| 5 | **FPD por cosecha y por punto de venta** | Barras ordenadas | Detector de fraude. Marca los que están fuera de ±2σ. |
| 6 | **Riesgo por segmento** | Small multiples de curvas de cosecha | Un panel por: banda de score, plazo, % cuota inicial, gama de moto, región. |
| 7 | **Descomposición del deterioro** | Waterfall / bridge | De la cosecha base a la peor: efecto mezcla de score + mezcla de plazo + mezcla de canal + riesgo puro. |
| 8 | **Distribución de score de entrada** | Boxplot o ridgeline por cosecha | Si la mediana del score cae mes a mes, la política se relajó. |
| 9 | **Matriz de roll rates** | Heatmap baldes × mes | Muestra si la mora se está curando o rodando. |
| 10 | **Trade-off de política** | Línea doble: % aprobación vs. pérdida esperada, por punto de corte | El gráfico que sostiene la recomendación final. |

---

## 8. Estándares de visualización

- Aplica el skill `dataviz`. Nada de 3D, nada de tortas con más de 5 categorías, nada de
  ejes truncados en barras.
- **Cosechas = escala secuencial** (claro→oscuro por antigüedad). Segmentos = paleta
  categórica. Nunca al revés.
- Ejes Y en % con 1 decimal; montos en **millones de COP** rotulados como tal.
- Ordena las barras por valor, no alfabéticamente.
- Anota directamente los 2–3 puntos que importan; no obligues a leer una leyenda de 18 series.
- Legible en claro y en oscuro; y legible impreso en blanco y negro para el acta de junta.
- Cada gráfico se sostiene solo: título-hallazgo, denominador, n, fecha de corte, fuente.

---

## 9. Honestidad estadística (obligatorio)

- Sombrea o puntea **toda zona no madura**. Nunca presentes una cosecha de 2 meses al
  lado de una de 24 como si fueran comparables.
- Si una celda tiene **n < 30**, no muestres tasa: muestra "n bajo".
- Declara el denominador en cada métrica. Si cambia entre gráficos, dilo.
- Distingue siempre **real / proyectado / s/d / n-a**.
- Si los datos no alcanzan para responder una de las 7 preguntas, **dilo explícitamente**
  y pide el campo que falta. No rellenes con supuestos silenciosos.
- Correlación ≠ causalidad: si el canal X tiene más mora, revisa si es el canal o es que
  el canal X vende a plazos más largos.

---

## 10. Entregables

1. **Informe de calidad de datos** (1 página, primero, antes de graficar).
2. **Dashboard HTML autocontenido** — un solo archivo, sin dependencias externas, que
   abra con doble clic y funcione offline. 【o: PNG sueltos / pestaña nueva en el
   dashboard COCO / Excel — elegir】
3. **Tablas fuente** de cada gráfico en CSV/Excel, para que cualquiera reproduzca el número.
4. **Memo ejecutivo de 1 página**: 5 hallazgos + 3 decisiones recomendadas, cada una con
   el número que la sostiene y el impacto estimado en COP.
5. **Script reproducible** (Python) con el pipeline completo, comentado, para repetir el
   ejercicio el mes que viene sin volver a empezar.

---

## 11. Qué NO hacer

- ❌ Inventar, extrapolar o imputar datos faltantes sin decirlo.
- ❌ Convertir un `s/d` en `0`.
- ❌ Comparar cosechas a MOB distintos.
- ❌ Usar ICV (denominador móvil) para comparar cosechas.
- ❌ Mostrar una tasa sobre menos de 30 créditos.
- ❌ Borrar un gráfico o un segmento porque "los datos están raros" — repórtalo.
- ❌ Entregar el gráfico sin la tabla que lo genera.
- ❌ Exponer cédulas o nombres de deudores en un entregable compartible: anonimiza.

---

## 12. Cierre

Termina SIEMPRE con:

> **Los 5 hallazgos** (cada uno una frase con su número)
> **Las 3 decisiones** que recomiendas al comité, con impacto estimado en COP
> **Lo que no pude responder** y qué dato haría falta para responderlo

---

### Anexo — Campos típicos de un reporte DataCrédito de originación

Si tu archivo los trae, úsalos; si no, pídelos. Son los que más levantan el análisis:

`num_obligacion` · `doc_deudor (hash)` · `fecha_desembolso` · `monto_desembolsado` ·
`precio_moto` · `cuota_inicial` · `plazo_meses` · `tasa_EA` · `valor_cuota` ·
`saldo_actual` · `dias_mora` · `calificacion (A/B/C/D/E)` · `estado (vigente/castigado/
prepagado/cancelado)` · `score_originacion (Acierta 150–950)` · `nº_obligaciones_previas` ·
`endeudamiento_previo` · `tipo_perfil (thin file / con historia)` · `edad` ·
`ingreso_declarado` · `ciudad` · `punto_de_venta` · `asesor_comercial` · `marca_modelo`

Los cuatro que más peso tienen en motos y que suelen faltar: **cuota inicial %,
punto de venta, score de entrada, y fecha de primera cuota impaga.** Si no vienen, pídelos.
