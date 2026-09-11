# Informe de cosechas para junta directiva

`Cosechas_Districolmotos_Junta.pptx` — 21 láminas, corte de datos 31-ago-2026.

## Cómo se rehace

```
python3 ../analizar_cosechas.py Cosechas_2026.xlsx   # refresca las tablas del análisis
python3 preparar_datos.py                            # arma deck_data.json y heatmap.png
node build_deck.js                                   # genera el .pptx
```

Requiere `pandas`, `openpyxl`, `matplotlib` y `pptxgenjs` (`npm install pptxgenjs`).

## Estructura del informe

| # | Lámina | Para qué está |
|---|---|---|
| 1 | Portada | El título ya es la conclusión |
| 2 | Resumen ejecutivo | Tres mensajes y tres decisiones. Si solo se lee una, es esta |
| 3 | **La salvedad** | El castigo de julio, **antes** de los resultados |
| 4 | Qué es una cosecha | 30 segundos, saltable si la junta ya lo maneja |
| 5 | Curvas por año | El gráfico central del informe |
| 6 | Mora al mes 12 por cosecha | Dónde está el daño |
| 7–11 | **Cosecha por cosecha** | Doce cuadros por año, misma escala, misma referencia |
| 12 | Triángulo completo | Las 52 cosechas de un vistazo |
| 13 | Matriz de transición | El punto de no retorno a los 60 días |
| 14 | La cartera hoy | Capital vencido vs exposición real |
| 15 | Punto de venta | Comparación a edad constante |
| 16 | Concentración | 122 créditos, 74% de la mora |
| 17 | Colocación | Crecimiento con calidad |
| 18 | **Lo que no sabemos** | Los tres campos que faltan |
| 19 | Las tres decisiones | Con dueño, plazo y monto |
| 20–21 | Anexos | Metodología, controles y tabla por cosecha |

Cada lámina lleva **notas del presentador** con lo que hay que decir al pasarla.

## Decisiones que se someten a aprobación

| # | Qué | Quién | Cuándo | En juego |
|---|---|---|---|---|
| 1 | Auditoría de originación en Cartagena-Suzuki y Puerto Berrío AKT | Gerencia comercial | 31-oct | COP 71 M al año |
| 2 | Cuadrar el castigo de julio con contabilidad | Contabilidad y CFO | 15-oct | COP 549 M |
| 3 | Incorporar cuota inicial, score de buró y asesor al reporte mensual | Sistemas y crédito | 30-nov | Habilita fijar política |

## Kit de defensa — las preguntas incómodas

**"¿Por qué este número no coincide con el de contabilidad?"**
Lámina 14. El indicador de cosechas mide **capital vencido** (COP 304 M): solo las cuotas
ya vencidas y no pagadas. La **exposición** —el saldo completo de esos mismos créditos—
son COP 728 M. Las dos cifras son correctas y miden cosas distintas.

**"¿De dónde sale el número?"**
La suma del capital vencido de los créditos con más de 30 días de mora en ago-2026 da
COP 301.160.233, idéntico peso a peso al total de la hoja `Consolidado` del archivo. El
control lo imprime `analizar_cosechas.py` cada vez que corre.

**"¿La mora bajó en julio, entonces?"**
No. Lámina 3. Salieron 242 créditos y COP 549 M del reporte, y la fecha de desembolso más
antigua saltó de jun-2016 a oct-2021. Es un castigo o una depuración.

**"¿No será que 2024 se ve mejor solo porque es más nueva?"**
Lámina 4 y 5. Todas las comparaciones son a edad constante: el mes 12 contra el mes 12.
Y por eso las líneas de 2025 y 2026 se cortan antes: no hay dato que mostrar todavía.

**"¿Entonces 2024 ya ganó?"**
No, y conviene decirlo antes de que lo pregunten. 2022 y 2023 convergen cerca del 8% hacia
los meses 24–30, y 2024 apenas va en el mes 18. Lo que está probado es que se deteriora a
la mitad del ritmo, no que vaya a perder la mitad.

**"¿Por qué el punto de venta se compara solo a los 6 meses?"**
Porque a esa edad el capital de los créditos vigentes todavía cubre el 81% de lo colocado.
Más allá, los créditos buenos ya se pagaron y salieron del archivo, así que la tasa se
dispara por sesgo de supervivencia, no por riesgo.
