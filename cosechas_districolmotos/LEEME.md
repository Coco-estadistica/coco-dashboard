# Cosechas de originación — Districolmotos

Análisis de cosechas (*vintage analysis*) de la cartera de crédito de motocicletas.
**Corte de datos: 31 de agosto de 2026.** Fuente: `Cosechas_2026.xlsx`.

| Archivo | Qué es |
|---|---|
| `Cosechas_Districolmotos.html` | El tablero. Doble clic, abre en el navegador, funciona sin internet |
| `analizar_cosechas.py` | Rehace todo el análisis desde el Excel y actualiza el tablero |
| `datos/` | Las tablas que alimentan cada gráfico, en CSV |

Para repetir el ejercicio el mes que viene:

```
python3 analizar_cosechas.py ruta/al/Cosechas_2026.xlsx
```

Imprime los controles, reescribe los CSV y reinyecta los datos en el HTML.
Requiere `pandas` y `openpyxl`.

---

## Memo ejecutivo

### Lo que pasó

**2023 fue el año que se deterioró más rápido.** Al mes 12 de vida, las cosechas de 2023
llegan a 4,4% de mora 30+ sobre lo colocado, contra 2,5% de las de 2022: un 80% peor
medido en el mismo punto de la vida del crédito. El daño no está repartido: tres cosechas
—mayo (8,5%), abril (7,4%) y septiembre (6,2%) de 2023— colocaron COP 889 millones entre
las tres y concentran el grueso.

**2024 y 2025 corrigieron, y la corrección se sostiene.** Al mes 18, las cosechas de 2024
van en 2,9% contra 4,7% de 2022 y 5,6% de 2023: prácticamente la mitad. Al mes 6, las de
2025 (1,1%) van incluso por debajo de las de 2024 (1,2%). Dos años consecutivos por debajo
no es ruido; es una política de originación que cambió y pegó.

**Y se creció al mismo tiempo.** La colocación pasó de COP 197 millones/mes en 2022 a
COP 450 millones/mes en 2026 — 2,3 veces — con mejor mora en todos los meses comparables.

### Lo que hay que mirar con cuidado

**El dato de julio y agosto de 2026 no se puede comparar con el de junio.** Entre esos dos
meses la mora reportada cayó COP 549 millones (−64%) y 22 cosechas bajaron su indicador de
golpe; doce quedaron en 0,00% exacto. Desaparecieron 242 créditos del reporte y la fecha de
desembolso más antigua saltó de jun-2016 a oct-2021. Eso es un castigo o una depuración, no
una recuperación de cartera. **Todas las curvas del tablero se cortan en junio de 2026.**

**La ventaja de 2024 es de velocidad, todavía no de resultado final.** Las cosechas de 2022
y 2023 convergen cerca del 8% hacia los meses 24–30, y 2024 apenas va en el mes 18. Lo
defendible hoy es "se deteriora a la mitad del ritmo", no "va a perder la mitad". Hay que
volver a medirla cuando cumpla 24 meses.

### Las tres decisiones

**1. Intervenir Puerto Berrío y Puerto Boyacá antes de seguir creciendo.**
En las cosechas desde ene-2025, esos dos puntos pusieron el 20% del capital originado y el
**32% de toda la mora**: COP 259 millones sobre COP 1.168 millones colocados (22,3% y 22,1%).
El resto de los puntos va en 11,7%. Si esos dos rindieran como el promedio de los demás,
habría **COP 123 millones menos en mora**. Y no es la plaza: en la misma Puerto Berrío,
la línea Suzuki va en 2,2% — diez veces mejor que la línea que lleva el nombre de la ciudad.
La diferencia está en la operación de cada punto, y ahí es donde hay que ir a mirar.

**2. Confirmar con contabilidad qué se castigó en julio, contra qué provisión y con qué
efecto en el P&G — antes de reportar cualquier mejora de la mora.** Son COP 549 millones.
Mientras eso no esté cuadrado, el indicador global de cierre de 2026 no es presentable ni a
junta ni a un banco.

**3. Pedir al sistema tres campos que hoy no existen en el archivo.** Sin ellos no se puede
responder la pregunta que de verdad importa —dónde poner el punto de corte de aprobación—:
- **cuota inicial / valor real de la moto** (`VALORGARAN` viene casi igual al capital
  prestado, así que no sirve para calcular LTV, que es el driver de riesgo número uno en motos);
- **score de buró en la originación** (no viene ninguno);
- **asesor comercial** (`NOMBREASES` viene vacía en los siete cortes).

Con esos tres campos, el mismo tablero permite fijar política. Sin ellos solo permite
diagnosticar.

---

## Metodología

**Cosecha** = mes de desembolso. **MOB** = meses transcurridos desde el desembolso
(MOB 0 = mes de desembolso).

**Indicador** = **capital vencido** (columna `SALDO EN MORA`) de los créditos con más de 30
días de mora *en esa fecha*, dividido por la **colocación original** de la cosecha. El
denominador es fijo: es lo que permite comparar cosechas entre sí.

Ojo con la lectura: el numerador es la porción de cuotas ya vencidas, no el saldo entero del
crédito. A los 31–60 días de mora es en promedio el 20% del saldo; pasados los 360 días, casi
el 100%. En ago-2026 el capital vencido son COP 301 millones, pero el **saldo total de esos
mismos créditos son COP 728 millones**: 2,4 veces más. El indicador mide lo vencido, no la
exposición.

Es un **saldo vigente, no un acumulado**: baja cuando los créditos se curan, se pagan o se
castigan, así que las curvas hacen pico y después ceden. Para medir pérdida definitiva por
cosecha haría falta el dato de castigos, que este archivo no trae.

Las curvas por año **ponderan cada cosecha por su monto colocado** (no promedian
porcentajes) y se dibujan con **composición constante**: cada línea llega solo hasta el mes
en que todas las cosechas de ese año siguen observadas. Si no, la curva daría un salto
artificial cuando las cosechas malas dejan de tener dato y quedan solo las buenas.

### Controles que corre el script

- **Numerador, cuadre exacto.** La suma del capital vencido de los créditos con
  `DIASMORA > 30` en el corte de ago-2026 da COP 301.160.233, idéntico peso a peso al total
  de la hoja `Consolidado`. Eso confirma qué mide el archivo.
- **Denominador constante.** Las 56 cosechas mantienen el mismo monto de colocación en los
  13 bloques de observación.
- **Ruptura de julio.** El script reporta cuánta mora salió y cómo se movió la fecha de
  desembolso más antigua.

### Salvedades

- **`Triángulo` vs `Consolidado`:** cuadran en 622 de 638 celdas. Las 16 que no cuadran
  están todas en la columna *julio 2025* de la hoja `Cosechas`, con valores redondeados a
  tres decimales (0,040 / 0,079 / 0,104…): esa columna se digitó a mano. El script usa el
  `Consolidado` como fuente y el triángulo solo para rellenar huecos.
- **La hoja `ENERO 26` está vacía** y el `Consolidado` no tiene bloque de enero de 2026.
  Ese mes es s/d, no cero.
- **Las hojas de febrero y abril traen otro formato** (99 y 100 columnas contra 105 del
  resto), con `FECHADESEM` en vez de `F_INICIOFI` y `SALDO EN MORA` con un espacio al final.
  El script las homologa.
- **Los gráficos de segmento son referenciales, no de cosecha.** Se calculan sobre los
  créditos vigentes a ago-2026 de cosechas desde ene-2025, usando el capital inicial como
  denominador. Los créditos ya cancelados no aparecen, así que la tasa está algo
  sobreestimada: sirve para **ordenar segmentos entre sí**, no como nivel absoluto.
- **La base crédito a crédito está anonimizada.** `CEDULASOCI` se reemplazó por un hash y la
  columna de nombre no se exporta.
- **Este archivo no es un reporte de DataCrédito.** Es el extracto de cartera del core
  (campos de reporte a Superintendencia: `CATEGORIAA`, `RPROVISION`, `CALHOMOLPE`). No trae
  score de buró ni endeudamiento externo del deudor.
