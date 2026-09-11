# Analítica de riesgo — comité de crédito y cartera

`Cosechas_Districolmotos_Comite.pptx` — 14 láminas, corte 31-ago-2026.
Complementa el informe a junta: la junta aprueba, el comité opera.

```
python3 ../analizar_cosechas.py Cosechas_2026.xlsx   # tablas base
python3 preparar_comite.py                           # datos_comite.json
node build_comite.js                                 # el .pptx
```

Requiere `pandas`, `numpy`, `scipy`, `statsmodels` y `pptxgenjs`.

## Los nueve análisis y lo que cada uno decide

| # | Análisis | Hallazgo | Decisión que habilita |
|---|---|---|---|
| 1 | Mapa de cobertura | De seis decisiones del comité, el archivo responde dos | Qué pedirle a sistemas |
| 2 | Qué discrimina el riesgo | **Solo la edad del crédito** (p<0,001) | Ninguna política de originación es defendible hoy |
| 3 | Curva de hazard | El riesgo de entrar en mora **sube** con la edad: 10,6% → 16,7% | Cobranza preventiva en toda la vida, no solo al inicio |
| 4 | Prepago y selección adversa | Los sanos salen 4 veces más rápido a los 18-23 meses | Explica el punto 3 y el sesgo de todo indicador sobre vivos |
| 5 | Curva de cura | 31-60 días: 31% se recupera. Más de 180 días: **cero de 1.036** | Foco 31-90 días; castigo a los 180 |
| 6 | Migración de calificación | La B es la bisagra; desde C el 53% cae a D en un mes | Provisionar C como D |
| 7 | Proyección de cosechas | Faltan por aflorar **COP 156 M** (rango 247-428 M de pico) | Provisión adicional |
| 8 | Maduración del libro | 62% del saldo sano tiene menos de 6 meses | Dimensionar cobranza para lo que viene |
| 9 | Concentración | Por deudor no hay problema; 122 créditos tienen el 74% de la mora | No crear tope por deudor; focalizar cobranza |

## Lo que este trabajo corrigió del informe a junta

El informe a junta afirmaba que **el punto de venta pesa diez veces más que el plazo o la
garantía**, y cuantificaba la brecha de los dos peores puntos en COP 71 millones al año.

Esa comparación mezclaba créditos de edades distintas. Los puntos no originan al mismo
ritmo, así que buena parte de la diferencia era edad, no riesgo. Controlando la edad:

- El punto de venta queda en **p = 0,083** y ningún punto es significativo por separado.
- Midiendo todo al mes 6 de vida, **p = 0,91**.

La conclusión de fondo se sostiene y hasta se refuerza —ninguna otra variable discrimina
nada— pero la acción se somete como **decisión bajo incertidumbre**, no como hallazgo
probado. El informe a junta y el tablero quedaron corregidos en ese punto.

## Las tres cosas que cambiaría mañana

1. **Foco de cobranza en 31-90 días.** Del balde 31-60 se recupera el 31%; pasados los 90,
   el 3%. Hoy el esfuerzo no está tramificado.
2. **Política de castigo a los 180 días.** Ninguno de los 1.036 créditos observados con más
   de 180 días volvió por debajo de 30. En cinco transiciones, cero.
3. **Provisionar la calificación C como D.** El 53% del saldo en C cae a D en un solo mes.

## Lo que más plata vale y hoy no se puede responder

Dónde poner el punto de corte de aprobación. Requiere score de buró en la originación,
cuota inicial efectiva y —esto es lo que casi nunca se guarda— **el registro de las
solicitudes rechazadas**. Sin ellas, cualquier modelo se entrena solo sobre aprobados y no
dice nada sobre a quién más se podría prestar.

## Salvedades que no hay que perder de vista

- La proyección es del **pico de mora 30+**, no de la pérdida final: el archivo no trae
  castigos ni recuperación de garantía.
- El rango 247-428 M es la **dispersión histórica de los factores** entre cosechas, no un
  intervalo de confianza.
- Todo lo medido sobre créditos vivos hereda el sesgo del análisis 4: los buenos se van antes.
- Se excluyen los 86 créditos vivos pasados de su plazo pactado. Siguen en el libro porque
  no pagaron y el 93% está en mora: incluirlos fabrica una relación entre plazo y riesgo
  que no existe.
- Solo 18 deudores tienen más de un crédito: no alcanza para medir si el cliente recurrente
  es mejor riesgo, aunque la diferencia observada apunte en esa dirección.
