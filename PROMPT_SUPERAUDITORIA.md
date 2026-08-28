# Master prompt — Superauditoría del Cockpit COCO

> Cópialo completo en una sesión nueva. Está escrito para que quien lo reciba no
> necesite nada de la conversación anterior.

---

Vas a auditar un tablero financiero. Actúa simultáneamente con **tres sombreros**, y
declara con cuál encuentras cada cosa:

- **Ingeniero de sistemas.** El código hace lo que dice. Filtros que no filtran, funciones
  con dos caminos de lectura, estado que se pierde, controles que mienten sobre lo que
  hacen, errores en consola, cosas que se rompen en combinaciones de filtros.
- **Contador.** Las cifras cuadran contra su fuente. El estado de resultados suma, la
  consolidación es la suma de sus partes, las conversiones de moneda usan la tasa del
  mes correcto, los acumulados acumulan, los signos son consistentes, y nada se cuenta
  dos veces.
- **Analista financiero.** Las cifras son *comparables* y significan lo que su rótulo
  dice. Series que mezclan bases distintas, indicadores con dos definiciones, tasas de
  crecimiento construidas sobre universos distintos, y KPIs que se leen mal aunque estén
  bien calculados.

---

## 1. Qué se audita

| | |
|---|---|
| Tablero | `Dashboard_COCO.html` — un solo archivo, ~3.765 líneas, HTML+CSS+JS+SVG, SheetJS embebido, sin dependencias de red |
| Versión | `JULIO-2026-R4` |
| Base | `migracion/BD_MAESTRA_COCO.xlsx` — 14 hojas; la tabla de hechos es `BD_Indicadores` |
| Archivo de auditoría | `migracion/BD_TRAZABILIDAD_COCO.xlsx` — el tablero **no** lo abre |
| Repositorio | `Moscorrofio/coco-dashboard` (privado) |

**Estado de la base al momento de escribir esto** — si no coincide, la base cambió y hay
que revalidar todo lo de abajo:

```
BD_Indicadores   1.877 filas · 37 periodos (2023-12 .. 2026-12) · 63 códigos distintos
segmentos        (vacío) 1506 · "Consolidación USD" 276 · "Acumulado H1 2026" 38
                 "Original enero" 12 · "Revisado mayo" 12
```

**Cómo arrancarlo: siempre por HTTP, nunca abriendo el archivo con doble clic.**

```bash
python -m http.server 8740 --bind 127.0.0.1
```

y entrar a `http://localhost:8740/Dashboard_COCO.html`. Abrirlo como `file://` da
comportamientos distintos. Si ves cifras raras, verifica primero que no estés leyendo la
caché vieja de `localStorage` (la barra de estado dice de dónde salió la base).

> **Antes de auditar nada, confirma que estás sobre la copia viva.** Han llegado a
> convivir tres archivos `Dashboard_COCO.html` en carpetas distintas, uno de ellos diez
> días desactualizado y visualmente idéntico. La copia buena es la del repositorio.
> Compruébalo así: en **03 Finanzas & P&L** el gráfico de crecimiento debe titularse
> «Crecimiento 2025 vs 2026 · Colombia, y el aporte internacional aparte» y llevar
> **tres** series. Si dice «Facturación 2025 vs 2026» con dos series, estás en una copia
> vieja y todo lo que audites será inútil: para y busca la correcta.

## 2. Método — esto no es negociable

**Recalcula cada cifra por tu cuenta leyendo `BD_Indicadores`, y compárala contra lo que
pinta la pantalla. No valides una cifra del tablero con otra cifra del tablero.** El modo
de fallo más común de este sistema es la coherencia interna aparente: dos vistas que se
alimentan de la misma función equivocada y por eso concuerdan entre sí.

Recorre la interfaz de verdad: hunde cada botón, cada pestaña, cada selector, y las
**combinaciones** entre ellos. No te limites al estado por defecto.

## 3. Alcance obligatorio

**7 pestañas:** 01 Resumen · 02 Burn & Runway · 03 Finanzas & P&L · 04 Ingresos & MRR ·
05 Cartera & Liquidez · 06 Gastos · 07 Catálogo.

**6 filtros globales:** Rango (12m / Últ. mes / 6m / Total) · Desde · Hasta · Escenario
(Todos / Presupuesto / Real) · País (Consolidado / Colombia / EE.UU. / Perú / Costa Rica)
· Moneda (USD / COP) · Base TRM (Promedio mensual / Fin de mes).

**Combinaciones que hay que probar sí o sí**, porque es donde han salido los errores:

1. Cada pestaña × cada país. Especialmente los cortes que solo existen para Colombia.
2. Escenario **Presupuesto** en todas las pestañas. Nada puede mostrar cifras Reales bajo
   ese rótulo.
3. USD ↔ COP en todas. Convertir y volver debe dar lo mismo; los acumulados de varios
   meses deben sumarse mes a mes, no reconvertirse con una tasa única.
4. Rango **Total** contra rangos cortos. Que ampliar el rango no haga desaparecer datos.
5. El mes de corte movido hacia adelante (agosto–diciembre 2026, que solo tienen
   presupuesto) y hacia atrás (2024).
6. El modal «Ingresar datos», sus dos solapas, y la validación con campos vacíos.

**No ingreses credenciales, tokens ni URLs de escritura. No modifiques la base. No pulses
nada que escriba en la nube.** Si un flujo no se puede probar sin eso, dilo y sigue.

## 4. Convenciones del proyecto — no las reportes como errores

- **`s/d`** = el dato no llegó. **`n/a`** = no aplica a esa entidad. **`0`** = un cero
  reportado de verdad. Un hueco marcado `s/d` es correcto; lo incorrecto sería rellenarlo.
- **La TRM tiene dos series.** `trm_cop_usd` es derivada (ingresos COP ÷ ingresos USD de
  la consolidación) y `trm_promedio_mercado` es de mercado. Julio 2026 = **3.268,95**;
  junio 2026 = **3.505,27**. Una cifra de julio convertida con 3.505 es un error.
- **Perú se convierte en dos pasos**: PEN → USD con su propio factor, y USD → COP con la
  TRM del mes.
- Cada tarjeta declara su origen y su fecha. Un dato de un mes anterior se muestra **con
  sello de fecha**; con más de un mes de rezago pasa a `s/d` — es política, no error.

## 5. Trampas conocidas de la base — verifica si el tablero cae en ellas

Estas ya están identificadas. **No las reportes como hallazgos nuevos; repórtalas solo si
encuentras que alguna vista cae en ellas.**

1. **El acumulado del semestre vive dentro del mes.** Las 38 filas con segmento
   `Acumulado H1 2026` están en el periodo `2026-06`, junto a las filas mensuales. Una
   suma que no filtre por segmento devuelve ~7 veces el ingreso real de junio.
2. **Hay tres filas de presupuesto por mes**: segmento vacío, `Original enero` y
   `Revisado mayo`. Sumar por escenario sin filtrar segmento **triplica** el presupuesto.
   Además, en enero–abril el `Revisado mayo` es copia exacta del ejecutado, así que
   cualquier cumplimiento contra esa versión da 100% por construcción.
3. **El bloque consolidado empieza en enero 2026.** Los `pyg_*` de los cuatro países no
   existen antes. Cualquier serie que cruce ese límite cambia de universo.
4. **La fila `Consolidado` no se deriva: está guardada.** Si un país se corrige y no se
   recompone, el consolidado deja de cuadrar con sus propias columnas y nada avisa.
5. **`costo_proveedor` de julio 2026 está incompleto** a propósito: llegaron Infobip y
   B2Chat, faltan AWS y Otros. El gráfico lo declara. No es un error del tablero.
6. **A Colombia le faltan 70 reglas de homologación** de cuentas PUC, así que entra con
   totales donde las filiales entran con detalle. Está documentado en la pestaña Gastos.
7. **`indice_morosidad` (77,3%) lo entrega cartera con su propia definición** y no
   coincide con vencida ÷ total (41,6%). Es una decisión tomada, no un descuadre. Lo que
   sí es válido reportar es si la pantalla no deja clara la diferencia.
8. **Costa Rica está confirmada por el negocio.** Sus ingresos concentrados en julio con
   costos cero son correctos. No la reportes.

## 6. Cifras de control — si alguna no cuadra, es un hallazgo

Corte **2026-07**, escenario **Real**:

```
Colombia (COP)
  Ingresos operacionales netos        679.890.552
  EBITDA                               29.342.360   (margen 4,3%)
  Utilidad neta                        24.707.051
  Egresos totales                     655.901.810
  Gastos operativos (sin financieros) 650.548.189   (diferencia = 5.353.621 financieros)
  MRR recurrente                      703.109.438
  Cartera total                     1.535.916.158   (= 469.850 USD × 3.268,95)

Consolidado (USD)
  Ingresos                                350.899
  EBITDA                                  133.162   (margen 37,9%)
  Utilidad neta                           131.563
  Egresos totales                         219.567
  Gastos operativos                       217.738

Puente de MRR jun -> jul (COP), debe cerrar exacto:
  759.860.184 + 9.249.060 + 14.828.122 − 43.146.258 − 37.681.670 = 703.109.438

Crecimiento Colombia ene-jul, USD:
  2025  1.035.634  ->  2026  1.431.986   = +38,3% orgánico
  Filiales 2026: 181.561                 = +55,8% total
```

## 7. Lo que se corrigió el 28-ago-2026 — verifica que siga bien, no lo re-reportes

Cada uno trae cómo comprobarlo:

| Qué se corrigió | Cómo verificar |
|---|---|
| El P&G ignoraba el selector de Escenario (`factVal` no miraba `f.escenario`) | Con Escenario = Presupuesto, el Resumen consolidado no debe mostrar **ninguna** cifra Real. Antes mostraba 11. |
| El formulario ofrecía Guatemala y Panamá, y no ofrecía EE.UU. ni Perú | La lista debe salir de la base: Colombia · EE.UU. · Perú · Costa Rica · Otros · Regional |
| «Egresos» y «Gastos operativos» mostraban el mismo número | Deben diferir exactamente en los gastos financieros |
| Datos viejos bajo encabezados nuevos | Con más de 1 mes de rezago la tarjeta pasa a `s/d` y dice cuánto lleva sin actualizarse |
| El crecimiento mezclaba Colombia-2025 con Consolidado-2026 (+119%) | Ahora son tres series y el Δ es Colombia contra Colombia (jul: +30%) |
| La serie de MRR y la cartera de julio | El puente debe cerrar exacto; la cartera debe reconvertir a 469.850 USD |

**Busca activamente regresiones en estos seis**, que es donde más fácil se rompe algo.

## 8. Títulos, rótulos y textos

Es la capa que más barato se rompe y la que más caro cuesta cuando alguien externo lee el
tablero. Revisa **cada texto visible**, no solo los que rodean una cifra sospechosa.

1. **El título, ¿dice lo que el gráfico muestra?** Un título que promete el consolidado
   sobre un gráfico que solo trae Colombia es un error de la misma familia que una cifra
   mal sumada, y se detecta igual: comparando el rótulo con lo que hay debajo.
2. **¿Algún rótulo promete algo que vive en otra pestaña?** Si una pestaña se llama
   «Finanzas & P&L», el estado de resultados tiene que estar ahí.
3. **Unidades.** Que ningún porcentaje lleve símbolo de moneda ni al revés. Que un
   puntaje que no es porcentaje —el NPS va de −100 a +100— no se muestre con `%`. Que los
   días digan días.
4. **Vocabulario consistente.** El mismo concepto debe llamarse igual en las siete
   pestañas y en el Catálogo. Lista los sinónimos que encuentres para una misma cosa.
5. **Las notas explicativas, ¿siguen siendo ciertas?** Varias describen el estado de los
   datos: «el auxiliar del ERP arranca en ene-2026», «faltan AWS y Otros». Si la base
   cambió y la nota no, la nota miente. Verifica cada una contra los datos.
6. **Signos.** Que un mismo concepto no salga positivo en una tabla y entre paréntesis en
   la cascada de al lado.
7. **Ortografía, tildes y formato numérico.** Todo el tablero está en español: miles con
   punto, decimales con coma. Reporta cualquier mezcla.
8. **Restos de versiones anteriores**: rangos de fechas fijos, nombres de mes escritos a
   mano, referencias a botones que ya no existen.

## 9. Navegabilidad

Recorre el tablero **como si fuera la primera vez**, y después **solo con el teclado**.

1. **Estado y URL.** ¿La dirección refleja la vista? ¿Recargar conserva pestaña, país,
   rango y moneda? ¿El botón Atrás vuelve a la pestaña anterior o saca del tablero?
2. **Teclado.** Cuenta cuántos elementos hay que pasar con Tab hasta llegar a la primera
   pestaña. ¿Las flechas ← → se mueven entre pestañas? ¿Se ve dónde está el foco en todos
   los controles o solo en algunos?
3. **Lectores de pantalla.** ¿Las pestañas se anuncian como pestañas y se sabe cuál está
   activa, o esa información vive solo en el color?
4. **Controles que no aplican.** Construye la matriz completa de **pestaña × filtro**:
   cuáles mueven de verdad los datos y cuáles no hacen nada. Los que no aplican, ¿están
   apagados y dicen por qué, o siguen encendidos invitando a moverlos?
5. **Orientación.** ¿Se sabe siempre qué se está mirando —país, rango, moneda— sin volver
   a la barra de arriba? ¿El scroll vuelve arriba al cambiar de pestaña? En pestañas de
   varios miles de píxeles, ¿hay forma de volver al inicio?
6. **La barra superior pegada.** Con la página bajada, comprueba con una prueba de
   impacto que un clic en el centro de cada pestaña le llega **a la pestaña** y no a la
   barra de arriba.
7. **El modal de carga.** Las dos solapas, la validación con campos vacíos, el idioma de
   los controles, y si los avisos se acumulan o se pisan.
8. **Arquitectura de la información.** ¿Las siete pestañas son las siete correctas? ¿La
   numeración 01–07 codifica un recorrido real o es decorativa? ¿Hay contenido duplicado
   entre pestañas sin que ninguna indique que la otra existe?

## 10. Lógica

Aquí no se auditan cifras sino **razonamientos**. Es la capa donde un error no se ve
justamente porque todo cuadra.

1. **Cada indicador derivado: ¿la fórmula es defendible?** Búscalas en el código y
   contrástalas con su ficha del Catálogo. Una fórmula que no coincide con su definición
   publicada es un hallazgo aunque el número esté bien calculado.
2. **Dos caminos para el mismo dato.** Busca activamente si hay más de una función que
   lea la base con reglas distintas. Es el patrón que ya produjo un crítico aquí: las
   tarjetas filtraban por escenario y la tabla del P&G no, porque leían por caminos
   distintos. Si encuentras dos, revisa **todos** los puntos de llamada de cada una.
3. **Precedencia entre fuentes.** Cuando un dato existe en dos sitios —P&G oficial y
   serie de gestión, cargado y derivado—, ¿la regla de cuál manda está escrita en un solo
   lugar y se aplica igual en todas las vistas?
4. **Semáforos y umbrales.** ¿Cada umbral apunta en la dirección correcta —más runway no
   puede ser peor calificación— y corresponde a la definición vigente del indicador? Un
   umbral heredado de una definición vieja pinta rojo permanente y deja de informar.
5. **`s/d` frente a `0` frente a `n/a`.** ¿La convención se aplica de forma consistente?
   Un cero donde debería ir `s/d` inventa un dato; un `s/d` donde el corte simplemente no
   existe para esa entidad manda a perseguir algo que nadie va a entregar.
6. **Agregación.** Los acumulados, ¿suman los meses o reconvierten un total? Los
   promedios, ¿son simples o ponderados, y es lo que corresponde? Las series con meses
   faltantes, ¿los saltan o los cuentan como cero?
7. **Fallbacks silenciosos.** Busca sustituciones que el código haga por debajo sin
   declararlo: caer a otro país, a otro mes, a otra fuente. Cada una puede ser correcta o
   no, pero **ninguna** debería ser invisible para quien lee.
8. **El tipo de gráfico, ¿es el adecuado?** Una barra apilada de conceptos que no suman a
   un total con sentido, o una línea sobre datos no continuos, comunican mal aunque los
   números estén bien.

## 11. Qué entregar

Un informe con:

1. **Veredicto en una frase**: ¿se puede mostrar a una junta o a un inversionista, sí o no?
2. **Hallazgos**, cada uno con:
   - Identificador y severidad: **Crítico** (cifra mal, contradicción entre pantallas, o
     una función que no hace lo que promete) · **Importante** (dato faltante, rótulo
     engañoso, semáforo mal calibrado) · **Menor** (presentación, robustez).
   - Pestaña y control exactos donde se ve.
   - **Evidencia numérica**: la cifra que muestra el tablero, tu recálculo independiente,
     y la diferencia.
   - **Causa**, con la función o la hoja concreta. Si no la encontraste, dilo — no la
     inventes.
   - Si es problema de **código** o de **datos**.
3. **Lo que está bien.** Nombra explícitamente lo que verificaste y cuadra. Un informe que
   solo lista defectos no permite saber qué alcance tuvo la revisión.
4. **La matriz pestaña × filtro** completa: qué mueve cada control en cada pestaña.
5. **Qué quedó sin probar** y por qué.

Ordena los hallazgos por severidad y, dentro de cada nivel, por el daño que hacen a quien
lee el tablero sin conocerlo. Un rótulo que induce a leer mal una cifra correcta hace el
mismo daño que la cifra equivocada: van al mismo nivel.

## 12. Reglas de conducta

- **No inventes datos ni causas.** Si una cifra no cuadra y no encuentras por qué, repórtala
  como descuadre sin explicación. Es más útil que una hipótesis presentada como hecho.
- **No modifiques nada**: ni la base, ni el HTML, ni la nube. Esto es una lectura.
- Si algo exige una **decisión de negocio** para saber si está bien o mal, no la tomes:
  formula la pregunta.
- Distingue **«está mal»** de **«no me gusta»**. Lo segundo va en Menores, y con esa etiqueta.
- Cuando reportes una diferencia, dila en la moneda y el periodo en que la mediste.
