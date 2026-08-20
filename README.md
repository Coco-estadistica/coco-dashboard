# COCO Tecnologías — Cockpit Financiero

Tablero financiero y operativo multipaís (Colombia, EE.UU., Perú, Costa Rica).
Corre en tu computador leyendo un Excel, y está preparado para publicarse en Google.

Última revisión de este documento: **20 de agosto de 2026**.

---

## 1. Uso diario — tres botones

En esta carpeta hay tres archivos `.bat`. **Doble clic, nada más.**

| Botón | Para qué | ¿Modifica algo? |
|---|---|---|
| **`Actualizar_Desde_GitHub.bat`** | Traer la última versión. **Empieza siempre por aquí** | Sí: baja archivos |
| **`Abrir_Dashboard_COCO.bat`** | Abrir el tablero | No |
| **`Revisar_Base.bat`** | Comprobar que la base esté sana, antes de publicar un cierre | **No** |
| **`Corregir_Base.bat`** | Alinear capas desfasadas. Muestra qué cambiaría y pide confirmación | Sí, con respaldo |
| **`Subir_A_GitHub.bat`** | Guardar tu trabajo en GitHub. **Termina siempre por aquí** | Sí: publica |

Los archivos `.py` que hay en las subcarpetas son las instrucciones que ejecutan esos
botones. **No hay que abrirlos ni ejecutarlos a mano.**

### 1.1 Una sola copia manda: la de GitHub

El mismo Excel puede existir en tu computador, en GitHub y en Google Sheets. Para que no
se repita el problema de la sección 3 —una cifra en varios sitios desalineándose sin
avisar—, **la copia oficial es la de GitHub**. La de tu computador es la copia de
trabajo: se baja antes de empezar y se sube al terminar.

**La regla que no se puede romper: un `.xlsx` es binario, y Git no lo puede fusionar.**
Si dos personas editan la base a la vez, no hay forma de combinar los dos archivos —
uno gana y el otro se pierde entero. Por eso los botones nunca fusionan: si detectan
que hay trabajo cruzado, se detienen y avisan en lugar de arriesgar cifras.

De ahí salen dos costumbres:

- **Antes de tocar nada**, `Actualizar_Desde_GitHub.bat`. Si tienes cambios sin subir,
  el botón se detiene y te lo dice: primero sube, luego baja.
- **Mientras alguien más esté trabajando sobre la base**, tú no la editas. Se avisa
  cuando queda libre. Nunca los dos a la vez sobre el mismo archivo.

> Si al abrir el tablero aparece "Carga la base de datos", espera unos segundos: la base
> vive en OneDrive y a veces tarda en descargarse. El tablero reintenta solo 3 veces.

---

## 2. Los dos libros de datos

La base se separó en dos el 13-ago-2026 para eliminar la duplicación de cifras.

### `migracion/BD_MAESTRA_COCO.xlsx` — 14 hojas · la que importa

Es la **única fuente de verdad**. La que lee el tablero, la que se respalda y la que
se sube a Google Sheets.

| Hoja | Papel |
|---|---|
| `BD_Indicadores` | **La tabla de hechos.** Todo el tablero sale de aquí |
| `Diccionario` | Catálogo de indicadores. Muchas hojas lo referencian con fórmulas |
| `TRM` / `TRM_Peru` | Tasas de cambio COP/USD y PEN/USD |
| `Calc_LTV_Fin` | Cálculo de LTV. Referenciado por `BD_Indicadores` y `Marketing` |
| `Analisis_Churn` / `Analisis_Reconciliacion` | Insumos de pestañas específicas |
| `Pipeline_Comercial` | Oportunidades comerciales |
| `CATALOGOS` | Qué códigos puede cargar cada área (lo usa Apps Script) |
| `LOG` | Bitácora de quién cargó qué (lo escribe Apps Script) |
| `MAP_DASHBOARD_GASTOS`<br>`BD_GASTOS_HOMOLOGADOS`<br>`BD_GASTOS_DASHBOARD_BRIDGE`<br>`BD_GASTOS_RESUMEN_MENSUAL` | **Cadena de gastos.** Las tres primeras producen la cuarta, que es la que el tablero lee. Se mantienen juntas porque son una cadena de producción activa |

### `migracion/BD_TRAZABILIDAD_COCO.xlsx` — 24 hojas · el archivo

Se conserva íntegro para auditoría. **El tablero nunca lo abre.**

- **Fuentes** (`EEFF_*`): de dónde salió cada cifra, homologaciones, movimientos,
  intercompañías, alertas y controles de los estados financieros.
- **Vistas por área** (`Finanzas`, `Contabilidad`, `Ventas`, `Marketing`,
  `Customer_Success`, `SaaS_Revenue`): se regeneran desde `BD_Indicadores`. No aportan
  dato nuevo.
- **Capas de soporte** (`BD_PYG_OFICIAL`, `Subsidiarias_PyG`, `PyG_Oficial`,
  `PyG_Mensual_Col`): repiten cifras que ya están en `BD_Indicadores`.

---

## 3. Por qué se separaron

La misma cifra vivía en tres o cuatro hojas. Cuando se actualizaba una y se olvidaban
las otras, **el tablero mostraba datos viejos sin avisar**. Pasó dos veces:

- **Perú julio**: cargado en el consolidado estructurado, pero en cero en la hoja que
  el tablero lee.
- **Costa Rica junio**: se corrigió en `BD_Indicadores` y `BD_PYG_OFICIAL`, pero
  `Subsidiarias_PyG` conservó la cifra vieja durante semanas.

Con la separación, **solo hay una capa que el tablero consume**. El resto es archivo
que nadie lee, así que no puede desalinearse.

El corte se decidió con tres verificaciones, no a ojo:

1. Se leyó el parser del tablero: abre exactamente 10 hojas.
2. Se mapearon las 5.229 fórmulas: solo `Diccionario` y `Calc_LTV_Fin` son
   referenciadas por otras hojas → se quedan.
3. Las `EEFF_*` y los puentes no tienen ninguna fórmula cruzada → moverlas es seguro.

---

## 4. Las 13 pestañas del tablero

| # | Pestaña | Qué muestra |
|---|---|---|
| 01 | **Resumen** | KPIs del mes y del año, Real vs Presupuesto. Es la que abre |
| 02 | Consolidado | Vista CEO: multipaís en USD, P&G por subsidiaria, puentes de resultado |
| 03 | **Burn & Runway** | Caja, quema y cobertura. Se alimenta del Modelo, no de la base |
| 04 | Finanzas & P&L | Indicadores de gestión, facturación contable vs ejecutada |
| 05 | P&G Colombia | Estado de resultados mensual con las líneas contables oficiales |
| 06 | Ingresos & MRR | Recurrencia y movimiento de MRR |
| 07 | Cartera & Liquidez | Cobranza y caja |
| 08 | Gastos Colombia | Costos por área, **solo Colombia** |
| 09 | Gastos Consolidado | Costos por área, **Colombia + EE.UU. + Perú + Costa Rica** |
| 10 | Ventas & Comercial | Pipeline y cierre |
| 11 | Marketing | Adquisición y eficiencia |
| 12 | Customer Success & Churn | Clientes, NPS y churn |
| 13 | Catálogo | Todos los indicadores con su definición y valor |

Las tres áreas comerciales (10 a 12) van al final **a propósito**: son las de menor
cobertura de datos —MRR y Customer Success llegan a mayo, Marketing a junio— mientras
que lo financiero va completo a julio. Así el tablero abre por lo que está al día.

**Gastos Colombia** y **Gastos Consolidado** eran una sola pestaña ("Costos & Gastos")
que seguía el selector de país de arriba — causaba la pregunta recurrente de "por qué
si cambio el país no cambian los números", porque no siempre era obvio que ese selector
también decidía lo que se veía ahí. Ahora son dos pestañas fijas: cada una fuerza su
propio país al entrar y el selector superior no puede sacarlas de su alcance mientras
se está en ellas. "Consolidado" es la suma real de los 4 países (con el arreglo del
puente de gastos de la sección 4.2), no la entidad "Consolidado" del selector — esa es
un corte contable más angosto que no trae desglose por área.

**Filtros**: rango de fechas, escenario (Real/Presupuesto), país, moneda (COP/USD) y
base de TRM (implícita de EEFF o promedio de mercado).

### 4.1 Cómo leer los vacíos

Todos los gráficos distinguen **tres** situaciones. La leyenda aparece bajo las pestañas
**solo en las vistas que tienen vacíos**; donde no falta nada, no estorba.

| Se ve | Significa | Qué hacer |
|---|---|---|
| Una cifra, p. ej. `$0` | **Cero reportado.** La fuente dice que ese mes fue cero | Nada. Es un dato |
| `s/d` sobre una línea punteada | **Dato no recibido.** La cuenta existe para ese país, pero el mes no se cargó | Perseguirlo antes del cierre |
| `n/a` | **No aplica.** Esa cuenta no pertenece a esa entidad | Nada. No falta nada |

La diferencia entre `s/d` y `n/a` **no está escrita en una lista fija**: si un código no
aparece nunca para ese país, no aplica. Así no hay que mantener excepciones a mano.

Ejemplo real, churn mensual: `abr 26 $326 · may 26 $0 · jun 26 s/d · jul 26 s/d`. Mayo
fue cero de verdad; junio y julio todavía no se han cargado. Antes ambos casos se veían
igual, y **un mes sin cargar se leía como una caída a cero**.

Las líneas de tendencia además **se cortan** en los meses sin dato: unir los extremos de
un hueco dibujaría una pendiente que nadie reportó.

### 4.2 El puente de gastos y Colombia

`Gastos Colombia` y `Gastos Consolidado` leen `BD_GASTOS_RESUMEN_MENSUAL` (el puente que
homologa los auxiliares de EE.UU./Perú/Costa Rica). Ese puente **nunca cubrió a
Colombia** — solo trae una fila puente en cero para julio. Hasta el 13-ago-2026 eso hacía
que:

- **"Consolidado" mostrara solo las filiales**, sin avisar: gastos comerciales de abril
  aparecían en $247 cuando Colombia sola movía $31.175 ese mes.
- **El desglose por área de Colombia colapsara** a un solo bloque gris "Sin desglose
  contable", aunque los 5 rubros reales (Financieros, TI, Proveedores, Comercial,
  Administración) seguían completos en `BD_Indicadores` desde 2024.

Se corrigió inyectando el detalle real de Colombia en el puente cuando este no la trae
para un código dado (`monthlyGastosBridgeData`, `Dashboard_COCO.html`). Si algún día se
homologa a Colombia dentro de `BD_GASTOS_RESUMEN_MENSUAL`, ese detalle real tiene
prioridad y el ajuste no duplica nada.

---

## 5. Cierre mensual

1. **`Actualizar_Desde_GitHub.bat`** — parte siempre de la última versión.
2. **Respalda** la base (o corre `Corregir_Base.bat`, que respalda solo).
3. Revisa la fuente de cada país **en su moneda de origen**.
4. Carga las cifras en `BD_Indicadores`.
5. Recalcula el bloque `Consolidado` en USD.
6. **Corre `Revisar_Base.bat`.** Si dice "todas las capas están alineadas", sigue.
7. Abre el tablero y revisa las pestañas afectadas.
8. Confirma que los periodos sin datos aparezcan como **`s/d`** y no como cero.
   Cada `s/d` es un pendiente de cargue; si sobra alguno, es que el dato ya llegó
   y no se subió.
9. **`Subir_A_GitHub.bat`** — el cierre no está terminado hasta que está publicado.
   El botón te muestra qué va a subir y pide confirmación antes de hacerlo.

### Si además actualizaste el Modelo financiero

> **Ojo: la carpeta `modelo/` no está en el repositorio de GitHub.** Solo existe en la
> copia local. Por eso `Revisar_Base.bat` reporta *"No encuentro modelo/Modelo_COCO.xlsx
> — no se puede verificar"* en el control E cuando se corre desde un clon del repo: no es
> una avería, es que el modelo nunca se subió. La pestaña 03 sí funciona, porque se pinta
> con `burn_runway.json`, que sí está versionado. Lo que no se puede desde el clon es
> **regenerar** ese JSON ni validar su huella. Pendiente de decidir: subir la carpeta o
> dejar constancia de que ese control solo corre en la máquina que tiene el modelo.

- Reemplaza `modelo/Modelo_COCO.xlsx` por la versión nueva.
- Doble clic en `modelo/Actualizar_Burn_Runway.bat` → refresca la pestaña 03.
- Si el modelo trae un cierre posterior a junio, **cambia el `corte`** en
  `modelo/extraer_burn_runway.py`: está escrito a mano y no se deduce solo.

`Revisar_Base.bat` compara la huella del modelo contra la que quedó sellada en
`burn_runway.json`. Si reemplazas el modelo y olvidas regenerar, avisa
**DESACTUALIZADO** en vez de dejar la pestaña mostrando cifras viejas. También avisa
si aparece un modelo más reciente en `Modelos Financieros` que aún no se ha copiado.

---

## 6. Reglas que no se deben romper

1. **No inventar datos.** Un dato que no llegó se marca **`s/d`**, **nunca cero**
   (ver sección 4.1).
2. **Perú se origina en PEN** y se convierte con `TRM_Peru`. Su proyecto especial
   (PROINNOVATE) va aparte y **no entra al consolidado**.
3. **Colombia se origina en COP.** El PDF oficial manda sobre el acumulado.
4. **No sumar monedas distintas.** El consolidado multipaís se presenta en USD.
5. **Acumulado ≠ movimiento del mes.** Para un mes se usa el movimiento del auxiliar
   (débitos/créditos), no el saldo.
6. Al cambiar cifras de un país, **recalcular el consolidado**.
7. Dejar la fuente, la tasa y el alcance en la columna `comentario`.
8. **No borrar gráficos ni funcionalidades** para resolver un problema de datos.

---

## 7. Cobertura actual (agosto 2026)

| Área | Último periodo con datos |
|---|---|
| P&G (Colombia, EE.UU., Perú, Costa Rica) | **julio** |
| Pipeline comercial | julio |
| Detalle contable colombiano | junio |
| ARR · Marketing · Cartera y liquidez | junio |
| MRR · Customer Success | mayo |

No completar estas áreas con estimaciones.

Los huecos hoy visibles en el tablero, y todos verificados contra la base:

- **Costos & Gastos**: 8 meses (ago-25 a mar-26). El P&G multipaís arranca en abril 2026.
- **Customer Success**: junio y julio de churn, y el NPS de ago-25 a feb-26.
- **Consolidado**: la reclasificación de honorarios de Perú sale `n/a` en el puente de
  Colombia — es un ajuste de consolidación que solo vive en EE.UU. y Consolidado.
- **Pipeline comercial**: la hoja `Pipeline_Comercial` de la base está **vacía** (solo la
  fila de encabezado), pese a que la tabla de arriba declara cobertura hasta julio.
  Verificado el 20-ago-2026. Es un **`s/d`**, no un pipeline en cero: hay que perseguir
  el dato con Comercial, no rellenarlo.

---

## 8. Cosas que conviene saber

- **Corregido (13-ago-2026): cualquier filtro devolvía a Resumen en silencio.** Cambiar
  país, moneda, fecha o escenario reconstruye el menú de pestañas (`buildTabs()`), y ese
  menú por diseño marca la primera pestaña como activa. Sin guardar cuál estaba activa
  antes de reconstruir, cada cambio de filtro sacaba a quien estuviera en cualquier otra
  pestaña y lo dejaba viendo Resumen sin avisar. Ya estaba así desde antes de esta sesión
  (confirmado contra el respaldo de las 11:08 del 13-ago). Se corrigió en `render()`.
- **Costa Rica factura pero casi no tiene costos.** Su margen aparente (~98%) no es un
  error de captura: los costos de entregar ese servicio están en otra entidad. Es una
  **pregunta de precios de transferencia** pendiente de resolver.
- **La utilidad de julio no trajo caja.** Los USD 134.104 de Costa Rica se registraron
  contra Clientes, no contra el banco. Sin Costa Rica, julio da pérdida de USD 886.
- **La tasa de Perú de julio ya es definitiva** (corregido el 18-ago-2026). `TRM_Peru`
  trae **3,400 PEN/USD**, promedio mensual completo del BCRP (punto medio compra/venta
  SBS). Ya no hay que reemplazarla: hasta el 13-ago era provisional, cortada al 9 de
  julio.
- **Burn & Runway va un mes atrás.** Está cortada a **junio**, mientras las otras once
  pestañas van a julio. Además convierte la caja de EE.UU. a **3.248,87** (tasa de
  cierre de un saldo) y no a los 3.505 del resto. No compares esas cifras de frente.
- **Corregido (20-ago-2026): el julio de Colombia en `BD_PYG_OFICIAL` estaba convertido
  a la TRM de junio.** Cuando el 14-ago se pasó la TRM de julio de 3.505,2683
  (autoderivada) a 3.268,95 (oficial), se reconvirtió `BD_Indicadores` pero no la capa de
  archivo. Las 12 líneas del P&G de Colombia diferían con un **ratio idéntico de
  1,072292** — que es exactamente 3.505,2683 / 3.268,95. Un ratio constante en todos los
  rubros es diferencia de *tasa*, no de cifra: sirve para diagnosticar de un vistazo.
  El tablero nunca mostró mal ese dato, porque no lee esa hoja. Se corrigió con
  `migracion/corregir_trm_julio_trazabilidad.py`, recalculando desde la fila **COP** de la
  propia hoja (Colombia se origina en COP), no copiando de `BD_Indicadores`. El
  Consolidado se ajustó **sumando el delta de Colombia**, no igualándolo a
  `BD_Indicadores`: igualarlo habría borrado también el desfase de EE.UU./Perú que está
  congelado a propósito desde el 13-ago. El control C bajó de 86 a 67 diferencias, y las
  5 que quedan en julio están atribuidas al céntimo a ese desfase congelado.
- **Las alertas de `EEFF_Alertas`** (libro de trazabilidad) se cierran, no se borran.
  El 13-ago-2026 se cerró la de cobertura de Costa Rica y se anotó lo verificado sobre
  los PEN 64.960,94 de Perú (es saldo acumulado, no ingreso de julio). Las de alta
  severidad sobre **tasas de conversión** y **eliminaciones intercompañía** siguen
  abiertas.
- **Texto con caracteres rotos** en algunas hojas heredadas (`Administraci?n`). No
  cambiar la codificación en masa sin comparar el comportamiento antes y después.

---

## 9. Publicar en Google (opcional, preparado y sin desplegar)

Ver **`DESPLIEGUE.md`**. Resumen: el Sheet queda como fuente de datos, Apps Script
sirve el tablero, y cada área carga su porción con su cuenta `@cocotech.ai` — sin
claves que repartir, con la identidad verificada por Google y todo registrado en `LOG`.

---

## 10. Para quien retome este proyecto

Lee también **`CLAUDE_HANDOFF.md`** (contexto de agosto 2026), pero ten en cuenta que
este README es posterior y corrige dos cosas que allí quedaron desactualizadas:

- La base ya **no** tiene 38 hojas en un solo archivo: se separó en dos.
- `Subsidiarias_PyG` **no** alimenta el tablero; es capa de soporte.

Antes de cambiar cifras: corre `Revisar_Base.bat`, y traza cada indicador hasta la
función del HTML y la hoja exacta que lo alimenta. No asumas que una hoja llamada
"oficial" es la que el tablero consume.
