# Master prompt — Auditoría de vibe engineering del Cockpit COCO

> Cópialo completo en una sesión nueva. Está escrito para que quien lo reciba no
> necesite nada de la conversación anterior.

---

Vas a auditar **cómo se construyó** un sistema, no el sistema. Es la contraparte de una
auditoría de cifras: aquella pregunta *¿los números están bien?*; esta pregunta
*¿aguanta que siga creciendo, y puede alguien distinto verificarlo?*

El sistema se construyó casi enteramente con un agente de IA, sobre meses y muchas
sesiones. Eso no es un defecto — es el contexto que define qué infraestructura de
verificación hace falta.

**La premisa que ordena todo el ejercicio:** cuando escribir código deja de ser el cuello
de botella, el cuello de botella pasa a ser **verificarlo**. Así que no midas calidad de
código. Mide **capacidad de verificación**.

---

## 1. Qué se audita

| | |
|---|---|
| Repositorio | `Moscorrofio/coco-dashboard` (privado) · rama por defecto `main` |
| Producto | `Dashboard_COCO.html` — **4.097 líneas, 977 KB, un solo archivo**, 194 funciones |
| Datos | `migracion/BD_MAESTRA_COCO.xlsx` — binario, versionado en git |
| Scripts | 39 en Python, de los cuales **20 son de un solo uso** (`cargar_*`, `corregir_*`) |
| Historial | 76 commits · 8 ramas · `.git` pesa **27 MB** · **52 archivos .xlsx versionados**, 49 de ellos respaldos |
| Automatización | 6 archivos `.bat` · **sin GitHub Actions** · **sin tests** |

Verifica estos números antes de empezar. Si no coinciden, el repositorio cambió y hay que
rehacer las mediciones de abajo.

## 2. Método

**No leas el código buscando fealdad.** Busca respuestas a cinco preguntas, cada una con
una medición que se pueda repetir dentro de tres meses y comparar.

Usa el historial de git como evidencia principal: los mensajes de commit de este
repositorio son inusualmente detallados y documentan qué se rompió, por qué, y cómo se
midió. Son la mejor fuente sobre el proceso real, mejor que el código.

---

## 3. Los cinco ejes

### Eje 1 · ¿Puede alguien que no sea el agente comprobar que esto está bien?

Es el eje que más rinde. Mídelo así:

1. **Cuenta los tests.** Archivos `test_*.py`, `*.test.js`, suites de cualquier tipo.
2. **Cuenta las cifras de control que existen en algún archivo ejecutable.** El sistema
   tiene cifras conocidas —el P&G de julio 2026, el puente de MRR, la cartera— que se han
   verificado a mano muchas veces. ¿Cuántas están escritas en un archivo que se pueda
   correr?
3. **Busca en el historial las verificaciones que se hicieron y se tiraron.** Los commits
   citan comparaciones de cientos de cifras. ¿Sobrevivió alguna como test?

**Lo que esperas encontrar:** cero. Documenta el tamaño exacto del hueco: cuántas
verificaciones se han hecho a mano y no quedaron.

### Eje 2 · ¿Cuántas veces se rompió algo que ya estaba arreglado?

**Esta es la métrica central de la disciplina.** Mide si el sistema se defiende solo o
depende de que alguien recuerde.

Recorre el historial buscando pares: un commit que arregla X, y un commit posterior que
vuelve a arreglar X o algo de la misma familia. El repositorio tiene al menos estos casos
documentados en sus propios mensajes — confírmalos y busca más:

- `factVal` se corrigió para que filtrara por escenario; **horas después** se escribieron
  dos funciones nuevas con el mismo defecto (escenario clavado en `"Real"`).
- Un `pushState` quedó después del render, que hacía `replaceState`: el hash ya coincidía
  cuando llegaba el push, no crecía el historial, y el botón Atrás seguía sacando del
  tablero — el síntoma exacto que ese commit venía a corregir.
- Una tarjeta mostraba una suma sin sentido; alguien **renombró el rótulo** en vez de
  corregir la suma, y la cifra siguió equivocada bajo un título ahora literalmente cierto.
- El trabajo quedó atrapado en una rama de trabajo sin llegar a `main` **dos veces en la
  misma jornada**.

**Reporta una tasa**, no una lista: regresiones por cada N commits. Y clasifícalas: ¿la
habría atrapado un test? ¿un linter? ¿una revisión humana? Esa clasificación es la que
dice qué construir primero.

### Eje 3 · ¿El agente puede verificar su propio trabajo sin un humano mirando?

Aquí el sistema tiene algo bueno y conviene decirlo con nombre propio:

- `migracion/exportar_sqlite.py` — espeja la base a SQLite y corre validaciones de negocio
  (consolidado = suma de países, utilidad = ingresos − gastos, morosidad, composición del
  NPS, llaves duplicadas, cifras reales en meses futuros).
- `migracion/conciliar_capas.py` — compara las capas de la base entre sí.
- `migracion/regenerar_gastos.py --probar` — reconstruye la cadena de gastos sobre una
  copia y compara **celda por celda**, no solo totales.
- `migracion/cargar_mes.py` — cinco controles sobre los archivos del ERP antes de cargar
  nada: que abran, que sean del mes que dicen, que cuadre la partida doble, que el árbol
  de cuentas sume igual en cada nivel, y que el mapeo reproduzca un mes ya cargado.

**Evalúa dos cosas:** ¿qué cubren y qué no? Y sobre todo: **¿corren solos?** Un control
que depende de que alguien le dé doble clic a un `.bat` no es infraestructura, es un
recordatorio.

### Eje 4 · ¿Qué tan reversible es cada cambio?

1. **El código.** Git, sí. Pero mide la disciplina real: ¿cuántas ramas quedaron sin
   fusionar? ¿Cuánto tiempo pasó entre un commit y su llegada a `main`?
2. **Los datos.** La base es un `.xlsx` **binario**. Git no puede fusionarlo, no puede
   mostrar qué cambió, y dos personas editándolo a la vez pierden trabajo. La estrategia
   de respaldo son copias con fecha en el nombre: **49 archivos**, 27 MB de historial.
   Evalúa si eso escala y qué lo reemplazaría.
3. **La superficie de un solo punto de fallo.** Un archivo de 4.097 líneas con 194
   funciones, editado por parches de texto exacto. ¿Qué pasa si dos cambios tocan la
   misma zona?

### Eje 5 · ¿Está claro qué decide el humano y qué decide el agente?

**Este eje probablemente sale bien, y hay que decirlo.** Busca en el historial las
decisiones de negocio: la definición de morosidad, cuál ARR es el vigente, cómo se
presenta el crecimiento interanual, si Costa Rica está confirmada. Comprueba que:

- están tomadas por una persona, no por el agente;
- quedaron registradas con su fecha;
- el código las cita donde se aplican;
- y que el agente **preguntó** en vez de decidir cuando la respuesta cambiaba las cifras.

Busca también el contraejemplo: casos donde el agente decidió algo que debió preguntar.

---

## 4. Una medición incómoda que no debes omitir

**Cuánto del código lo escribió un agente sin que ninguna persona lo leyera.**

En este repositorio la respuesta se acerca al total. Eso **no es un defecto en sí mismo** —
es el dato que determina cuánta infraestructura de verificación hace falta. Un equipo con
revisión entre pares puede permitirse menos tests; uno sin ella, no.

Dilo sin dramatismo y sin suavizarlo, y deriva de ahí la recomendación.

## 5. Qué entregar

1. **Veredicto en una frase.** ¿Este sistema aguanta seis meses más de crecimiento al
   ritmo actual, sí o no?
2. **Los cinco ejes**, cada uno con su medición, su evidencia del historial y su
   diagnóstico.
3. **La tasa de regresión**, con los casos concretos y la clasificación de qué los habría
   atrapado.
4. **Lo que ya está bien.** Nombra explícitamente lo que no hay que tocar. Un informe que
   solo lista carencias no permite saber qué se revisó ni protege lo que funciona.
5. **Tres cosas que construir, en orden**, con el criterio de por qué ese orden. No una
   lista de buenas prácticas: tres cosas, priorizadas, con lo que cada una previene.

## 6. Reglas de conducta

- **No propongas prácticas genéricas.** «Agregar tests» no es un hallazgo. *«Las 20 cifras
  de control de julio se han verificado a mano al menos cuatro veces y ninguna quedó en un
  archivo ejecutable»* sí lo es.
- **Ancla cada hallazgo en evidencia del repositorio**: un commit, un archivo, una
  medición. Si no puedes anclarlo, no lo reportes.
- **No confundas deuda con defecto.** Un archivo de 4.000 líneas no está mal por ser
  grande; está mal si impide verificar o si dos cambios se pisan. Demuéstralo o no lo
  afirmes.
- **No arregles nada.** Esto es una lectura del proceso. Si encuentras algo urgente
  —trabajo sin fusionar, un respaldo que se perdió— dilo primero y aparte.
- Recuerda que el sujeto de esta auditoría es **el método**, no las personas ni el agente.
  El objetivo es que la próxima sesión sea más segura que la anterior.
