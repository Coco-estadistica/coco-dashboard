# Scripts ya ejecutados — no volver a correr

Estos 22 scripts hicieron su trabajo una vez y **ya no forman parte del cierre mensual**.
Se guardan porque documentan cómo entró cada dato a la base, no porque sirvan para
volver a entrarlo.

**Lo que hacen hoy lo hace `migracion/cargar_mes.py`**, con controles que estos no tenían:
que el archivo abra, que sea del mes que dice, que cuadre la partida doble, que el árbol
de cuentas sume igual en cada nivel, y que el mapeo reproduzca un mes ya cargado.

## Por qué se archivaron y no se borraron

Dos razones. La primera es que explican decisiones que el código vigente ya no muestra:
por qué julio de Perú se recargó, con qué criterio se homologaron los gastos de Colombia,
de dónde salió la cifra de junio del porcentaje de proveedores.

La segunda es que **son armas cargadas**. Correr cualquiera de ellos hoy volvería a
escribir cifras que ya se corrigieron. Sacarlos de la carpeta de trabajo es la forma
barata de que nadie lo haga por accidente.

## Los cuatro que más daño harían

| Script | Qué pasa si se corre |
|---|---|
| `cargar_cartera_julio.py` | Tiene `TRM_JULIO = 3505.2683` escrito dentro, que es **la TRM de junio**. Volvería a inflar la cartera de julio un 7,2% — el error que se corrigió el 27-ago |
| `construir_bd_maestra.py` | Reconstruye la base desde cero y **escribe sin pedir `--escribir`** |
| `agregar_h1_2026.py` | Vuelve a agregar el acumulado del semestre, duplicándolo. Escribe directo |
| `agregar_presupuesto.py` | Vuelve a agregar las tres versiones del presupuesto. Escribe directo |

Los tres últimos no tienen la protección de `--escribir` que sí tienen los demás: un doble
clic basta.

## Si necesitas algo de aquí

No lo corras: **léelo**. Si el procedimiento sigue siendo válido, llévalo a `cargar_mes.py`
en vez de resucitar el script. Y si de verdad hay que correr uno, copia primero la base —
todos escriben sobre `BD_MAESTRA_COCO.xlsx`.

Nada de esto se pierde: `git log --follow migracion/_ejecutados/<archivo>` tiene la historia
completa, y `git mv` conservó la de antes del archivado.
