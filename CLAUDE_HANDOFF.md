# Traspaso del proyecto COCO Dashboard

Estado documentado: 13 de agosto de 2026.

## 1. Objetivo

Mantener un dashboard financiero y operativo multipais para COCO Tecnologias. El dashboard consolida Colombia, EE.UU., Peru y Costa Rica, permite filtrar periodos y moneda, y consume una base maestra Excel o su equivalente en Google Sheets.

La prioridad inmediata es estabilizar el dashboard, asegurar que cada grafico use datos existentes y evitar que datos faltantes se representen como ceros.

## 2. Archivos oficiales

- Dashboard local: `COCO_Dashboard_Cloud/Dashboard_COCO.html`.
- Base local oficial: `COCO_Dashboard_Cloud/migracion/BD_MAESTRA_COCO.xlsx`.
- Lanzador habitual: `COCO_Dashboard_Cloud/Abrir_Dashboard_COCO.bat`.
- Implementacion Google Apps Script: `COCO_Dashboard_Cloud/apps_script/Code.gs` y `Dashboard.html`.
- Fuente financiera estructurada: `Consolidado Financiero Estructurado COCO 2026.xlsx` en la raiz del proyecto.
- Respaldo anterior a la ultima correccion: `COCO_Dashboard_Cloud/migracion/BD_MAESTRA_COCO_RESPALDO_ANTES_PERU_JULIO_20260813.xlsx`.

No usar como base vigente los archivos con nombres `ORDENADA`, `FINAL`, `ACTUALIZADA` o `RESPALDO`. Son versiones historicas.

Los tres HTML actualmente existentes (`Dashboard_COCO.html`, `Dashboard_COCO_CORREGIDO_JULIO_2026.html` y `Dashboard_COCO_JULIO_2026_R4.html`) tienen el mismo SHA-256. Conviene conservar uno como canon y convertir los otros en respaldos o eliminarlos solo con autorizacion del usuario.

## 3. Ejecucion local

El lanzador inicia `python -m http.server 8740` y abre la version R4. El HTML servido por HTTP intenta cargar automaticamente:

`migracion/BD_MAESTRA_COCO.xlsx`

URL habitual: `http://localhost:8740/` o `http://localhost:8740/Dashboard_COCO.html`.

No evaluar la carga automatica abriendo el HTML mediante `file://`; en ese modo no puede leer la base por `fetch` y puede mostrar cache antiguo del navegador.

## 4. Flujo real de datos

El parser principal esta dentro de `Dashboard_COCO.html`, funcion `parseWorkbook()`.

- `BD_Indicadores`: fuente principal de los KPI y del P&G que se visualiza.
- `BD_PYG_OFICIAL`: soporte financiero oficial y trazabilidad; no reemplaza la necesidad de sincronizar `BD_Indicadores`.
- `Subsidiarias_PyG`: resumen de P&G por subsidiaria.
- `TRM`: conversion COP/USD.
- `TRM_Peru`: conversion PEN/USD.
- `BD_GASTOS_RESUMEN_MENSUAL`: puente para graficos de costos y gastos.
- `BD_GASTOS_HOMOLOGADOS` y `BD_GASTOS_DASHBOARD_BRIDGE`: detalle y homologacion de gastos.
- `Pipeline_Comercial`, `Marketing`, `Customer_Success`, `Analisis_Churn`, `Calc_LTV_Fin` y otras hojas alimentan pestanas operativas especializadas.

Regla critica: cualquier correccion del P&G debe reflejarse de manera consistente en `BD_Indicadores`, `BD_PYG_OFICIAL` y `Subsidiarias_PyG`. Si cambia un pais, tambien debe recalcularse el bloque `Consolidado`.

## 5. Ultima correccion: Peru julio 2026

La fuente `Consolidado Financiero Estructurado COCO 2026.xlsx`, hoja `Mov_Mensuales`, filas 514:536, contiene movimientos ordinarios de Peru de julio en PEN. Antes de esta correccion, la base maestra contenia ceros.

Se cargaron los movimientos ordinarios y se excluyo el proyecto especial de Peru. Tambien se excluyo el saldo acumulado de `Otros ingresos` por PEN 64,960.94 porque no tuvo movimiento en julio y requiere clasificacion separada.

Tasa aplicada desde `TRM_Peru`:

- PEN/USD: 3.405.
- USD/PEN: 0.293686.
- Estado: provisional, promedio hasta el 9 de julio.

Peru julio cargado:

| Rubro | PEN | USD |
|---|---:|---:|
| Ingresos operacionales | 6,202.25 | 1,821.51 |
| Gastos de administracion | 878.00 | 257.86 |
| Gastos de ventas | 11,218.31 | 3,294.66 |
| Gastos financieros | 314.15 | 92.26 |
| Utilidad bruta | 6,202.25 | 1,821.51 |
| Utilidad operativa | -5,894.06 | -1,731.00 |
| Utilidad neta | -6,208.21 | -1,823.26 |

Consolidado julio resultante:

| Rubro | USD |
|---|---:|
| Servicio de software | 345,033.44 |
| Devoluciones | -9,132.00 |
| Ingresos operacionales | 335,901.45 |
| Costo de ventas | 91,838.29 |
| Utilidad bruta | 244,063.16 |
| Administracion | 46,098.13 |
| Proyectos | 29,366.19 |
| Ventas | 36,890.16 |
| Utilidad operativa | 131,708.69 |
| Otros ingresos | 215.68 |
| Gastos financieros | 1,783.85 |
| Utilidad neta | 130,140.53 |

Las cifras quedaron sincronizadas en `BD_Indicadores`, `BD_PYG_OFICIAL` y `Subsidiarias_PyG`.

## 6. Reglas financieras que no deben romperse

1. No inventar ni extrapolar datos faltantes.
2. Un dato no recibido debe ser `null`, vacio o mostrarse como `Dato no recibido`; nunca cero salvo que la fuente confirme cero real.
3. Peru se origina en PEN y debe convertirse mediante `TRM_Peru`.
4. El proyecto especial de Peru debe permanecer separado del P&G ordinario.
5. Colombia se origina en COP; el PDF oficial de Colombia controla el acumulado oficial.
6. No sumar monedas locales diferentes. El consolidado multipais se presenta en USD.
7. Los acumulados no deben confundirse con movimientos mensuales. Para julio se usaron movimientos del auxiliar o diferencias YTD solo cuando la fuente lo permitia.
8. Al cambiar cifras por pais, recalcular ingresos, utilidad bruta, utilidad operativa y utilidad neta del consolidado.
9. Mantener notas de fuente, tasa y alcance en la columna `comentario`.

## 7. Cobertura disponible y faltantes conocidos

- P&G julio: disponible para Colombia, EE.UU., Peru ordinario y Costa Rica.
- Detalle contable colombiano por cuenta/tercero/subcategoria: ultimo periodo identificado junio.
- MRR: ultimo periodo identificado mayo.
- ARR: ultimo periodo identificado junio.
- Marketing: ultimo periodo identificado junio.
- Customer Success: ultimo periodo identificado mayo.
- Cartera y liquidez: ultimo periodo identificado junio.
- Pipeline comercial: julio disponible.

No completar estas areas con estimaciones. El dashboard debe indicar el ultimo periodo real disponible.

## 8. Problemas y riesgos conocidos

- La documentacion `README.md` esta desactualizada: habla de nueve pestanas, pero la base actual contiene muchas mas.
- Hay texto con mojibake en HTML y documentacion (`Administraci?n`, caracteres rotos). No cambiar codificacion masivamente sin comparar el comportamiento.
- Existen tres HTML identicos con nombres diferentes, lo cual genera confusion operativa.
- El lanzador abre `Dashboard_COCO_JULIO_2026_R4.html`, aunque `Dashboard_COCO.html` deberia ser el archivo canon.
- El navegador usa `localStorage`; si falla la carga automatica puede mostrar una base cacheada y antigua.
- La tasa de Peru julio es provisional y debe sustituirse cuando se tenga el promedio mensual definitivo.
- Los graficos de gastos dependen de hojas puente diferentes al P&G. Tener subtotales en `BD_PYG_OFICIAL` no garantiza que el detalle por area aparezca.
- No borrar graficos de cascada ni componentes existentes al corregir datos. Separar correcciones de datos de cambios visuales.

## 9. Procedimiento recomendado para una actualizacion mensual

1. Crear un respaldo fechado de `BD_MAESTRA_COCO.xlsx`.
2. Revisar la fuente mensual de cada pais y su moneda.
3. Homologar cuentas sin cambiar la naturaleza contable.
4. Separar proyectos especiales e intercompanias antes de consolidar.
5. Cargar valores por pais en `BD_Indicadores`, `BD_PYG_OFICIAL` y `Subsidiarias_PyG`.
6. Recalcular el bloque `Consolidado` en USD.
7. Actualizar puentes de gastos si existe detalle nuevo.
8. Validar ingresos, utilidad bruta, utilidad operativa y utilidad neta por pais y consolidado.
9. Abrir el dashboard mediante servidor HTTP, cargar la base y revisar todas las pestanas.
10. Confirmar que periodos faltantes no aparezcan como cero.

## 10. Criterios de aceptacion antes de modificar el proyecto

- Trabajar sobre copias o dejar respaldo antes de editar la base oficial.
- No asumir que una hoja con nombre `oficial` es la unica consumida por el dashboard.
- Trazar cada KPI hasta la funcion del HTML y la hoja exacta que lo alimenta.
- Probar desktop y una ventana estrecha despues de cambios visuales.
- Mantener los botones de periodo, filtros, carga local, nube, PDF y navegacion funcionales.
- Informar cualquier inconsistencia de fuente antes de sobrescribir cifras.

## 11. Prompt inicial recomendado para Claude

```text
Trabaja como ingeniero senior y responsable FP&A sobre este proyecto COCO. Lee primero COCO_Dashboard_Cloud/CLAUDE_HANDOFF.md completo y luego inspecciona el codigo y la base; no asumas que la documentacion antigua sigue vigente.

Archivos oficiales:
- COCO_Dashboard_Cloud/Dashboard_COCO.html
- COCO_Dashboard_Cloud/migracion/BD_MAESTRA_COCO.xlsx
- Consolidado Financiero Estructurado COCO 2026.xlsx

Reglas obligatorias:
1. No inventes cifras ni conviertas faltantes en cero.
2. Peru usa PEN y TRM_Peru; su proyecto especial se mantiene separado.
3. Toda correccion de P&G debe sincronizar BD_Indicadores, BD_PYG_OFICIAL, Subsidiarias_PyG y el Consolidado.
4. Crea respaldo antes de editar la base oficial.
5. No elimines graficos o funcionalidades para resolver un problema de datos.
6. Verifica el dashboard servido por HTTP y revisa todas las pestanas afectadas.

Antes de hacer cambios, dime: archivo canon identificado, flujo de datos de la funcionalidad solicitada, fuentes que usaras, cifras o campos faltantes y plan de validacion. Luego implementa, prueba y documenta exactamente lo realizado.
```

## 12. Primeras tareas sugeridas para Claude

1. Verificar visualmente que Peru julio y el consolidado actualizado aparecen al cargar `BD_MAESTRA_COCO.xlsx`.
2. Unificar los tres HTML y hacer que el `.bat` abra `Dashboard_COCO.html`.
3. Actualizar `README.md` con la arquitectura real de la base.
4. Crear pruebas automatizadas de conciliacion entre paises y consolidado.
5. Auditar por pestana que cada grafico distingue `dato faltante` de `cero real`.

