# Poner el dashboard en línea — guía paso a paso

> **Este plan nunca se activó.** El tablero real está publicado en **Cloudflare**
> (`coco-cockpit.cocotec.workers.dev`), no en Google. Esta guía describe un camino
> alterno que se dejó preparado y quedó en pausa — no la sigas pensando que es
> cómo está desplegado hoy. Para el despliegue real, ver `README.md`, secciones 1 y 9.

Tiempo estimado: **20 minutos**. Todo es gratis (Google Workspace que ya tienes).

Al terminar vas a tener:
- La base en un **Google Sheet** (fuente única, con historial de versiones)
- Un **enlace** que entrega los datos al tablero
- Cada área cargando **su propia porción** de datos, con validación y bitácora
- Solo correos **@cocotech.ai** y **@cfocus.co** pueden escribir

---

## Antes de empezar

Ten a mano estos **tres** archivos de la carpeta del proyecto:

| Archivo | Dónde está | Qué se hace con él |
|---|---|---|
| `BD_MAESTRA_COCO.xlsx` | `migracion\` | **Se sube** a Google Drive |
| `Code.gs` | `apps_script\` | Se **copia y pega** en Apps Script |
| `Dashboard.html` | `apps_script\` | Se **copia y pega** en Apps Script |

> Si cambiaste la base, regenera el primero con `migracion\preparar_google_sheets.py`.
> Si cambió el tablero o el informe de caja, regenera el tercero con
> `apps_script\generar_Dashboard_html.py`.

### Cómo queda al final

El tablero vive **dentro de Google**, en la misma dirección que entrega los datos.
Todos entran por **una sola URL** y siempre ven la última versión. Google verifica
quién entra: **no hay claves que repartir ni que rotar**.

---

## Paso 1 — Subir la base a Google Sheets (5 min)

1. Entra a **drive.google.com** con tu correo de COCO.
2. Crea una carpeta llamada **COCO Dashboard** (clic derecho → Nueva carpeta).
3. Arrastra ahí el archivo `BD_MAESTRA_COCO.xlsx`.
4. **Doble clic** en el archivo subido → se abre en modo vista previa.
5. Arriba dice *"Abrir con Hojas de cálculo de Google"* → haz clic.
6. En el menú: **Archivo → Guardar como hoja de cálculo de Google**.
7. Se crea un archivo nuevo. Renómbralo a **`BD_MAESTRA_COCO`** (sin el .xlsx).
8. Ya puedes borrar el .xlsx que subiste: la fuente de verdad es el Sheet.

**Verifica:** el Sheet debe tener **19 pestañas** abajo, entre ellas `BD_Indicadores`,
`CATALOGOS`, `LOG` y `Pipeline_Comercial`.

---

## Paso 2 — Instalar el script y el tablero (7 min)

1. Con el Sheet abierto: **Extensiones → Apps Script**.
2. Se abre un editor. Borra todo lo que diga (suele haber un `function myFunction()`).
3. Abre `apps_script\Code.gs` en tu computador, **copia todo** y pégalo ahí.
4. Guarda con **Ctrl + S**. Ponle nombre al proyecto: *COCO Dashboard*.

### Agrega el tablero

5. En el panel izquierdo, junto a **Archivos**, clic en **＋ → HTML**.
6. Nómbralo exactamente **`Dashboard`** (Google le pone `.html` solo).
7. Borra el contenido de ejemplo.
8. Abre `apps_script\Dashboard.html` con el **Bloc de notas**, selecciona todo
   (**Ctrl+E**), copia (**Ctrl+C**) y pégalo ahí. Guarda con **Ctrl+S**.

   > Es un archivo grande (≈880 KB). Puede tardar unos segundos en pegar.

### Autoriza

9. Arriba, elige la función **`probarLectura`** y presiona **Ejecutar**.
10. La primera vez Google pide permisos:
    - *Revisar permisos* → elige tu cuenta
    - Aparece *"Google no ha verificado esta aplicación"* → **Configuración avanzada**
      → *"Ir a COCO Dashboard (no seguro)"* → **Permitir**

    > Ese aviso es normal: la aplicación eres tú mismo, no un tercero.

11. Abajo, en *Registro de ejecución*, debe listar las hojas con su número de filas.
    Si dice `BD_Indicadores: 1615 filas` (o parecido), vas bien.

---

## Paso 3 — Publicar (3 min)

1. Arriba a la derecha: **Implementar → Nueva implementación**.
2. Clic en el engranaje ⚙ junto a *"Seleccionar tipo"* → **Aplicación web**.
3. Configura **exactamente así** — de esto depende la seguridad:

   | Campo | Valor | Por qué |
   |---|---|---|
   | Descripción | `v1` | |
   | Ejecutar como | **Yo (tu correo)** | Nadie necesita permisos sobre el Sheet, así que nadie puede saltarse la validación editándolo a mano |
   | Quién tiene acceso | **Usuarios de cocotech.ai** | Google bloquea a cualquiera fuera del dominio, antes de que el código corra |

4. **Implementar** → copia la **URL de la aplicación web** (termina en `/exec`).

**Esa URL es el tablero.** Ábrela en el navegador: debe cargar con los datos del Sheet
y mostrar tu correo junto al nombre de la base.

---

## Paso 4 — Repartir

A cada persona solo le mandas **la URL**. Nada más.

- No hay archivo que instalar
- No hay clave que memorizar
- Entra con su cuenta @cocotech.ai y el tablero ya sabe quién es

Entrégale además su plantilla de `plantillas\` según el área.

> **No compartas el Sheet.** Nadie necesita acceso directo: el tablero hace todo. Si
> alguien edita el Sheet a mano, se salta la validación y no queda rastro en la bitácora.

### Cómo carga datos un área

- **Un dato suelto**: ✎ Ingresar datos → *Registro individual*
- **Varios datos**: ✎ Ingresar datos → *Carga masiva* → descarga su plantilla, la llena,
  la sube. Ve una **previsualización** con los errores fila por fila; nada se escribe
  hasta que confirme.

Marketing no puede escribir códigos de Finanzas: cada área está limitada a los suyos,
y en la bitácora queda su correo real (verificado por Google, no declarado).

---

## Cuando cambie el tablero o el informe de caja

1. En tu computador: `apps_script\generar_Dashboard_html.py`
2. Copia el nuevo `Dashboard.html` y pégalo en Apps Script (reemplazando)
3. **Implementar → Administrar implementaciones → editar (lápiz) → Versión: Nueva → Implementar**

> El paso 3 es el que más se olvida: sin él, la URL sigue mostrando la versión anterior.

### Cómo carga datos un área

- **Un dato suelto**: ✎ Ingresar datos → *Registro individual*
- **Varios datos**: ✎ Ingresar datos → *Carga masiva* → descarga su plantilla, la llena,
  la sube. Ve una **previsualización** con los errores fila por fila; nada se escribe
  hasta que confirme.

Marketing no puede escribir códigos de Finanzas: cada área está limitada a los suyos.

---

## Cada mes: actualizar el modelo financiero

Esto sigue siendo local (Python no corre en la nube, y así **tú controlas cuándo se
toma la foto del modelo**):

1. Reemplaza `modelo\Modelo_COCO.xlsx` por la versión nueva
2. Doble clic en `modelo\Actualizar_Burn_Runway.bat` → actualiza el informe de caja
3. Si activaste líneas en `mapa_modelo.csv`: doble clic en `modelo\Cargar_del_Modelo.bat`

---

## Si algo falla

| Síntoma | Causa probable | Solución |
|---|---|---|
| Al abrir la URL sale "no tienes acceso" | La persona no es de @cocotech.ai | Debe entrar con su cuenta de COCO, no con Gmail personal |
| "No pude verificar tu identidad" | Sesión de Google confusa (varias cuentas) | Abre la URL en ventana de incógnito con la cuenta de COCO |
| "El dominio X no tiene permiso" | Correo fuera de @cocotech.ai / @cfocus.co | Usa el correo corporativo |
| Cambié algo y no se ve | Falta publicar la versión nueva | **Implementar → Administrar implementaciones → editar (lápiz) → Versión: Nueva → Implementar** |
| Sale la pantalla de "arrastra la base" | El script no pudo leer el Sheet | Ejecuta `probarLectura` en Apps Script y mira el error |

---

## Respaldo

- El Sheet guarda **historial de versiones**: *Archivo → Historial de versiones*.
  Un error de carga se deshace desde ahí.
- Una vez al mes: *Archivo → Descargar → .xlsx* y guárdalo en
  `OneDrive\…\COCO\1.0 Coco Digital\Respaldos_BD\AAAA-MM.xlsx`.
  Ese archivo funciona con el botón **Cargar BD** del tablero — plan B completo sin nube.

---

## Lo que esta solución NO hace

1. **Se agrega, no se corrige.** Un dato mal cargado se arregla en el Sheet directamente
   (y queda en el historial de versiones de Google). El tablero solo suma filas.
2. **El CFO externo (@cfocus.co) queda fuera del acceso por dominio.** Google solo deja
   entrar a @cocotech.ai. Dos salidas: darle una cuenta @cocotech.ai (lo recomendable),
   o que use el tablero local con la clave — ese camino sigue funcionando.
3. **Los motores del modelo siguen siendo locales.** Actualizar el modelo financiero y
   el informe de caja se hace en tu computador con los `.bat`, y luego se republica el
   tablero. Es a propósito: **tú decides cuándo se toma la foto del modelo**.
4. **Google muestra un aviso de "app no verificada"** la primera vez. Es normal en
   proyectos internos; solo lo ve quien instala, no los usuarios.
