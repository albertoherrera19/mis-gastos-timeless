# Resumen de estado — actualizado 2026-07-17

> Este archivo describe el ESTADO ACTUAL (cambia con el tiempo). Las reglas fijas del proyecto están
> en `CLAUDE.md` — léelo primero.

## Punto de sincronización entre los dos repos
**Sincronizado el 2026-07-17** (commit amigos `5d88df1`). Se portaron a `mis-gastos`
(carpeta `mis-gastos-timeless`) todas las funcionalidades pendientes desde el sync anterior
(`300eccb` en personal / `65c7596` en amigos), es decir el rango `300eccb..099dc22`:

- Presupuesto general y por grupo (además del ya existente por categoría), botón 🎯 junto al total.
- Gastos recurrentes de monto variable (pelo/skincare/suplementos) dentro de la página de 🔁, en su
  propia sub-pestaña. Renombre: la funcionalidad vieja de suscripciones/monto fijo pasa a llamarse
  **"Gastos fijos"**, y la nueva de monto variable pasa a llamarse **"Gastos recurrentes"**.
- Gastos etiquetables a **varios grupos a la vez** (ej. un gasto en "Personal" y "Gym" simultáneamente)
  sin duplicar el total general/Predeterminado — tanto al crear como al editar un gasto.
- Orden de la lista "Días con gasto" (detalle de categoría): más antiguo (por defecto), más reciente,
  monto menor a mayor, monto mayor a menor. No afecta el orden del gráfico de barras por día.
- **Editar y borrar cualquier categoría, incluidas las 8 base** (antes solo se podían borrar las
  personalizadas). Pensado para que amigos pueda quitar/renombrar categorías que no le aplican (ej.
  Inversión, Ads). Los cambios se guardan como overrides/exclusiones locales en localStorage
  (`timeless_cat_overrides`, `timeless_deleted_base_cats`) sin tocar el arreglo `BASE_CATEGORIES` del
  código — cada dispositivo/repo puede tener su propio set. "Otros" queda protegida como destino de
  respaldo (no se puede eliminar).
- **Íconos actualizados de la PWA** también para el repo de amigos esta vez (a pedido explícito del
  usuario) — antes eran exclusivos del repo personal. Set completo `icons/` (any 192/512, maskable
  192/512, apple-touch-icon-180, favicon 32/16), `manifest.json` y `service-worker.js` actualizados.

Quedaron intencionalmente FUERA del repo de amigos (como siempre): la integración con Google Sheets
completa (constantes, funciones, botón, toast) y la pre-creación de grupos (en amigos, `catGroups`
sigue empezando **vacío** — solo aparecen las pestañas "Predeterminado" y "+ Grupo", sin
"Timeless"/"Personal"). Verificado en navegador sin errores de consola antes del push (cache del
service worker de amigos subida a `v21`).

Todo cambio FUTURO en el repo personal (después de este punto) vuelve a quedar pendiente de portar a
amigos hasta el próximo pedido explícito de sincronización.

## ✅ Funcionalidades implementadas y funcionando (repo personal, verificadas en navegador)
- Registro de gastos: monto, nota, categoría, **fecha manual opcional** (por defecto hoy, no permite
  futuro), **etiqueta de grupo opcional** (uno o varios) por gasto individual, editable después.
- 8 categorías base + categorías personalizadas (emoji + nombre). **Cualquier categoría (base o
  personalizada) se puede editar (nombre/ícono) y borrar**, con confirmación y oferta de mover sus
  gastos a "Otros" si tenía. "Otros" no se puede borrar (es el destino de respaldo).
- Selector de categoría con **toggle** (tocar de nuevo deselecciona).
- 13 temas de color, con la regla de acento neutro (Negro/Blanco) y color independiente por categoría.
- Donut por categoría (interactivo, click muestra %), con monto central compacto y escalado por dígitos.
- Comparativo de barras por mes (Comparar meses) — verificado que julio se conserva como historial al
  simular el paso a agosto, sin mezclar totales.
- **Detalle de categoría** (página completa): gráfico de barras por día con zoom/pinch, posicionado para
  que la barra del **primer día con gasto** quede pegada al borde izquierdo (sin scroll manual, sea el
  día 1, el 9 o el 30); lista "Días con gasto" expandible con montos individuales y **selector de orden**
  (más antiguo/reciente/monto asc/desc); **comparativo ▲/▼ % vs mes anterior** (oculto si no hay datos
  previos); **presupuesto mensual opcional** con barra verde/amarillo/rojo; engranaje ⚙️ coloreado que
  abre color + presupuesto juntos.
- **Presupuesto general y por grupo**: botón 🎯 junto al total del mes, contextual (cambia según la
  pestaña de grupo activa), misma barra verde/amarillo/rojo que el de categoría.
- **Editar gastos** en página completa (lápiz, desde Movimientos y desde Días con gasto): edita monto,
  nota, categoría, fecha y grupo(s); actualiza el mismo registro (no duplica).
- **Movimientos recientes**: SIN límite de cantidad (verificado con 1000 gastos, todos visibles, sumas
  exactas); orden "Predeterminado" (por fecha, más reciente arriba), "Más antiguo" (fecha ascendente),
  "Por categoría" (agrupado, expandible); **buscador y filtros** (texto de nota/categoría, rango de
  monto, rango de fechas), compatible con los 3 modos de orden; confirmación antes de borrar un gasto.
- **Grupos de categorías**: pestañas Predeterminado + grupos personalizados (crear/editar/eliminar,
  nombre + categorías, una categoría puede estar en varios grupos); filtra total del mes, donut y "Por
  categoría"; además, un gasto individual puede etiquetarse a **uno o varios** grupos aunque su categoría
  no pertenezca a ellos, sin duplicar el total general. Barra de grupos ubicada entre "Mis gastos" y el
  mes actual, en el header.
- **Gastos fijos y recurrentes**: vista de página completa (ícono 🔁 en el header), con dos sub-pestañas:
  - *Gastos fijos*: nombre/monto/día/categoría, toggle Pagado/Pendiente por mes; al marcar pagado,
    ofrece registrar el gasto real (o no); al desmarcar, ofrece quitar el gasto vinculado.
  - *Gastos recurrentes*: nombre/categoría, SIN monto fijo (varía cada vez); marcar "Hecho" lleva al
    formulario de agregar gasto pre-llenado con la categoría, para escribir el monto real de esa vez;
    banner cancelable mientras está pendiente de completar.
- **Google Sheets** (solo repo personal): cada gasto se envía automáticamente al registrarse
  (`queueForSheets`), con cola de reintento si no hay internet (`timeless_sheets_pending`, se reintenta
  en el evento `online`). Además hay un **botón 💾 manual** en el header que fuerza el reenvío y muestra
  un toast: "✓ Guardado en Google Sheets" (éxito), "Sin conexión, intenta de nuevo" (offline).
- **Respaldo de datos**: exportar (descarga `.json` con `BACKUP_KEYS`) e importar (reemplaza localStorage
  y recarga), en el panel ⚙️. Usado exitosamente por el usuario para migrar datos de un ícono viejo de
  iOS a la instalación nueva.
- **Auto-actualización de la PWA**: el service worker revisa actualizaciones en `load`, `visibilitychange`
  y `pageshow` (no solo la primera vez), para minimizar el problema de iconos instalados que quedan
  pegados en versiones viejas de iOS.
- **Íconos nuevos**: set completo en `icons/` (any 192/512, maskable 192/512, apple-touch-icon-180,
  favicon 32/16), rutas relativas, verificado con HTTP 200 en vivo en el subpath. Ahora también en el
  repo de amigos (ver "Punto de sincronización" arriba).

## ⏳ Pendiente / a medio hacer
- **Sincronizar al repo de amigos** todo cambio futuro del repo personal — el usuario aún no lo ha
  pedido explícitamente. No hacerlo hasta que lo pida.
- No hay ninguna funcionalidad a medias dentro del repo personal en este momento — todo lo commiteado
  fue probado en navegador sin errores de consola.

## 🐛 Bugs conocidos
- Ninguno abierto actualmente. Notas de comportamiento (no bugs, pero conviene recordarlos):
  - El envío a Google Sheets usa `fetch(..., {mode:'no-cors'})` → es "fire-and-forget", no se puede leer
    la respuesta real del servidor desde el navegador (limitación intencional para evitar bloqueos CORS).
    El botón 💾 asume éxito si el fetch no lanza error de red.
  - Los pushes a GitHub a veces fallan una vez con "Deployment failed, try again later" (infraestructura
    de GitHub Pages, no del código) — se resuelve reintentando con un commit vacío.
  - En iOS, un ícono instalado antes de un cambio de código puede tardar hasta 24h en autoactualizarse
    aunque el fix de `visibilitychange` ya esté desplegado, porque ese fix necesita llegar primero. Un
    ícono MUY viejo puede necesitar una recarga manual una sola vez vía Safari normal para "despegarse".
  - Si un grupo personalizado tiene una categoría en su lista (`cats`) y esa categoría se borra después,
    el id queda huérfano dentro del grupo (no se limpia automáticamente). No causa errores porque todos
    los lugares que resuelven una categoría por id ya tienen fallback (`catById(id) || {icon:'🗂️',
    name:'Otros'}`), pero si se nota alguna inconsistencia visual en un grupo tras borrar una categoría,
    esta es la causa más probable.

## Decisiones de diseño de esta sesión (no cubiertas en CLAUDE.md)
- La edición/borrado de categorías base NO modifica `BASE_CATEGORIES` en el código (es compartido por
  ambos repos) — en su lugar usa dos claves de localStorage por dispositivo: `timeless_cat_overrides`
  (nombre/ícono editado) y `timeless_deleted_base_cats` (ids ocultos). `allCategories()` combina ambas
  sin tocar el arreglo original. Esto permite que personal y amigos tengan catálogos de categorías
  distintos usando el mismo código fuente.
- El formulario "Nueva categoría" (`#newCatForm`) se reutilizó también para **editar** (categoría base o
  personalizada), alternando el texto del botón entre "Crear categoría" y "Guardar cambios" según
  `catFormEditId`, en vez de crear un formulario separado.
- El ícono de editar (✎) se agregó en la esquina opuesta al de borrar (✕) sobre cada `.cat-btn`, mismo
  patrón visual (círculo pequeño flotante), para no rediseñar la grilla de categorías.
- El engranaje de categoría es un SVG inline (no emoji ni imagen), coloreado vía `style.color` con el
  hex de la categoría, para que combine visualmente sin depender de assets externos.
- El toast (`#toast`) es un elemento único reutilizado por `showToast(msg, kind)`, con clases `.ok`/`.err`
  para el borde de color; se autolimpia con `setTimeout` guardado en `t._timer` (cancela el anterior si
  se dispara de nuevo rápido).
- La carpeta de íconos fuente que trajo el usuario (`Downloads/app-icons/`) se copió a `icons/` dentro
  del repo y luego se borró el original — si en el futuro se pide cambiar íconos de nuevo, ya no existe
  esa carpeta fuente.
- El README.md de instrucciones de íconos traía rutas absolutas `/icons/...`; se decidió usarlas
  **relativas** (`icons/...`) porque el proyecto entero es 100% rutas relativas y se sirve en subpath.
- Para portar los cambios al repo de amigos se usó `git merge-file` (3-way merge real, con la base común
  siendo el commit del último sync) en vez de aplicar un patch plano — esto reconcilia automáticamente
  la mayoría de diferencias intencionales (Sheets removido, `PRECREATE_GROUPS=false`) y solo deja
  conflictos genuinos para resolver a mano (ej. el título del botón 🔁 y el orden de `BACKUP_KEYS`).

## Preguntas sin responder
- Ninguna pendiente en este momento. La única decisión abierta es **cuándo** el usuario quiere
  sincronizar los cambios acumulados al repo de amigos — queda a su criterio, sin fecha definida.
