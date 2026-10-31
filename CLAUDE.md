# CLAUDE.md — Mi panel

Contexto del proyecto para Claude Code. Léelo entero antes de tocar nada y mantenlo actualizado: cuando se tome una decisión nueva o cambie una existente, actualiza este fichero en el mismo commit.

## Reglas de seguridad (obligatorias)

Este repositorio es **público**. Nunca añadas al repositorio, ni en código ni en este fichero ni en mensajes de commit:

- Tokens, contraseñas, claves de API o secretos de cualquier tipo.
- Identificadores de la organización del trabajo: Tenant ID, Client ID de App Registrations, nombres de dominio internos, nombres de servidores o de la empresa. Refiérete a ella como "la cuenta del trabajo" o "el tenant del trabajo".
- Datos personales: tareas reales, correos, nombres de personas, direcciones.

Cualquier identificador que la app necesite (token de Todoist, Client ID de Microsoft, Client ID de Google) lo introduce el usuario en la pantalla de Ajustes y se guarda en el navegador, nunca en el código. Si una tarea parece requerir lo contrario, para y pregunta.

## Qué es

Panel personal de tareas, en modo oscuro y muy visual, que centraliza en un solo sitio:

- **Personal**: tareas de Todoist (plan gratuito).
- **Trabajo**: tareas de Microsoft Planner (plan personal del usuario en el tenant del trabajo). *Pendiente.*
- **Calendario**: citas de Google Calendar. *Pendiente.*

Encima de las tareas de origen, la app añade su propia capa: **etiquetas con nombre y color**, filtros y una vista de tablero cómoda. Se usa desde el PC del trabajo, el PC de casa y un iPad, siempre por navegador.

## Arquitectura

- **Web estática, sin backend y sin paso de compilación.** HTML + CSS + JavaScript plano. No introducir frameworks, bundlers, npm ni TypeScript salvo que el usuario lo pida expresamente.
- **Alojamiento:** GitHub Pages, desplegando desde la rama `main`, carpeta raíz. Cada push a `main` publica.
- **Ficheros:**
  - `index.html`: estructura estática (barra lateral, cabecera, contenedores).
  - `styles.css`: todo el diseño. Variables de color en `:root`.
  - `app.js`: toda la lógica, en un IIFE. Sin módulos ES.
  - `Sortable.min.js`: SortableJS 1.15.2 (MIT), incluido en el repo a propósito para no depender de CDN. No sustituir por un enlace a CDN.
  - `favicon.svg`: icono de la pestaña. `apple-touch-icon.png` (180 px): icono de pantalla de inicio en iPad/iPhone, generado a partir del mismo dibujo.
- **El navegador habla directamente con las APIs** (Todoist admite CORS desde cualquier origen). No hay servidor intermedio.

### Almacenamiento en el navegador (localStorage)

| Clave | Contenido |
|---|---|
| `mp.todoist.token` | Token personal de Todoist |
| `mp.todoist.cache` | Última copia de proyectos, secciones, tareas y comentarios |
| `mp.tags` | `{ list: [{id, name, color}], assign: { [taskId]: [tagId, ...] } }` |
| `mp.ui` | Preferencias de interfaz (`sidebarHidden`, `projectId`) |

Las etiquetas viven solo en el navegador hasta la fase de OneDrive. En Ajustes hay Exportar/Importar en JSON como puente.

## Integración con Todoist (hecha)

- API v1, endpoint único `https://api.todoist.com/api/v1/sync`, autenticación `Authorization: Bearer <token>`, cuerpo `application/x-www-form-urlencoded`.
- Lectura: `sync_token='*'` con `resource_types` `["user","projects","sections","items","notes"]`. Se descartan tareas completadas, borradas y archivadas, y los comentarios borrados o de tareas no activas.
- Escritura mediante `commands` (cada uno con `uuid` propio; `temp_id` al crear): `item_add`, `item_update` (renombrar), `item_move` (cambiar de sección; a "Sin sección" se mueve con `project_id`), `item_reorder`, `item_close` (completar; respeta tareas recurrentes), `item_uncomplete` (deshacer), `note_add` / `note_update` / `note_delete` (comentarios).
- Se comprueba `sync_status` de cada comando; si algo falla se muestra aviso y se vuelve a leer todo de Todoist.
- Plan gratuito de Todoist: 5 proyectos, 20 secciones, 300 tareas activas. La API es gratuita.

## Decisiones de producto (respetarlas)

- **Sincronización manual.** Solo se conecta con las fuentes cuando el usuario pulsa Sincronizar o hace un cambio (mover, crear, renombrar, completar). Al abrir la app se muestra la última copia guardada. No añadir sincronización periódica ni al arrancar.
- **Tres pantallas separadas: Trabajo, Personal y Calendario.** No mezclar Planner y Todoist en un mismo tablero.
- **Vista de tablero** con las columnas de la fuente: secciones de Todoist, buckets de Planner. "Sin sección" solo aparece si tiene tareas o si no hay secciones.
- **Las columnas son de solo lectura en la app.** Crear, renombrar, reordenar o borrar columnas se hace en el origen.
- **Mover tarjetas entre columnas es crucial** y sí se escribe en el origen.
- **Tarjetas compactas:** casilla de completar + nombre en una línea (con elipsis). Si tiene etiquetas, van en una segunda fila, en pequeño. Borde lateral del color de la primera etiqueta.
- **Etiquetas:** gestor propio (nombre + color, sin iconos, ordenables). Varias por tarea. Son solo de la app, no se escriben en Todoist. Se comparten entre las tres pantallas.
- **Asignar etiqueta:** arrastrándola desde la bandeja superior hasta la tarjeta. Pulsar una etiqueta de la bandeja filtra el tablero.
- **Quitar etiquetas:** desde el panel de la tarea, no arrastrando.
- **Añadir/editar tarea:** panel sencillo con nombre y etiquetas. "Añadir tarea" al pie de cada columna crea en esa columna; el botón de la cabecera permite elegir columna.
- **Menú lateral estrecho y ocultable** (preferencia recordada por navegador): igual de ancho que el logo (116 px), con el logo arriba, cada opción como icono con su nombre en pequeño debajo, sin contadores, la última sincronización compacta abajo y el botón de ocultar al pie. En pantallas medianas y pequeñas pasa a ser un panel flotante.
- **Adaptable:** columnas que se reparten el ancho (mínimo 300 px), desplazamiento horizontal si no caben; en pantallas pequeñas una columna por pantalla con encaje al desplazar. En iPad, arrastrar tarjetas requiere mantener pulsado un instante.
- **Comentarios de Todoist:** se ven en el panel de la tarea (del más antiguo al más reciente, texto plano con saltos de línea; los adjuntos como enlace). Se pueden añadir, editar y borrar, y se guardan en Todoist al momento, aparte del botón Guardar. Si una tarea tiene comentarios, su tarjeta muestra un icono de bocadillo a la derecha del nombre (sin número).
- La descripción de las tareas de Todoist no se muestra ni se edita.
- Las subtareas de Todoist no se muestran por ahora.

## Estilo visual

Oscuro y elegante. Tipografía Manrope (Google Fonts) con alternativas del sistema.

- Fondo `#0B0F15`, barra lateral `#0F141B`, columnas `#121821`, tarjetas `#18202B`, paneles `#10161E`.
- Texto `#E7EAF0`, secundario `#C9D0DC`, apagado `#9AA4B5` / `#8A94A6`.
- Acento principal lima `#DBF266` (botón primario, marca), azul `#7FA2E0` (selección, foco), rojo suave `#F0857A` (errores, borrar).
- Esquinas redondeadas (9–18 px), bordes de 1 px semitransparentes, sin degradados ni sombras exageradas.
- Iconos de línea en SVG, sin emojis. Toda la interfaz en español.
- **Marca:** un vegvísir dibujado en línea, en oscuro `#0B0F15` sobre cuadrado lima redondeado. Es el mismo dibujo detallado en la barra lateral (SVG en línea en `index.html`), en `favicon.svg` y en `apple-touch-icon.png`; se acepta que pierda definición a tamaño pequeño. Si se cambia, cambiar los tres.

## Convenciones de código

- Crear el DOM con el helper `h()` de `app.js`. **Nunca usar `innerHTML` con datos de las APIs ni del usuario.**
- Validar colores con `isHex()` antes de usarlos en estilos.
- Cambios optimistas en pantalla y después la llamada a la API; si falla, avisar con `toast()` y re-sincronizar.
- Mantener los textos de la interfaz en español y con tono directo.
- Código legible y comentado en español solo donde aporte.
- **Versión anti-caché:** en `index.html`, `styles.css`, `app.js`, `favicon.svg` y `apple-touch-icon.png` se cargan con `?v=N` (mismo número en todos). En cada cambio que se suba, incrementar `N` en uno en todos, aunque solo haya cambiado uno de los ficheros. `Sortable.min.js` no lleva versión porque no cambia.

## Forma de trabajar con el usuario

- El usuario prefiere **revisar un plan antes de que se escriba código**: para cambios no triviales, explica primero qué vas a hacer y espera su visto bueno.
- Entrega cambios **completos y probados**, no a medias ni por ensayo y error.
- Antes de hacer push, comprueba la sintaxis (`node --check app.js`) y, si es posible, prueba la app en local (`python3 -m http.server`).
- Mensajes de commit en español, breves y descriptivos.
- Respuestas en español, directas y prácticas.

## Próximas fases

1. **OneDrive personal para las etiquetas** (sincronizar entre dispositivos).
   - Login con MSAL.js (`@azure/msal-browser`, versión fija, incluida en el repo como Sortable) contra la autoridad `https://login.microsoftonline.com/consumers` para que solo acepte cuentas personales.
   - Permisos delegados: `Files.ReadWrite.AppFolder`, `User.Read`, `offline_access`.
   - Fichero de datos en la carpeta de la app (`/me/drive/special/approot`), leído y escrito por Graph con control de versiones mediante eTag (`If-Match`) y fusión si hay conflicto.
   - Migrar automáticamente las etiquetas locales existentes en el primer login.
   - La App Registration (tipo SPA, cuentas de cualquier organización y personales) la crea el usuario; su Client ID se introduce en Ajustes.
2. **Trabajo: Microsoft Planner.** Misma App Registration. Login contra el tenant del trabajo (el Tenant ID lo introduce el usuario en Ajustes). Permiso delegado `Tasks.ReadWrite`. Mismo tablero y mismas etiquetas que Personal, con los buckets como columnas. El login solo se pide al sincronizar.
3. **Calendario: Google Calendar.** Pantalla propia con vista de calendario (por definir). Login con Google Identity Services desde el navegador; el Client ID de Google se introduce en Ajustes.
4. Más adelante: elementos propios de la app (pagos recurrentes, cumpleaños, recordatorios) y captura rápida de tareas.
