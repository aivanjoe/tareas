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
- **Trabajo**: tareas de Microsoft Planner (el plan "Tareas mías" del usuario en el tenant del trabajo).
- **Calendario**: el calendario de Google elegido por el usuario, en vista de mes, donde la app crea además sus propios eventos.

Encima de las tareas de origen, la app añade su propia capa: **etiquetas con nombre y color**, filtros y una vista de tablero cómoda. Se usa desde el PC del trabajo, el PC de casa y un iPad, siempre por navegador.

## Arquitectura

- **Web estática, sin backend y sin paso de compilación.** HTML + CSS + JavaScript plano. No introducir frameworks, bundlers, npm ni TypeScript salvo que el usuario lo pida expresamente.
- **Alojamiento:** GitHub Pages, desplegando desde la rama `main`, carpeta raíz. Cada push a `main` publica.
- **Ficheros:**
  - `index.html`: estructura estática (barra lateral, cabecera, contenedores).
  - `styles.css`: todo el diseño. Variables de color en `:root`.
  - `app.js`: toda la lógica, en un IIFE. Sin módulos ES.
  - `Sortable.min.js`: SortableJS 1.15.2 (MIT), incluido en el repo a propósito para no depender de CDN. No sustituir por un enlace a CDN.
  - `msal-browser.min.js`: MSAL.js (`@azure/msal-browser`) 4.30.0 (MIT), incluido en el repo como Sortable. Se carga bajo demanda solo si el usuario ha puesto el Client ID de OneDrive. Rama 4 (LTS).
  - `chart.umd.min.js`: Chart.js 4.5.1 (MIT). Se carga bajo demanda al pintar un gráfico.
  - `pdf.min.mjs` y `pdf.worker.min.mjs`: pdf.js 6.4.299 (Apache-2.0), compilación *legacy* para navegadores menos recientes (iPad). Se cargan con `import()` dinámico solo al leer facturas, con `isEvalSupported: false`. Las librerías no llevan versión anti-caché porque no cambian; si se actualizan, cambiar el nombre del fichero.
  - `favicon.svg`: icono de la pestaña. `apple-touch-icon.png` (180 px): icono de pantalla de inicio en iPad/iPhone, generado a partir del mismo dibujo.
- **El navegador habla directamente con las APIs** (Todoist admite CORS desde cualquier origen). No hay servidor intermedio.

### Almacenamiento en el navegador (localStorage)

| Clave | Contenido |
|---|---|
| `mp.todoist.token` | Token personal de Todoist |
| `mp.todoist.cache` | Última copia de proyectos, secciones, tareas y comentarios |
| `mp.datos` | Almacén de datos propios de la app (ver abajo) |
| `mp.datos.base` | Última versión sincronizada con OneDrive (base de la fusión) |
| `mp.onedrive` | `{ clientId, account, etag, syncedAt }` |
| `mp.planner` | `{ account, plans, planId, buckets, tasks, cats: { [planId]: {etag, names} }, syncedAt }` (copia de Planner) |
| `mp.google` | `{ clientId, account, token, exp, syncedAt, write, cal }` (token de acceso de Google, dura una hora; `write`/`cal`: si el permiso concedido permite editar contactos / usar Calendar) |
| `mp.calendar` | `{ events, from, to, syncedAt }` (copia de los eventos leídos del calendario para la vista de mes) |
| `msal.*` | Caché de sesión de MSAL (la gestiona la librería) |
| `mp.ui` | Preferencias de interfaz (`sidebarHidden`, `projectId`) |

`mp.tags` es el formato antiguo de las etiquetas: al arrancar, si no existe `mp.datos`, se migra solo. En Ajustes, Exportar descarga el almacén completo (`{ app, version: 2, datos }`); Importar sustituye solo las colecciones que trae el fichero (así se importan datos de otras fuentes sin tocar el resto) y acepta también el formato antiguo de etiquetas.

## Almacén de datos y OneDrive (hecho)

- **Almacén general:** `{ schema, colecciones: { [nombre]: { [id]: registro } } }`. Todo dato propio de la app (etiquetas hoy; gastos, pagos, consumos, cumpledías… después) es una colección de registros con `id`. `SCHEMA` en `app.js` es la versión del formato: si cambia, migrar al leer, y nunca escribir sobre datos de un esquema más nuevo.
- **Etiquetas en el almacén:** `etiquetas` (`{id, name, color}`), `asignaciones` (`{id: taskId, tags: [...]}`) y el orden aparte en `orden` (`{id: 'etiquetas', ids: [...]}`). Regla general: **un registro solo cambia cuando el usuario cambia ese registro** (por eso el orden no va dentro de cada etiqueta).
- **Fusión a tres bandas por registro** (base, local, remoto): si solo cambió un lado gana ese cambio, borrados incluidos; si cambiaron los dos, gana este dispositivo.
- **OneDrive:** cuenta personal (autoridad `consumers`), permisos `Files.ReadWrite.AppFolder` y `User.Read`. Fichero `datos.json` en la carpeta de la app (`/me/drive/special/approot`). Lectura con los metadatos (eTag) y la `downloadUrl`; escritura con `If-Match` (o `If-None-Match: *` si no existe). Si responde 412, se vuelve a leer, se fusiona y se reintenta.
- **Cuándo se sincroniza:** al pulsar Sincronizar (junto con Todoist) y, si está conectado, 1,5 s después de cada cambio en los datos de la app. Nunca al arrancar.
- **Login por redirección, sin ventana emergente:** la página de login de Microsoft aísla su ventana de la que la abre y el sondeo de MSAL 4 deja de funcionar (da `interaction_in_progress` en el siguiente intento). Se usa `loginRedirect` (Conectar, Reconectar, Sincronizar con la sesión caducada, o el botón del aviso) con `prompt: 'select_account'`; al volver, `handleRedirectPromise()` en `loadMsal()` lo procesa. Antes de cada login se borra la marca `msal.interaction.status` que deje un intento anterior. URI de redirección: la propia página (`https://aivanjoe.github.io/tareas/` y `http://localhost:8000/`).
- **Estado:** la cabecera muestra "Sincronizado hace…" de la fuente de la pantalla activa (Todoist, Planner, Google u OneDrive) y, si una conexión necesita atención, un aviso pulsable ("OneDrive: reconectar", "Planner: error"…), también en el móvil. La barra lateral ya no muestra estados de sincronización. En Ajustes, sección OneDrive: Client ID, Conectar, cuenta, Sincronizar ahora y Desconectar (desconectar no borra datos).
- **Pruebas:** sin cuenta real, se prueba en local con MSAL y Graph simulados (dos "dispositivos" contra el mismo OneDrive simulado, con conflictos de eTag).

## Trabajo: Microsoft Planner (hecho)

- **Misma App Registration y misma MSAL que OneDrive**, con la cuenta del trabajo: `loginRedirect` con autoridad `organizations`, `state: 'planner'` y permisos `Tasks.ReadWrite` y `User.Read` (en Entra: permiso delegado `Tasks.ReadWrite` con consentimiento de administrador). En la caché de MSAL conviven las dos cuentas: la personal es la del directorio de cuentas de Microsoft (`tenantId` `9188040d-…`, `consumers`); la otra es la del trabajo. Al volver del login, `handleRedirectPromise` decide por `state` y por el directorio a qué conexión corresponde. Token en silencio con la autoridad del directorio de la cuenta.
- **Lectura (al sincronizar):** `/me/planner/plans` (se elige por defecto el plan "Tareas mías"; selector de plan en la cabecera), `/planner/plans/{id}/buckets` (columnas, por `orderHint` **descendente**, que es como las muestra Planner; comprobado con el plan real), `/planner/plans/{id}/tasks` (sin las completadas; por `orderHint` ascendente) y `/planner/plans/{id}/details` (nombres de las etiquetas).
- **Escritura** con el `etag` de cada objeto (`If-Match`; en PATCH, `Prefer: return=representation` para recibir el etag nuevo): mover de columna y reordenar (`bucketId` + `orderHint` con el formato de Planner `"<anterior> <siguiente>!"`), crear (al final de la columna), renombrar, completar (`percentComplete: 100`, con deshacer a 0), etiquetas de la tarea (`appliedCategories` con `true`/`false`), **notas** (`/planner/tasks/{id}/details` → `description`; se leen al abrir el panel y se guardan con Guardar) y nombres de las etiquetas del plan (`categoryDescriptions`). Si Planner responde 412 (cambiado entretanto) se avisa y se vuelve a leer.
- **Etiquetas de Trabajo = etiquetas de Planner** (bidireccional): las 25 de Planner (`category1`…`category25`, colores fijos aproximados en `PL_COLORS`); en la bandeja salen las que tienen nombre o están puestas en alguna tarea; "Gestionar" edita sus nombres. Personal sigue con las etiquetas de la app.
- Tarjetas como las de Personal, sin fechas; icono de bocadillo si la tarea tiene notas (`hasDescription`). Planner no tiene comentarios accesibles: solo notas.
- La sesión del trabajo (SPA) dura ~24 h: al caducar, la cabecera muestra el aviso "Planner: reconectar" y Sincronizar lleva al login.

## Google: contactos (hecho, lectura y escritura)

- **Proyecto de Google Cloud del usuario** (modo pruebas, el usuario como usuario de prueba), People API activada, ID de cliente de tipo "Aplicación web" con orígenes `https://aivanjoe.github.io` y `http://localhost:8000` y URIs de redirección `https://aivanjoe.github.io/tareas/` y `http://localhost:8000/`. El ID de cliente se introduce en Ajustes → Google. El mismo proyecto servirá para Google Calendar.
- **Login sin librerías de Google:** OAuth 2.0 para aplicaciones de navegador (flujo implícito, `response_type=token`) por redirección a `accounts.google.com/o/oauth2/v2/auth`, con `state` aleatorio guardado en `sessionStorage` (así la vuelta de Microsoft no se confunde con la de Google). Permisos: `contacts` (leer y editar) y `email` (para saber la cuenta); en el proyecto de Google tiene que estar añadido `.../auth/contacts` en Acceso a datos. `mp.google.write` indica si el permiso concedido permite editar (con uno de solo lectura, al editar se vuelve a pedir con `prompt=consent`). Primera vez con `prompt=select_account`; después con `login_hint`. El token vale una hora: no se renueva solo. Google muestra el permiso de contactos como casilla desmarcada: si vuelve sin permiso de contactos, se avisa y el botón «Volver a pedir» repite el login con `prompt=consent`.
- **Lectura:** `people/me/connections` (paginado) con `names, emailAddresses, phoneNumbers, birthdays, biographies, memberships, addresses, organizations`, y `contactGroups` para los nombres de las categorías (solo las del usuario; "Destacados" se guarda como `favorito`). Se sustituyen enteras las colecciones `contactos` (`{id: 'g-<id>', resourceName, etag, nombre, nombrePila, apellidos, telefonos/emails/direcciones: [{valor, tipo (tipo de Google: mobile, home, work…), etiqueta}], cumple: 'AAAA-MM-DD'|'--MM-DD', notas, etiquetas: [nombres], grupos: [resourceName], favorito, org: {nombre, cargo}, empresa}`) y `gruposContactos` (`{id: resourceName, nombre}`), que se sincronizan con OneDrive como lo demás.
- **Escritura** (formulario de contacto): `people:createContact`, `{resourceName}:updateContact` con el `etag` guardado y `updatePersonFields=names,phoneNumbers,emailAddresses,birthdays,biographies,organizations,addresses`, `contactGroups` (crear categoría), `{grupo}/members:modify` (categorías y favorito = `contactGroups/starred`) y `{resourceName}:deleteContact`. Tras cambiar grupos se relee el contacto para tener su `etag` nuevo. Si Google responde que el `etag` no coincide (alguien lo cambió en Google), se relee el contacto, se avisa y se reabre el formulario con los datos nuevos (no se pisa nada). Si el permiso caduca al pulsar Editar o al guardar, se guarda lo escrito en `sessionStorage` (`mp.google.next`), se va a Google y al volver se reabre el formulario con el borrador. Un contacto guardado por una versión anterior (sin `etag`) se relee antes de editarlo.
- **Cuándo:** al volver del login, con «Actualizar contactos» (va a Google si el permiso caducó) y al pulsar Sincronizar **solo si el permiso sigue vigente** (Sincronizar nunca saca de la página).
- Borrar un contacto se confirma y recuerda que Google Contactos permite "Deshacer cambios" hasta 30 días.

## Calendario: Google Calendar (hecho)

- Mismo proyecto de Google y mismo login que los contactos, con los permisos `calendar.events` y `calendar.calendarlist.readonly` (añadidos en Acceso a datos; Google Calendar API activada). Con un permiso antiguo sin Calendar, la pantalla ofrece "Dar permiso" (login con `prompt=consent`).
- **Calendario elegido** en Ajustes → Google (lista de calendarios en los que se puede escribir); se guarda en el almacén (colección `ajustes`, registro `calendario`: `{calendarId, nombre, color}`), así vale para todos los dispositivos. Su nombre no va en el código.
- **Eventos de la app**, marcados con `extendedProperties.private { mpApp: '1', mpKey }`, de todo el día, transparentes y con los avisos por defecto del calendario (`reminders.useDefault`):
  - Título con el nombre corto primero: nombre de pila y primer apellido (o las dos primeras palabras del nombre si Google no los tiene separados), p. ej. "Ana López · cumpleaños" y "Ana López · 15.000 días".
  - `cumple:<contacto>`: cumpleaños anual (`RRULE:FREQ=YEARLY`; los del 29-F, `BYMONTH=2;BYMONTHDAY=-1`) de los contactos seguidos en Cumples, desde su fecha de nacimiento (o el año actual si no tiene año). Color 4.
  - `dias:<contacto>:<N>`: los dos próximos cumpledías especiales (múltiplos de 5.000 días). Color 6.
  - `gasto:<gasto>`: gastos activos con la casilla "Mostrar en el calendario", "Concepto · importe", repetido según sus meses y día (`FREQ=MONTHLY;BYMONTH=…;BYMONTHDAY=d`; si el día es 29-31, `BYMONTHDAY=28..d;BYSETPOS=-1` para caer en el último día de los meses cortos). Color 5.
  - `tarea:<tarea>`: tareas de Personal con "Fecha en el calendario" (colección `fechasTareas`: `{id: taskId, fecha}`), con el nombre de la tarea. Al completar la tarea se borra la fecha (y el evento); deshacer la recupera. No toca la fecha de Todoist. Color 2.
- **Puesta al día (`calReconcile`)**: lee los eventos con la marca de la app, crea los que faltan, cambia los que no coinciden (título, fechas, repetición, color) y borra los que sobran. **La app manda sobre sus eventos**: si se mueven o borran en Google, vuelven. Nunca toca los demás eventos. Sin datos de Todoist cargados no toca los eventos de tareas.
- **Cuándo:** al pulsar Sincronizar (si el permiso de Google sigue vigente), al volver del login de Google, al elegir calendario y ~2,5 s después de cada cambio en el almacén (`saveDatos`) o al renombrar una tarea.
- **Pantalla Calendario** (antetítulo: el nombre del calendario; botón principal "Hoy"): vista de mes (lunes a domingo) con los eventos de cada día (en el móvil, puntos de color) y, debajo, la lista del día elegido con hora y enlace a Google Calendar. Colores: cumpleaños rosa, cumpledías ámbar, gastos amarillo, tareas lima y el resto con el color del calendario. Al sincronizar se leen desde dos meses antes hasta un año después; al ir a un mes fuera de ese tramo se lee si el permiso sigue vigente.

## Integración con Todoist (hecha)

- API v1, endpoint único `https://api.todoist.com/api/v1/sync`, autenticación `Authorization: Bearer <token>`, cuerpo `application/x-www-form-urlencoded`.
- Lectura: `sync_token='*'` con `resource_types` `["user","projects","sections","items","notes"]`. Se descartan tareas completadas, borradas y archivadas, y los comentarios borrados o de tareas no activas.
- Escritura mediante `commands` (cada uno con `uuid` propio; `temp_id` al crear): `item_add`, `item_update` (renombrar), `item_move` (cambiar de sección; a "Sin sección" se mueve con `project_id`), `item_reorder`, `item_close` (completar; respeta tareas recurrentes), `item_uncomplete` (deshacer), `note_add` / `note_update` / `note_delete` (comentarios).
- Se comprueba `sync_status` de cada comando; si algo falla se muestra aviso y se vuelve a leer todo de Todoist.
- Plan gratuito de Todoist: 5 proyectos, 20 secciones, 300 tareas activas. La API es gratuita.

## Decisiones de producto (respetarlas)

- **Sincronización manual.** Solo se conecta con las fuentes cuando el usuario pulsa Sincronizar o hace un cambio (mover, crear, renombrar, completar, cambiar datos de la app). Al abrir la app se muestra la última copia guardada. No añadir sincronización periódica ni al arrancar.
- **Pantallas separadas: Personal, Trabajo, Calendario y las propias de la app (Contactos, Gastos, Consumos, …).** No mezclar Planner y Todoist en un mismo tablero. La pantalla activa se recuerda por navegador (`mp.ui.view`). La cabecera cambia con la pantalla: título, antetítulo y el botón principal ("Añadir tarea", "Añadir persona"…); Sincronizar es común.
- **Contactos** (en el menú; antetítulo "Agenda y cumples"), con dos pestañas, **Agenda** y **Cumples** (la elegida se recuerda en `mp.ui.ctTab`, internamente `agenda` / `cumpledias`; el botón principal es "Nuevo contacto" en Agenda y "Elegir contactos" en Cumples):
  - **Agenda:** contactos de Google con buscador (nombre sin tildes, teléfono o correo), filtros Todos, Favoritos, En Cumples y por categoría, botón "Colores" (elegir el color de cada categoría: colección `coloresContactos`, `{id: resourceName del grupo, color}`; sin elegir, el color sale del nombre), botón "Actualizar desde Google", un **pastel en cada fila** para marcar o desmarcar en Cumples sin abrir la ficha (rosa si está marcado), y ficha (teléfonos con `tel:`, correos con `mailto:`, cumpleaños con edad, categorías, empresa, direcciones, notas, casilla "Seguir en Cumples", "Abrir en Google" y **Editar**). Formulario de contacto: nombre y apellidos, teléfonos, correos y direcciones como filas (valor + tipo + quitar), cumpleaños (día, mes y año opcional), categorías (fichas; se pueden crear nuevas), favorito, empresa y cargo, notas, y Borrar. Las **categorías de Google van separadas de las etiquetas de tareas**.
  - **Cumples** (cumpleaños y cumpledías juntos): **los contactos que marca el usuario** (colección `cumpledias`: `{id: id del contacto}`; se marcan con el pastel de la Agenda, con "Elegir contactos", que lista solo los que tienen cumpleaños, o con la casilla de la ficha). Las personas añadidas a mano en versiones anteriores (colección `personas`) siguen saliendo con la marca "a mano" mientras existan, pero ya no se pueden crear nuevas; si un contacto marcado se llama igual, no se repiten. Un cumpleaños sin año solo genera el cumpleaños (sin edad ni días de vida). Muestra, desde hoy incluido, los tres próximos **múltiplos de 1.000 días de vida** de cada persona y su próximo cumpleaños (los nacidos el 29 de febrero lo celebran el 28 en años no bisiestos). Fechas calculadas en UTC por día para evitar saltos de horario. Arriba: próximo evento, eventos este mes y este año, y número de personas; filtro Todo / Cumpleaños / Días de vida; **colores**: cumpleaños destacados en rosa `#E68AC4` (insignia rellena con la edad y fila tintada), hitos redondos (múltiplos de 5.000 días) en ámbar `#F2B35B` y el resto de hitos (cada 1.000) en azul, más discretos; leyenda encima de la lista; lista de personas para añadir, editar o borrar. Al cambiar de día se recalcula sola.
- **Vista de tablero** con las columnas de la fuente: secciones de Todoist, buckets de Planner. "Sin sección" solo aparece si tiene tareas o si no hay secciones.
- **Las columnas son de solo lectura en la app.** Crear, renombrar, reordenar o borrar columnas se hace en el origen.
- **Mover tarjetas entre columnas es crucial** y sí se escribe en el origen.
- **Tarjetas compactas:** casilla de completar + nombre en una línea (con elipsis). Si tiene etiquetas, van en una segunda fila, en pequeño. Borde lateral del color de la primera etiqueta.
- **Etiquetas de Personal:** gestor propio (nombre + color, sin iconos, ordenables), en Ajustes y en la bandeja. Varias por tarea. Son solo de la app, no se escriben en Todoist. **Trabajo usa las etiquetas de Planner** (ver arriba), no estas.
- **Asignar etiqueta:** arrastrándola desde la bandeja superior hasta la tarjeta. Pulsar una etiqueta de la bandeja filtra el tablero.
- **Quitar etiquetas:** desde el panel de la tarea, no arrastrando.
- **Añadir/editar tarea:** panel sencillo con nombre, etiquetas y, en Personal, "Fecha en el calendario" (la tarjeta muestra un icono de calendario si la tiene). "Añadir tarea" al pie de cada columna crea en esa columna; el botón de la cabecera permite elegir columna.
- **Menú lateral estrecho y ocultable** (preferencia recordada por navegador): 140 px (el logo de 88 px con aire alrededor), sin desplazamiento vertical desde 600 px de alto (por debajo de 720 px de alto pasa a una versión compacta, con el logo de 64 px); las etiquetas de tareas se gestionan desde Ajustes ("Gestionar etiquetas") y desde la bandeja del tablero, no desde el menú. Con el logo arriba, cada opción como icono con su nombre en pequeño debajo, sin contadores y el botón de ocultar al pie. En pantallas medianas y pequeñas pasa a ser un panel flotante.
- **Adaptable:** columnas que se reparten el ancho (mínimo 300 px), desplazamiento horizontal si no caben; en pantallas pequeñas una columna por pantalla con encaje al desplazar. En iPad, arrastrar tarjetas requiere mantener pulsado un instante.
- **Comentarios de Todoist:** se ven en el panel de la tarea (del más antiguo al más reciente, texto plano con saltos de línea; los adjuntos como enlace). Se pueden añadir, editar y borrar, y se guardan en Todoist al momento, aparte del botón Guardar. Si una tarea tiene comentarios, su tarjeta muestra un icono de bocadillo a la derecha del nombre (sin número).
- **Gastos** (antetítulo "Pagos recurrentes"; botón principal "Nuevo gasto"): colecciones `gastos` (`{id, concepto, proveedor, categoria, importe, periodicidad, meses: [1-12], dia, proximoPago: 'AAAA-MM-DD'|null, cobro: 'automatico'|'manual', notas, activo, creado}`) y `pagos` (`{id, gastoId, concepto, categoria, fecha: 'AAAA-MM-DD', importe, notas}`; el pago copia concepto y categoría para que el historial sobreviva al borrado del gasto).
  - Categorías fijas: Seguros, Impuestos, Comunidad, Inversiones, Suscripciones, Consumos, Préstamos, Alquileres, Streaming, Otros (cada una con su color).
  - **Próximo vencimiento** (misma lógica que el portal anterior, comprobada con sus datos reales): vencimientos según los meses de cobro y el día (ajustado al último día del mes si no existe); el último pago cubre el vencimiento que cae entre 15 días antes y 20 después, y toca el siguiente; sin pago, el primero desde hoy. La fecha fija `proximoPago` manda mientras no haya un pago que la cubra.
  - Pestaña **Próximos**: totales de lo que queda hasta fin de mes (destacado: el usuario cobra a final de mes), del mes que viene (todos sus vencimientos, incluida la siguiente cuota de los mensuales) y vencidos; filtro por categoría; cada gasto con su vencimiento, importe y botón Pagar; los pausados (`activo: false`) abajo y fuera de los totales. Pestaña **Historial**: pagos del más reciente al más antiguo, filtros por categoría y año, total y borrar. "Registrar pago" permite elegir gasto, fecha, importe y notas. Si el importe no es el habitual del gasto, aparece la elección **Solo este pago** (por defecto) / **Nuevo importe fijo** (actualiza el importe del gasto para los próximos cobros).
  - Importes escritos a mano en campo de texto con teclado decimal: se admite coma o punto.
- **Móvil:** Safari de iPhone/iPad amplía la página al tocar un campo con letra de menos de 16 px y la deja más ancha que la pantalla. Para evitarlo, en iOS `app.js` añade `maximum-scale=1` a la etiqueta viewport al arrancar (en iOS pellizcar para ampliar sigue funcionando) y la reaplica al cargar y en `pageshow`, porque Safari recuerda el zoom de la pestaña al recargar, y además en pantallas táctiles los campos de formulario van a 16 px.
  - Pestaña **Resumen** (por año, elegible): pagado en el año, media mensual (sobre los meses transcurridos) y previsión del año (pagado + vencimientos pendientes hasta el 31 de diciembre de los gastos activos, vencidos incluidos); gráfico de barras apiladas por categoría con lo pagado cada mes y línea discontinua con el total mensual del año anterior; totales por categoría con porcentaje y barra.
- **Consumos** (antetítulo "Luz, gas y agua"; botón principal "Subir facturas"): colección `consumos`, **un registro por factura** con id `<tipo>-<fin>` (`{id, tipo: 'luz'|'gas'|'agua', anio, mes, inicio, fin, emision, kwh, m3, coste, costeEnergia, costeServicios, precioKwh, precioM3, factura, notas}`). **Cada factura cuenta en el mes en que termina el consumo** (`anio`/`mes` salen de `fin`). Unidades: luz y gas en kWh, agua en m³. Colores: luz amarillo `#F5D76E`, gas verde `#8BD17C`, agua azul `#6FC3E8`.
  - Tarjetas de luz, gas y agua (último importe; pulsar elige tipo). **Próxima factura estimada:** fin del periodo por la duración habitual (mediana de los últimos periodos) e importe como media entre el mismo periodo del año pasado y la media de las tres últimas. Gráfico de barras del año actual frente al anterior por mes, con selector Coste / Consumo. Historial con periodo, consumo, precio, importe, ver factura, editar y borrar (borrar no quita el PDF de OneDrive).
  - **Subir facturas:** uno o varios PDF; se leen en el navegador con pdf.js y `parseFactura()` (quita los espacios del texto y busca patrones; `#` en los patrones es un número "1.234,56"). Formatos soportados: la compañía actual de luz y gas, y la de agua. Si ya existe la factura (mismo tipo y fin), se actualiza. Con OneDrive conectado, el PDF se guarda en `facturas/AAAA/<tipo>_<fin>.pdf` (subida simple, máximo 4 MB); sin OneDrive se guardan solo los datos y se avisa. Ver factura: descarga desde OneDrive y abre el PDF en una pestaña nueva (abierta en el clic para que no la bloqueen).
  - **Añadir a mano:** tipo, inicio (opcional) y fin del consumo, importe y consumo.
  - Si energía + servicios no suman el total de una factura (descuentos, regularizaciones), manda el total a pagar.
- La descripción de las tareas de Todoist no se muestra ni se edita.
- Las subtareas de Todoist no se muestran por ahora.

## Estilo visual

Oscuro y elegante. Tipografía Manrope (Google Fonts) con alternativas del sistema.

- Fondo `#0B0F15`, barra lateral `#0F141B`, columnas `#121821`, tarjetas `#18202B`, paneles `#10161E`.
- Texto `#E7EAF0`, secundario `#C9D0DC`, apagado `#9AA4B5` / `#8A94A6`.
- Acento principal lima `#DBF266` (botón primario, marca), azul `#7FA2E0` (selección, foco), rojo suave `#F0857A` (errores, borrar).
- Esquinas redondeadas (9–18 px), bordes de 1 px semitransparentes, sin degradados ni sombras exageradas.
- Iconos de línea en SVG, sin emojis. Toda la interfaz en español.
- **Marca:** un vegvísir dibujado en línea, en oscuro `#0B0F15` sobre cuadrado lima redondeado. El dibujo detallado está en la barra lateral (SVG en línea en `index.html`) y en `apple-touch-icon.png`. `favicon.svg` usa una versión simplificada (círculo, anillo central y ocho brazos con un remate corto), que se lee mejor a 16 px; la eligió el usuario.

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

La App Registration ya existe (tipo SPA, cuentas de cualquier organización y personales, registrada por el usuario en el tenant del trabajo); su Client ID se introduce en Ajustes.

1. **Bloques traídos de un portal anterior del usuario** (hechos: Cumpledías, en la pantalla Contactos; Gastos y pagos; y Consumos) (web con servidor propio), rehechos con el estilo y la arquitectura de esta app. Cada uno es una colección del almacén y una pantalla en la barra lateral. Los datos reales se importan con un fichero generado fuera del repo; nunca se suben aquí.
   - **Inversiones:** precios y gráficas de mercado. BTC y ETH con datos públicos de Binance; SPY necesita un proveedor que admita llamadas desde el navegador, con clave gratuita introducida en Ajustes.
2. Más adelante: elementos propios de la app (pagos recurrentes, cumpleaños, recordatorios) y captura rápida de tareas.
