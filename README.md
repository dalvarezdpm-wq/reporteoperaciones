# Oñate Reporte

Aplicación web (Vite + Firebase Realtime Database) para digitalizar el reporte operativo diario de previos, despachos, guías revalidadas y pendientes de Oñate, Willy & Cía. Captura 100% manual desde el navegador — no usa inteligencia artificial ni servicios de pago.

## Estructura del proyecto

```
reporte-ow/
├── .github/
│   └── workflows/
│       └── deploy.yml       # Compila y publica en GitHub Pages automáticamente
├── index.html               # HTML raíz (Vite)
├── package.json
├── vite.config.js
├── src/
│   ├── main.js               # Lógica de la app, UI y estado
│   ├── firebase.js            # Inicialización de Firebase (tu configuración real)
│   ├── storage.js             # Lectura/escritura del historial en Realtime Database
│   ├── parseDocx.js           # Lector gratuito de Word (.docx), sin IA
│   ├── parsePdf.js            # Lector gratuito de PDF, sin IA (best-effort)
│   ├── fields.js               # Definición de columnas de previos/despachos
│   └── style.css
└── .gitignore
```

## Qué hace

- Sube un `.docx` o `.pdf` ya llenado con el formato de la empresa (Fecha, Coordinador, Tramitador, tabla de previos, tabla de despachos, otras actividades, firmas).
- Lee automáticamente los datos — Word por estructura real de tabla (exacto), PDF por posición del texto (aproximado).
- Pantalla de revisión para corregir cualquier dato antes de guardar.
- Guarda cada reporte en **Firebase Realtime Database**, registrando: fecha del reporte, nombre del archivo original, tipo de origen (Word/PDF), quién lo subió, fecha/hora de carga, y todos los datos capturados.
- Historial de reportes anteriores agrupado por fecha, con estadísticas (previos vs. despachos, top clientes, actividad de 14 días) y exportación a CSV por día.

## 1. Firebase — ya configurado

Este proyecto ya trae tu configuración real de Firebase en `src/firebase.js` (proyecto `reporteaduana`). Verifica solamente que las **reglas de Realtime Database** permitan lectura/escritura:

En Firebase Console → Realtime Database → pestaña **Reglas**:
```json
{
  "rules": {
    ".read": true,
    ".write": true
  }
}
```
(Reglas abiertas están bien para este uso interno de dos personas; si más adelante quieres restringirlas por autenticación, dímelo y lo ajustamos.)

**No necesitas Firebase Storage.** El Excel compartido se guarda como datos dentro de Realtime Database (el mismo servicio que ya usas), evitando el requisito de plan de pago que pide Storage.

## 2. Correr en local

```bash
npm install
npm run dev
```
Abre la URL que muestra la terminal (normalmente `http://localhost:5173`).

## 3. Desplegar en GitHub Pages (automático, sin Vercel)

Este proyecto ya incluye `.github/workflows/deploy.yml`: cada vez que subas cambios a la rama `main`, GitHub compila y publica la app automáticamente — no necesitas Vercel ni compilar nada a mano.

1. Sube este proyecto a un repositorio de GitHub (rama `main`).
2. En el repositorio: **Settings → Pages → Source** → selecciona **"GitHub Actions"** (no "Deploy from a branch").
3. Haz cualquier cambio y `git push`, o entra a la pestaña **Actions** del repo y ejecuta manualmente el workflow "Desplegar a GitHub Pages".
4. En unos minutos tu app estará en `https://TU_USUARIO.github.io/TU_REPO/`. La URL exacta aparece en **Settings → Pages** una vez publicada.

Cada push futuro a `main` vuelve a compilar y publicar solo, sin que tengas que hacer nada más.

## 4. Desplegar en Vercel (alternativa opcional)

Si prefieres Vercel en vez de GitHub Pages:

1. Entra a [vercel.com](https://vercel.com) → "Add New… → Project" → selecciona el repositorio.
2. Vercel detecta Vite automáticamente (build command `vite build`, output `dist`) — no necesitas configurar nada.
3. "Deploy". Tu app quedará en una URL tipo `https://reporte-ow.vercel.app`.

## Novedades de esta versión

- **Corregido: "No se pudo guardar la asignación en la nube" al editar asignaciones viejas.** El error real era que Firebase Realtime Database rechaza CUALQUIER guardado que tenga un valor `undefined` en algún campo — y una asignación vieja sin `fechaCompletado` guardado producía justo eso al editarla. Se corrigió el origen del problema (siempre usa un valor válido, nunca `undefined`), y además se agregó una limpieza automática en todos los guardados a Firebase (reportes, catálogos, asignaciones, respaldos), para que este tipo de error no pueda volver a tumbar un guardado por ningún otro campo en el futuro.

- **Administrador y Coordinador ahora son dos entradas separadas.** Antes, Coordinación compartía la contraseña de Admin y (si había coordinadores en el catálogo) pedía elegir el nombre después. Ahora:
  - **Administrador**: sigue igual — una sola contraseña compartida, entra directo.
  - **Coordinador**: flujo aparte — elige tu nombre de la lista y entra con tu **propia contraseña individual** (mismas iniciales + 2026 que ya usa Trámite). Dentro de la app tiene exactamente los mismos permisos que Administrador (KPIs, Asignaciones, Catálogos, borrar directo sin pedir autorización) — nada más queda tu propio nombre marcado en cada cosa que hagas, en vez de un genérico "Administrador".
  - El botón "Coordinador" solo aparece en la pantalla de entrada si hay nombres cargados en el catálogo de Coordinadores; si está vacío, no se muestra.
  - En Catálogos → Coordinadores, cada nombre muestra su contraseña individual junto al chip, igual que ya pasa con Tramitadores.

- **Corregido: guías revalidadas aparecían siempre en "Completar referencias".** Ese apartado no tiene campo de Referencia, así que el sistema las contaba como si SIEMPRE les faltara una — se excluyeron del todo de ese conteo.
- **Nuevo: unificar duplicados de Clientes** (igual que ya existía para Ejecutivos) — botón "🔀 Unificar duplicados de Clientes" en Catálogos cuando detecta variantes como "stanley" y "STANLEY". A diferencia de Ejecutivos, la detección revisa los datos reales capturados (Previos, Despachos, Revalidadas, Pendientes, Asignaciones), no solo el catálogo, ya que un cliente puede estar escrito con variaciones sin nunca haberse agregado ahí. Al unificar, corrige todos los renglones guardados para que apunten al mismo nombre.
- **Las estadísticas de clientes (Top clientes, KPIs por cliente) ya agrupan de forma tolerante a mayúsculas/espacios de inmediato**, sin esperar a correr la unificación manual — mismo criterio que ya se usaba para tramitadores.

- **Nueva escala de puntos en Productividad ponderada**: Previos (A=0.06, B=0.10, C=0.14, D=0.18), Despachos (A=0.02, B=0.05, C=0.08, D=0.11), Guías revalidadas (0.02 fijo), Otras actividades (0.02 fijo por cada una). Un "2do previo" ahora **resta -0.06** en vez de sumar. Se actualizaron el panel del inicio y la tabla de KPIs con el desglose completo.
- **Corregido: pendientes borrados que volvían a aparecer solos.** La reconciliación automática "revivía" pendientes que alguien había borrado a propósito, pensando que nunca habían llegado a su destino. Ahora, al borrar la copia de un pendiente que vino de un traslado, se marca el original histórico como descartado para que nunca regrese solo.
- **Ahora se puede borrar cualquier pendiente**, incluyendo el respaldo histórico fijo (con confirmación aparte, por ser un respaldo).
- **Nuevo: varios coordinadores por aduana.** Se agregó el catálogo "Coordinadores" — si tiene nombres cargados, después de la contraseña de Admin se pide elegir cuál de ellos es, para quedar identificado con su propio nombre (pensado para Toluca, que tiene varios).
- **Querétaro ahora corta operaciones a medianoche**, no a las 8pm como el resto de las aduanas — el horario de corte ya es configurable por aduana.
- **Nuevo: control de eliminación — solo Admin borra directo.** Trámite y Ejecutivo ahora solo pueden *solicitar* borrar un renglón o quitar una foto — queda marcado "🗑️ Esperando autorización" hasta que Admin lo apruebe o lo rechace. Nueva pantalla **"🗑️ Solicitudes"** (barra superior, con contador) donde Admin ve y resuelve todas las solicitudes pendientes de cualquier reporte, sin tener que entrar uno por uno.

- **Nuevo: detección de asignaciones "completada" sin rastro real.** El cierre automático marca una asignación como completada en el momento que detecta la guía en el reporte del tramitador, pero nunca revisaba hacia atrás — si esa guía se borraba o corregía después, la asignación se quedaba "completada" para siempre sin que nadie lo notara. Ahora, en la lista de "Completadas", cada una se compara contra todos los reportes reales — si no se encuentra la guía en ningún lado, aparece un aviso amarillo claro y un botón **"↩️ Regresar a pendiente"** para corregirlo a mano (no borra nada del reporte, solo corrige el estatus de la asignación).

- **Nuevo: respaldo automático diario.** Cada vez que un Admin entra (máximo una vez al día por aduana), la app guarda sola una copia completa de reportes, catálogos y asignaciones en una rama aparte de Firebase — un punto de recuperación real sin depender de que alguien se acuerde de descargar el Excel a mano. Se conservan los últimos 14; los más viejos se borran solos. Nueva pantalla **"🗄️ Respaldos"** (barra superior, solo Admin) para ver la lista y descargar cualquiera como archivo JSON.
- **Confirmado: las notificaciones push del navegador ya estaban construidas** (permiso, aviso del sistema operativo al recibir/completar asignaciones, historial dentro de la app) — no hizo falta agregar nada ahí.
- **Auditoría de 31 puntos donde un error se ignoraba en silencio** — se corrigieron los 7 con riesgo real de que algo se perdiera o quedara desactualizado sin que nadie se enterara: guardar el caché local (ahora con fotos puede llenarse), marcar algo como "pendiente de subir" (el caso más grave — sin esto ni siquiera quedaba rastro para reintentar), guardar catálogos, reasignar un pendiente entre tramitadores, crear asignaciones masivas, y eliminar una asignación. Ahora, si cualquiera de estos falla, aparece un aviso claro en vez de fallar callado.

- **La autoasignación de guías ("🎯 Guías disponibles") ya no es exclusiva de Toluca — ahora aplica a las 5 aduanas** (GDL, Toluca, AIFA, AICM, Querétaro). Cualquier tramitador puede ver la "bolsa" de asignaciones sin tramitador y agarrar las que él va a trabajar, siempre asignándose a sí mismo, sin poder elegir a nadie más.

- **Corregido: tomar una foto de guía revalidada podía cerrar la app a medio proceso.** La forma en que se leía la foto antes de comprimirla cargaba el archivo completo (varios MB sin comprimir) como texto en memoria antes de siquiera empezar a procesarlo — en celulares con poca memoria libre, eso podía tumbar la app antes de terminar. Se cambió a un método mucho más ligero (referencia directa al archivo, sin copiarlo entero primero), que debería evitar el problema en la gran mayoría de los casos. También se agregó un aviso claro si la foto pesa más de 20 MB, en vez de intentar procesarla a fuerzas.

- **Corregido un bug serio de sincronización: datos viejos revivían solos y tapaban lo nuevo.** Cuando un dispositivo se quedaba con una copia local "pendiente de subir" que nunca lograba sincronizarse (por señal intermitente, o porque el guardado fallaba por otro motivo), esa copia vieja se volvía a pegar encima de los datos frescos del servidor **cada vez que la app cargaba o se actualizaba sola** — sin importar qué tan vieja fuera, ni si alguien más ya había cambiado algo en ese mismo día desde entonces. Esto explicaba tanto pendientes que "revivían" después de borrarlos, como reportes que parecían perderse al crear algo nuevo.
  - Ahora, antes de traer datos frescos (al cargar la app, y cada 25 segundos mientras haya señal), se intenta subir primero cualquier cosa pendiente — así una copia atorada tiene muchísimas más oportunidades de sincronizarse sola en segundos, en vez de quedarse tapando lo nuevo indefinidamente.
  - Si algo de verdad lleva más de 10 minutos sin poder subirse a pesar de tener señal, ahora aparece un aviso rojo bien visible en la barra superior ("⚠️ N sin subir hace rato — reintentar") con un botón para forzar el reintento a mano, en vez de fallar en silencio.

- **Nuevo: autoasignación de guías para Trámite, exclusivo de Toluca.** Cuando Coordinación crea asignaciones (individual o masiva) sin decidir todavía el tramitador, cualquier persona de Trámite en Toluca ahora ve un botón nuevo "🎯 Guías disponibles" con la "bolsa" de esas guías sin asignar. Puede tocar "Agarrar esta guía" en las que él vaya a trabajar — se asigna **siempre a su propio nombre** (con el que inició sesión), nunca puede elegir a otra persona. En cuanto se la asigna, aparece en su panel normal de "Mis asignaciones", igual que si Coordinación se la hubiera puesto directamente.

- **Corregido: la foto de una guía revalidada podía desaparecer si no se llegaba a guardar el reporte completo.** Antes, la foto solo se guardaba de verdad en la nube hasta darle clic al botón final "Guardar cambios" — si algo interrumpía antes de eso (se cerró la app, se salió sin querer, se fue la señal), la foto se perdía aunque se hubiera visto bien en el momento. Ahora, al tomar o quitar una foto en un reporte que ya existe guardado, se sube a la nube de inmediato, sin esperar al botón final.

- **Se quitó el enlace a SOIA** en Despachos (no resultó útil) — se revirtió a como estaba en v5.30.0.
- **Nuevo: tomar foto en Guías revalidadas.** Cada renglón de "Control de guías revalidadas" ahora tiene un botón 📷 — al tocarlo, abre la cámara del celular (o el explorador de archivos en compu) para tomar/elegir una foto y guardarla ahí mismo, en ese renglón. La foto se comprime automáticamente antes de guardarse (para no inflar el espacio en la base de datos) y se puede ver en grande tocando la miniatura, o quitar con la ✕. También se ve en la vista de detalle de solo lectura del reporte ya guardado.

- **Nuevo enlace directo a SOIA junto al campo Pedimento en Despachos.** Un botón 🔗 abre la consulta rápida oficial de SOIA (sin usuario ni contraseña) en una pestaña aparte. No llena nada automáticamente — el sitio de gobierno bloquea el acceso automatizado — pero ahorra tener que buscar el link cada vez; solo hay que escribir el pedimento ahí y copiar el resultado a mano al campo "Resultado".

- **La reparación de Ejecutivos ya no es un botón manual — corre sola, una sola vez, y desaparece para siempre.** Se quitó de la pantalla de Asignaciones; ahora se dispara automáticamente en segundo plano la próxima vez que un Admin entre a la app (sin ningún aviso ni interrupción), rellena lo que falte, y queda marcada con una bandera permanente para nunca volver a correr — ni la vas a volver a ver.

- **Nueva herramienta: "🔧 Rellenar Ejecutivos faltantes desde asignaciones existentes"** (botón en Asignaciones, solo Admin). Repara renglones viejos de Previos/Despachos/Pendientes que se capturaron antes de que existiera la columna Ejecutivo — busca, guía por guía, si hay una asignación ya guardada con ese mismo dato, y lo rellena. Nunca pisa un Ejecutivo que ya esté puesto, solo llena huecos. Pide confirmación antes de aplicar y muestra cuántos renglones se actualizaron.

- **Se muestra el Ejecutivo asignado en "Completar referencias".** Cada guía pendiente ahora tiene una etiqueta "👤 Le toca a: [nombre]" (o "Sin ejecutivo asignado" si no tiene uno) — útil cuando varios ejecutivos manejan al mismo cliente, para que cada quien identifique de un vistazo cuáles guías le corresponden a él.

- **Nuevo campo "Pedimento" en Asignaciones**, individual y en carga masiva. Se agregó un cuadro más en la carga masiva (5 columnas ahora: Guías, Clientes, Referencia, Pedimento, Ejecutivo). Viaja correctamente por todo el flujo: Asignación → Pendientes → (al marcar "Listo" en tipo Despacho) → columna Pedimento en Despachos. También se muestra en todas las listas de asignaciones (pendientes, completadas, sin tramitador, y el panel de "Mis asignaciones" del tramitador).

- **Las asignaciones sin tramitador ahora respetan el corte de las 8:00 PM antes de moverse a Pendientes.** Antes aparecían en Pendientes de inmediato, aunque se hubieran creado ese mismo día. Ahora: si se crean HOY (antes de las 8pm), se quedan vivas en la pantalla de **Asignaciones**, en su propia sección "⚠️ Sin tramitador asignado — hoy". Solo si cruzan el cierre de operaciones sin que se les ponga tramitador, se mueven solas a la pantalla de **Pendientes** — el mismo criterio que ya se usa para los pendientes normales.

- **Corregido: las asignaciones sin tramitador no se podían editar ni borrar.** Al mover esa sección a Pendientes, se quedó solo con el selector para asignar tramitador — le faltaban los botones de editar y eliminar. Ahora tiene los mismos tres: "Asignar", "✏️ Editar" (corrige guía, cliente, referencia, sector, tipo — cualquier dato, no solo el tramitador) y "✕" (eliminar). El botón de editar te lleva directo al formulario en Asignaciones con todo precargado, sin importar desde qué pantalla lo abriste.
- Se renombró el botón "✏️ Reasignar" a "✏️ Editar" en la lista normal de Asignaciones, para que quede claro que corrige cualquier dato, no solo el tramitador.

- **Todo el texto secundario de la app (antes gris) ahora es negro**, para que se lea mucho mejor — aplica en toda la aplicación de golpe (Asignaciones, Pendientes, KPIs, Catálogos, etc.), ya que estaba controlado por un solo color centralizado. Las etiquetas importantes (Guía, Tramitador, Ejecutivo, Cliente, etc.) siguen en negritas como ya estaban, ahora simplemente también en negro en vez de azul marino.

- **La sección "Sin tramitador asignado" se movió de Asignaciones a Pendientes.** Ya no aparece en la pantalla de Asignaciones — ahora vive como su propio apartado especial ("⚠️ Asignaciones sin tramitador") arriba de la lista de pendientes de reportes, en la pantalla de Pendientes. Mismo semáforo de siempre: rojo si se creó hoy, amarillo con contador de días si lleva más tiempo sin asignarse. El contador del botón "Pendientes" en la barra superior ahora también incluye estas.

- **Corregido: el Ejecutivo se perdía al traer una asignación a Pendientes.** Nunca se copiaba, y tampoco había una columna visible para mostrarlo. Ahora "Ejecutivo" es una columna real en Previos, Despachos y Pendientes — viaja correctamente desde la asignación hasta el reporte final, se ve tal cual se escribió, y queda incluida en el Excel y CSV exportados.
- **Ampliado el ancho del contenido** (de 1100px a 1360px) para dar más espacio a las tablas del reporte, que ahora tienen una columna más.
- **Corregida la tipografía del "almacén fijo"** (Querétaro) — antes se veía con letra más grande que el resto de la tabla; ahora combina exactamente con el tamaño y fuente que usa toda la app.

- **Corregido: los 4 cuadros de carga masiva se borraban solos.** No guardaban nada mientras se escribía/pegaba, así que la actualización automática de cada 25 segundos los vaciaba. Ahora cada uno guarda lo que llevas al instante, así que ya no se pierde nada aunque tardes en pegar todo.
- **El cuadro de Clientes ahora es más ancho** (y Guías también), para que se puedan leer bien los nombres completos, uno por línea, sin que se corten.

- **Rediseñada la carga masiva de asignaciones para pegar columna por columna.** Ahora son 4 cuadros de texto separados — Guías, Clientes, Referencia, Ejecutivo — pensado para cuando en el manifiesto/Excel se selecciona una columna completa a la vez, en vez de tener que armar renglones combinados a mano. La app junta todo por posición (el renglón 1 de cada cuadro se junta con el renglón 1 de los demás). Si algún cuadro queda con más o menos renglones que Guías, se avisa después de crear para poder revisar el orden.

- **Ahora se puede crear una asignación (individual o en carga masiva) sin elegir tramitador todavía** — útil cuando llega el manifiesto antes de saber quién lo va a trabajar. Aparece la opción "— Sin asignar aún —" en ambos formularios.
- **Nueva sección "⚠️ Sin tramitador asignado"** arriba de "Pendientes" en la pantalla de Asignaciones — junta todas las que se crearon sin tramitador, con un selector y botón "Asignar" por cada una (o varias de un jalón) para ponerles tramitador cuando ya se sepa. En cuanto se les asigna, aparecen solas en el reporte de esa persona al entrar a capturar, como cualquier otra asignación.

- **Simplificado: se quitaron las casillas de selección manual en Pendientes.** Ya no hace falta marcar nada a mano — cuando 2 o más despachos pendientes vienen de la misma carga masiva, la app los detecta sola y muestra un solo botón "✅ Marcar TODO este lote como Listo (N guías — Cliente)" para resolverlos todos de un jalón. Sigue aplicando solo a Despacho (Previo se mantiene siempre individual). Cualquier guía se puede sacar del lote con su ✕ normal antes de usar el botón.

- **La selección múltiple y los lotes de "Listo" ahora son exclusivos de Despacho.** En Previo, cada renglón puede terminar con un tramitador distinto, así que ahí se quedó siempre individual (solo el botón "✅ Listo" de cada renglón, sin casillas ni lotes). Despacho es normalmente una sola persona con todo el consolidado, así que ahí sí se mantiene la casilla de selección múltiple y el botón de "marcar todo el lote".

- **Almacén fijo "210 TERMINAL" para Querétaro** (exclusivo de esa aduana, por trabajar con manifiesto y un solo almacén) — se autocompleta y queda bloqueado 🔒 en todos los renglones de captura y en el formulario de asignaciones.
- **Carga masiva de asignaciones**, disponible para todas las aduanas: nuevo botón "📋 Carga masiva de asignaciones" en Asignaciones — pega varias guías de un jalón (desde Excel o separadas por comas), elige Tipo/Sector/Tramitador una sola vez, y crea todas de golpe. Pensado para manifiestos grandes, pero útil para cualquier aduana.
- **Lotes: marcar TODO un envío consolidado como "Listo" de un jalón**, disponible para todas las aduanas. Las asignaciones creadas juntas en una misma carga masiva quedan "amarradas" como un lote — en Pendientes aparece un botón para resolver el lote completo de un jalón. Cualquier guía se puede sacar del lote a mano con su ✕ normal antes de usar el botón, si necesita tratarse aparte. También se puede seguir marcando manualmente con casillas cualquier grupo de pendientes sueltos (no solo los que vienen de un lote).

- **Nuevo: corte de operaciones a las 8:00 PM.** Como la aduana cierra a esa hora, la app ya considera "el día siguiente" a partir de las 8:00 PM, no hasta medianoche. Esto afecta a toda la app de forma consistente: el reporte nuevo que se empiece a capturar después de las 8pm se fecha automáticamente para el día siguiente, y los pendientes sin resolver a esa hora ya se pueden trasladar a "mañana" desde el cierre real de operaciones. El "día operativo" corre de 8:00 PM a 8:00 PM sin huecos.
- Nota: si alguien deja abierta a medio llenar una hoja justo cuando cruza las 8pm, esa hoja puntual conserva la fecha con la que se creó — es un caso raro y no afecta los reportes ya guardados.

- **Se quitó "Tipo despacho" del Control de despachos** — quedaba redundante con "Tipo de operación" (Importación/Exportación).

- **Desglose individual por dificultad en Productividad ponderada** — ahora se ve cuántos A, B, C y D hizo cada tramitador, no solo el total de puntos. En la pantalla completa de KPIs aparece como columnas propias en la tabla; en el panel rápido del inicio aparece como texto chico debajo de cada barra (ej. "A:3 · B:5 · C:2 · D:1").

- **Nueva pantalla "Pendientes"** (botón en la barra superior, solo Admin) — junta TODOS los pendientes abiertos de todos los tramitadores y todos los días en un solo lugar, con opción de reasignar cualquiera a otra persona (venga o no de una asignación formal), sin tener que buscarlo dentro de cada reporte diario.
- **Corregido: los números de KPIs y Productividad no cuadraban.** La causa raíz: las estadísticas agrupaban por tramitador usando el nombre tal cual estaba escrito, sin tolerar variaciones de mayúsculas o espacios — si "Luis Arreola" quedó capturado alguna vez con una diferencia mínima, sus números se partían en dos filas distintas en todos los tableros (KPIs, Productividad ponderada, panel del inicio), inflando el conteo de "personas" y partiendo sus totales. Ahora todas las estadísticas agrupan de forma tolerante, y siempre se muestra el nombre tal como está en el catálogo actual.
- **Prevención de duplicados en Catálogos**: al agregar un nombre (uno por uno o en carga masiva), si ya existe uno igual con distintas mayúsculas/espacios, ya no se vuelve a agregar. Además, cada categoría muestra una alerta si detecta posibles duplicados ya existentes, para poder limpiarlos a mano.
- **Panel de Productividad ponderada** del inicio ahora tiene un link "Ver detalle diario y filtrar fechas →", y la pantalla completa de KPIs incluye la tabla de productividad ponderada con su propio filtro de fechas.

- **Nuevo botón "📥 Traer pendientes de días anteriores"** en la sección de Pendientes del reporte de hoy — igual que ya existía para asignaciones. Antes el traslado de pendientes solo se disparaba automáticamente al crear el reporte del día; ahora cualquier tramitador puede dispararlo a mano en cualquier momento mientras edita su reporte de hoy, con un número que muestra cuántos hay pendientes por traer. Es el mismo mecanismo de siempre (deja fijo el renglón original en su día viejo), solo que ahora se puede activar manualmente sin depender de que el automático se haya disparado en el momento correcto.

- **Corregido el espacio vacío en el inicio.** "Top clientes" (que ahora tiene 3 secciones internas) es mucho más alto que "Actividad" y "Distribución" — se separó a su propia fila para que ya no deje un hueco enorme abajo de esos dos paneles más cortos.
- **Nuevo: Dificultad (opcional) también en Despachos** — igual que ya existía en Previos (A/B/C/D), para poder medir productividad ponderada en ambos.
- **Nueva medición: Productividad ponderada por dificultad.** Cada previo y despacho suma puntos según su dificultad — A = 25 pts, B = 50, C = 75, D = 100 (un renglón sin dificultad asignada cuenta como A). Las guías revalidadas no entran en esta cuenta, a propósito. Aparece como panel en el inicio (últimos 7 días, solo Admin) y como tabla completa con su propio filtro de fechas en la pantalla de KPIs — muestra puntos totales, número de operaciones, y el promedio de dificultad por operación, para distinguir a alguien con pocas operaciones muy complejas de alguien con muchas operaciones sencillas.

- **Corregido: un pendiente podía quedar "trabado" sin llegar al día de hoy.** El traslado automático de pendientes marcaba el renglón viejo como "ya trasladado" y lo guardaba de inmediato, pero la copia hacia el día de hoy se guardaba en un paso aparte y posterior — si esa persona cerraba la app a medio camino (o algo fallaba en ese momento), la copia nunca llegaba a existir en ningún lado, aunque el original ya se había marcado como resuelto. Ahora, antes de trasladar pendientes nuevos, la app revisa si algún pendiente viejo quedó en ese estado a medias y lo destraba solo, para que se vuelva a intentar el traslado correctamente. Se corrige solo, sin que nadie tenga que hacer nada — pasa la próxima vez que esa persona entre a capturar su reporte.

- **Pendientes que no se resuelven ya no se pierden de vista al día siguiente.** Cuando un tramitador entra a capturar (o sigue editando) su reporte del día, la app revisa automáticamente sus pendientes de días anteriores que sigan sin resolver (por asignación o escritos a mano, da igual) y los agrega solos al día de hoy. El renglón original del día viejo queda congelado como respaldo — de solo lectura, sin poder borrarse ni editarse — así que si alguien borra por accidente el de hoy, el de ayer sigue existiendo.
- **Asignaciones: ahora se pueden reasignar/editar sin borrar y volver a crear.** Botón "✏️ Reasignar" en cada asignación pendiente — carga sus datos en el formulario de arriba para corregir cualquier cosa (incluyendo cambiar de tramitador), conservando quién la creó originalmente y cuándo.
- **Colores de asignaciones pendientes por antigüedad**: rojo = se creó hoy; **amarillo** = se quedó pendiente de un día anterior sin resolver (antes todas se veían del mismo tono rojo sin distinguir).
- **Se quitó "N° despacho"** del control de despachos.
- **Control de despachos reordenado**: ahora Pedimento, Referencia y Guía van primero; el resto sigue igual.
- **"Dificultad" ahora dice "(opcional)"** en la etiqueta, para dejar claro que no es obligatorio llenarla.
- **Nuevo campo "Tipo de operación" (Importación/Exportación)** en Previos y en Despachos.
- **"Resultado" en Despachos** ya no incluye "Verde"/"Rojo" — solo quedan "Desaduanamiento libre" y "Reconocimiento aduanero".
- **Encabezados de tabla en negro y negritas**, para que resalten más en los reportes.
- **Nuevo en el inicio (solo Admin/Coordinación): "Productividad — últimos 7 días"** y **"Pendientes atrasados por tramitador"** — dos paneles nuevos, más específicos que las gráficas generales que ya había, para ver de un vistazo quién lleva pendientes arrastrados de días anteriores.

- **Fila de "TOTAL" al final de cada tabla en KPIs y Clientes**, sumando cada columna por separado (total de Previos, total de Despachos, total de Revalidadas, etc.), no solo el total horizontal por renglón que ya existía. Aplica a las 4 tablas: KPIs por tramitador, detalle día por día, todos los clientes, y días con más movimiento.

- **Se quitó el escáner de código de barras por cámara.** Después de varios intentos no se logró que funcionara de forma confiable en todos los dispositivos — la app queda limpia y más ligera sin ese código a medio funcionar. El campo Guía sigue siendo texto normal, lo cual significa que **ya funciona con un lector físico USB/Bluetooth** sin necesitar ningún cambio adicional, si deciden ir por esa opción.
- **3 aduanas nuevas activadas: AIFA, AICM y QUERETARO** — mismo esquema que Toluca: catálogos de clientes/almacenes/tramitadores/ejecutivos completamente vacíos y separados (para que cada una arme el suyo desde cero), contraseñas de tramitador de 4 letras, y su propia contraseña de Administrador/Coordinación:
  - AIFA → `OWAIFA`
  - AICM → `OWAICM`
  - QUERETARO → `OWQUERETARO`

- **Se reconstruyó el escáner desde cero, usando el método propio y probado de la librería** (en vez del código hecho a mano que se armó en los últimos intentos, que resultó tener más bugs de los que resolvía — dejó de detectar cualquier código, ni cerca ni lejos). Ahora usa `decodeFromConstraints`, el ciclo de escaneo oficial de ZXing, optimizado internamente. Se conservan las dos mejoras que sí valían la pena: el filtro que ignora automáticamente códigos que no tienen forma de guía real, y la confirmación antes de aplicar el número al campo.

- **El escáner ahora filtra automáticamente los códigos que no parecen una guía real.** Las guías de FedEx son puramente numéricas (12, 15, o hasta 20-22 dígitos según el servicio) — si detecta un código que no tiene esa forma (como pasó antes con un "417992" de 6 dígitos, que era otro código de la misma hoja), lo ignora solo y sigue buscando, mostrando un aviso breve de "se detectó X pero no parece guía". Solo pide confirmar cuando encuentra algo con forma real de guía.

- **Corregido: el escáner dejó de detectar cualquier código, incluso bien encuadrado.** La causa: el "zoom digital" que se agregó en la versión anterior recortaba la imagen sobre los píxeles crudos de la cámara, sin tomar en cuenta que la pantalla ya recorta/reescala esa imagen para que quepa en el celular — el recuadro azul que se veía en pantalla no correspondía con la zona que realmente se analizaba. Se quitó ese recorte; ahora se analiza el cuadro completo de la cámara tal como se ve, igual que en las primeras versiones del escáner (que sí detectaban códigos). El paso de confirmación ("✅ Sí, usar" / "🔁 Escanear de nuevo") se mantiene.

- **El escáner ahora pide confirmar antes de usar el código leído.** Antes, en cuanto detectaba algo, lo metía directo al campo Guía y cerraba la cámara — si detectaba mal (otro código de la misma hoja, o una lectura corrupta), el número equivocado ya estaba puesto sin darse cuenta. Ahora, al detectar un código, la cámara se pausa y muestra el número en grande con dos botones: "✅ Sí, usar" (lo aplica al campo) o "🔁 Escanear de nuevo" (descarta esa lectura y sigue intentando, sin tener que volver a abrir la cámara ni pedir permiso otra vez).

- **Rediseñado el escáner para leer mejor cuando el código se ve chico en cámara.** Antes se analizaba la imagen completa de la cámara; ahora la app recorta digitalmente solo la franja central (donde está el recuadro guía en pantalla) y la **amplía** antes de intentar leerla — como un zoom digital automático. Esto ayuda mucho cuando se encuadra la hoja completa en vez de acercarse solo al código de barras.

- **Mejoras al escáner de código de barras**, para que lea mejor:
  - Pide a la cámara la mayor resolución posible (antes usaba la calidad por defecto, muy baja para códigos chicos) — si el celular no la soporta, reintenta automáticamente con algo más simple en vez de fallar.
  - **Nuevo botón de linterna 🔦** (aparece solo si el celular la tiene), para escanear con poca luz.
  - Mensaje en pantalla sugiriendo la distancia ideal (10-15 cm) del código.

- **Nuevo: escáner de código de barras con la cámara**, para llenar el campo Guía sin teclear. Aparece un botón 📷 junto al campo Guía en cada renglón de Previos/Despachos/Revalidadas/Pendientes, y también en el formulario de "Crear asignación". Al tocarlo se abre la cámara de pantalla completa (usa la cámara trasera en celular/tablet); en cuanto detecta el código de barras (funciona con el formato que trae FedEx en sus guías), lo escribe solo en el campo y cierra la cámara automáticamente — sin necesidad de tomar foto ni confirmar nada. Si esa guía ya tenía una Referencia capturada antes, se autocompleta sola también, igual que al escribirla a mano.

- **Corregido: a algunos tramitadores no les "jalaban" sus asignaciones a Pendientes.** La causa: la app comparaba el nombre del tramitador de forma exacta (sensible a mayúsculas/minúsculas y espacios extra) contra el nombre con el que inició sesión. Si el nombre se volvió a capturar en el catálogo con una letra distinta de mayúscula o un espacio de más (por ejemplo, al usar la carga masiva pegando desde otra fuente), la asignación dejaba de "verse" como suya aunque el nombre pareciera idéntico. Ahora la comparación ignora mayúsculas/minúsculas y espacios de más en todos los puntos donde se relaciona una asignación con su tramitador (traer a Pendientes, cierre automático, notificaciones, badge de asignaciones).

- **Nuevo catálogo de Ejecutivos, sin contraseña.** El login de Ejecutivo ya no pide escribir el nombre libremente — ahora se elige de una lista fija (como ya pasa con Trámite), sin necesidad de contraseña. Se agregó el catálogo inicial para GDL: Alberto Pichardo, Fernanda Narez, Claudia Barrera, Fernanda Ramirez, Ulises Bautista, Erendira Calderon. Coordinación/Admin puede editar esta lista en Catálogos, igual que con Tramitadores.
- **Se guarda quién dejó cada observación.** En Previos, Despachos y Pendientes, debajo de cada observación aparece el nombre del Ejecutivo que la escribió — visible para todos, pero solo el Ejecutivo puede seguir editándola.
- **Nuevo campo "Ejecutivo" en Asignaciones**, para elegir a quién le corresponde esa operación al crearla. Se muestra en las listas de asignaciones pendientes/completadas y en el panel de "Mis asignaciones" del tramitador.

- **"Observaciones" ahora también en Pendientes** (antes solo estaba en Previos y Despachos ya resueltos). Mismo comportamiento: solo el Ejecutivo puede escribir ahí, todo lo demás sigue de solo lectura. Y si esa guía se marca "✅ Listo" y sube de Pendientes a Previos/Despachos, la observación viaja con ella — no se pierde.

- **Nueva columna "Observaciones" en Previos y Despachos, editable solo por el Ejecutivo.** Desde "Ver reportes (solo lectura)", el Ejecutivo ahora puede escribir una observación directo en esa columna — es la ÚNICA celda editable para ese rol; todo lo demás (Guía, Cliente, Referencia, etc.) sigue siendo de solo lectura para él, tal como antes. Se guarda sola al salir del campo, sin necesidad de un botón aparte.

- **Asignaciones ahora incluye Referencia y Sector.** Coordinación/Admin puede capturarlos al crear una asignación, y viajan automáticamente hasta Pendientes y luego hasta Previos (columna Sector) cuando el tramitador trae la asignación y la marca "✅ Listo" — ya no hay que volver a escribirlos a mano. También se muestran en las listas de asignaciones pendientes/completadas y en el panel de "Mis asignaciones" del tramitador.

- **Nuevo: días pendientes en "Completar referencias" (Ejecutivo).** Cada guía sin referencia ahora muestra desde cuándo lleva pendiente y cuántos días exactos ("🕒 Pendiente desde 10/08/2026 — 4 días"), con color de alerta: gris si es reciente, ámbar a partir de 2 días, rojo a partir de 5. Dentro de cada cliente, las guías se ordenan mostrando primero las más antiguas (las más urgentes).
- **Nueva columna "Sector" en Previos**, con selector: Químico, Perecedero, Metalúrgico, Textil, Agricultura, Manufactura, Salud, Digital. Se incluye automáticamente en la vista de detalle, el CSV y el Excel exportado.

- **Se quitó la categoría "Aduanas" de la pantalla de Catálogos.** Era un residuo que no estaba conectado a nada real (la lista real de aduanas del login se controla aparte, en el código) — por eso se veía distinta entre GDL y Toluca, lo cual generaba confusión. No tenía ningún efecto funcional, así que quitarla no cambia nada del comportamiento de la app.

- **Contraseñas propias para Toluca, distintas de GDL.** Administrador/Coordinación en Toluca ahora usan `OWTOLUCA` (GDL sigue usando `ow2026`, sin cambios). Los tramitadores de Toluca tienen contraseña automática de **4 letras** iniciales de su nombre + 2026 (GDL sigue siendo 3 letras, como siempre — no se invalidó ninguna contraseña ya compartida en GDL).

- **La app se renombró a "Oñate Reporte"** — nuevo nombre en la pestaña del navegador, en la pantalla de login, en el encabezado, y al instalarla como app (PWA).
- **Se quitó la subida de Word y PDF.** Ya no existe la pantalla intermedia de "elegir cómo capturar" — el botón "+ Nuevo reporte" lleva directo al llenado manual. Esto también aligeró bastante la app (pesa menos de la mitad que antes, carga más rápido).
- **Ya no se pueden crear hojas duplicadas.** Antes, una persona de Trámite podía apretar "+ Nuevo reporte" varias veces el mismo día y terminar con varias hojas separadas a su nombre (repartiendo sus datos entre ellas por accidente). Ahora, si ya tiene una hoja capturada hoy, "+ Nuevo reporte" la abre directo para seguir editándola, en vez de crear otra.
- Recordatorio de algo que ya existía: solo Administrador/Coordinación pueden borrar una hoja completa — Trámite puede seguir editando libremente el contenido de la suya, pero no borrarla por accidente.

- **El Excel descargado (día individual e histórico completo) ahora tiene la identidad de OW.** Cada hoja incluye el logo de la empresa arriba, el título "Oñate, Willy y Cía., S.C." en los colores de marca, y los datos formateados como una **Tabla de Excel real** (con flechitas de filtro y franjas de color alternadas), con el encabezado en el azul marino exacto de OW y letras blancas en negrita. Aplica tanto a la descarga manual como al Excel compartido con el link automático.
- Cambio técnico de fondo: se reemplazó la librería que arma el Excel (la anterior no soportaba colores, logos ni tablas con estilo) por una que sí lo permite.

- **Agregada la opción "2do despacho"** al selector de N° despacho, por si se llega a necesitar.

- **Corregido: el autocompletado de Referencia no detectaba una guía repetida dentro del MISMO reporte sin guardar todavía.** La búsqueda solo revisaba reportes ya guardados; ahora revisa primero el reporte que se está llenando en ese momento (aunque no se haya guardado), y si no la encuentra ahí, busca en el historial. También se volvió más tolerante a mayúsculas/minúsculas y espacios al comparar guías.
- **N° previo y N° despacho ahora son selectores, no texto libre.** Previo: "1er previo", "2do previo", "3er previo". Despacho: "1er despacho". Quedan iguales a como se generan automáticamente al marcar un pendiente como "✅ Listo".

- **Catálogos preparado para volúmenes grandes (ej. miles de clientes).** Cuando una categoría pasa de 25 registros, ya no se dibujan todos los nombres de un jalón (eso volvería la pantalla lentísima con miles) — en su lugar aparece un buscador: escribes parte del nombre y solo se muestran las coincidencias (hasta 60 a la vez). El almacenamiento en sí no es problema (miles de nombres pesan poquísimo), el ajuste es puramente para que la pantalla siga siendo rápida y usable.

- **Nuevo: carga masiva en Catálogos.** En Clientes, Almacenes, Tramitadores y Aduanas ahora hay un botón "+ Agregar varios de un jalón" — abre un cuadro donde pegas muchos nombres de golpe (uno por línea, o separados por coma) y se dan de alta todos juntos con un solo clic, sin repetir los que ya existan. Pensado justo para cuando se arma el catálogo de una aduana nueva desde cero (ej. Toluca) y sería muy lento capturarlos uno por uno.

- **Nuevo: filtro Año / Mes en "Reportes por fecha".** Antes se listaban todos los días de golpe, uno tras otro — a la larga (varios meses/años de operación) se iba a volver una lista eterna. Ahora hay dos selectores arriba de la lista: elige el Año y luego el Mes (ej. "2026 → Agosto") y solo se muestran los reportes de ese mes. Por defecto están en "Todos" (se ve igual que antes); el filtro es opcional. Los contadores generales ("Días con reportes", el botón de Excel histórico completo, etc.) siguen contando TODO el historial sin importar el filtro — el filtro solo afecta qué se ve en la lista.

- **Activada la aduana TOLUCA.** Ya aparece como opción en la pantalla de inicio de sesión, junto a GDL. Sus datos (reportes, catálogos, asignaciones, Excel histórico) quedan completamente separados de GDL — cada aduana vive en su propio espacio dentro de la misma base de datos, y nadie ve los datos de la otra sin cambiar explícitamente de aduana.
- **Corregido: al ser una aduana nueva, su catálogo de tramitadores ya no se llena solo con los nombres del equipo de GDL.** Antes, cualquier aduana nueva heredaba por accidente esa lista fija; ahora empieza vacía, para que Coordinación dé de alta ahí a su propio equipo de Toluca.

- **Corregido: espacios en blanco feos en los paneles del inicio** (Actividad / Distribución / Top clientes). Como el panel de "Top clientes" creció (ahora tiene 3 listas separadas), estiraba a los otros dos paneles cortos a su misma altura, dejando un hueco vacío abajo. Ahora cada panel toma solo la altura de su propio contenido.

- **"Top clientes" en la página principal ahora está separado por Previos, Despachos y Guías revalidadas** (antes venía todo mezclado en un solo número). Incluye un enlace directo a "ver con filtro de fechas" que lleva al nuevo dashboard de Clientes.

- **Nuevo: Dashboard de clientes** (botón "Clientes" en la barra superior, junto a KPIs). Igual que el dashboard de tramitadores, pero por cliente y con filtro de fechas (Hoy / Últimos 7 días / Este mes / Todo). Incluye:
  - **Top clientes — Previos**, **Top clientes — Despachos** y **Top clientes — Guías revalidadas** por separado (antes "Top clientes" en el inicio los mezclaba todos en un solo número).
  - Tabla con todos los clientes del rango y sus totales.
  - **"Días con más movimiento"**: los 20 renglones día + cliente con más actividad en el rango — responde directo a "qué día se movió más tal cliente".

- **Autocompletado de Referencia por Guía repetida.** Al escribir una Guía en Previos, Despachos o Pendientes y salir del campo, si esa misma guía ya tiene una Referencia capturada en algún reporte anterior, se rellena sola — ya no hace falta buscarla ni volver a teclearla cuando se repite una guía (ej. un 2do previo de la misma guía). Si el campo Ref. ya tiene algo escrito, no se lo pisa.

- **"Control de guías revalidadas" simplificado**: ahora solo pide Guía, Cliente y Almacén — sin Ref., Motivo ni N° de revalidación.
- **Coordinación/Administrador ahora también puede crear asignaciones de "Guía revalidada"**, junto a Previo y Despacho, en la pantalla de Asignaciones. Se cierran solas igual que las demás en cuanto el tramitador guarda un reporte con esa guía en el nuevo control.

- **Nueva sección: "Control de guías revalidadas".** Se agregó como tercer control (después de Previos y Despachos, antes de Pendientes) con sus propios campos: Ref., Guía, Cliente, Almacén, Motivo de revalidación y N° de revalidación. Aparece en el formulario de captura, en el detalle de cada reporte, en la exportación a CSV y en el Excel descargable (nueva hoja "Revalidadas").
- **El selector "Tipo" de Pendientes ahora incluye "Revalidada".** Antes solo se podía elegir previo o despacho, y las guías revalidadas pendientes no tenían dónde clasificarse. Ahora, al elegir "Revalidada" en un pendiente, aparece el botón "✅ Listo" que lo sube directo al nuevo control de guías revalidadas (con su numeración automática "1ra/2da revalidación..." si la guía se repite).
- Guías revalidadas ahora también se incluyen en: la búsqueda global, la detección de guía duplicada, el aviso de "referencia faltante" para el Ejecutivo, los KPIs por tramitador (nueva columna "Revalidadas"), y la tarjeta de estadística "Guías revalidadas" en el inicio.

- **Corregido: el formulario de "Crear asignación" se borraba solo mientras lo llenabas.** La causa: la app se actualiza sola cada 25 segundos en segundo plano para traer datos nuevos, pero ese formulario no guardaba lo que ibas escribiendo hasta apretar el botón — entonces la actualización automática lo regeneraba vacío a media captura. Ahora lo que escribes (Guía, Cliente, Almacén, Tipo, Tramitador) se guarda al instante, letra por letra, así que ya no se borra sin importar cuánto tardes en llenarlo.

- **El "Historial de cambios" de cada reporte ahora está colapsado detrás de un botón.** Antes se mostraba siempre abierto, ocupando espacio y compitiendo visualmente con los datos reales capturados. Ahora aparece como un botón pequeño "🕒 Historial de cambios (N)" — al tocarlo se despliega quién creó o editó el reporte y cuándo; al tocarlo de nuevo se cierra. Así el cuadro de cada reporte se ve más limpio y la información importante resalta más.

- **Corregido bug importante: "hoy" se calculaba en hora UTC, no en hora de México.** Como México está detrás de UTC, desde cierta hora de la tarde/noche en adelante, la app ya pensaba que era el día siguiente — esto afectaba la fecha por defecto al capturar un reporte nuevo, el botón "Hoy" y "Últimos 7 días" en KPIs, y la gráfica de actividad de los últimos 14 días. Ya corregido para usar siempre la fecha del huso horario local del dispositivo, no UTC.

- **KPIs con filtro de fechas**: en la pantalla de KPIs ahora se puede elegir un rango específico ("Desde"/"Hasta"), o usar los atajos "Hoy", "Últimos 7 días", "Este mes", o "Todo". Todos los totales y la gráfica se recalculan al instante según el rango elegido.
- **Nueva tabla "Detalle día por día"**: además del acumulado por tramitador, ahora se ve una tabla con cada fecha y cada persona por separado — para saber exactamente qué hizo cada quien cada día, no solo el total general.

- **Acceso de solo lectura a reportes para Ejecutivo**: nuevo botón "📋 Ver reportes (solo lectura)" en la pantalla principal del Ejecutivo. Puede ver el dashboard, entrar al detalle de cualquier reporte, y consultar todo — pero no ve el botón de "Nuevo reporte", ni "Editar", ni ningún botón de borrar. Protegido tanto en la interfaz como en el código (por si alguien intentara forzarlo).

- **Corregido bug crítico: no se podía guardar al editar el único reporte del día.** Al editar un reporte que era la única hoja capturada ese día (muy común — normalmente hay un reporte por persona por día) y guardarlo sin cambiar la fecha, el sistema borraba el día completo por quedar vacío al quitar la versión anterior, y luego intentaba volver a agregarle la hoja editada — a algo que él mismo acababa de borrar. Esto rompía el guardado para Trámite (y cualquier rol) al editar sus propios reportes en el caso más común. Ya corregido: ahora reemplaza el renglón en su lugar en vez de borrar y volver a agregar.

- **Corregido: el logo no aparecía en GitHub Pages.** El logo y los íconos usaban rutas absolutas (`/logo-ow.jpg`), que funcionan bien en Vercel (que sirve desde la raíz del dominio) pero no en GitHub Pages, que sirve la app desde una subcarpeta (`tuusuario.github.io/reporteoperaciones/`) — con ruta absoluta, el navegador buscaba el logo en la raíz del dominio en vez de dentro de esa subcarpeta. Ya corregido a rutas relativas, que funcionan igual de bien en ambos.

- **Corregido el bug real de "Listo" no movía el pendiente.** Había un error de paréntesis en el código: al contar cuántas veces se repetía una Guía en el historial completo (para poner "2do previo", etc.), el respaldo contra reportes con datos faltantes solo cubría el caso "despacho", no "previo" — así que si CUALQUIER reporte antiguo en tu historial tenía datos con una forma distinta, la función truena a media ejecución, justo antes de mover el renglón. Ya corregido para ambos casos.
- Además, tanto "Traer asignaciones" como "Marcar Listo" ahora muestran un mensaje de error específico en pantalla si algo más llegara a fallar, en vez de quedarse en silencio.

- **Corregido: el botón "✅ Listo" no aparecía al elegir el Tipo manualmente.** Si en un renglón de Pendientes elegías "Previo" o "Despacho" desde el desplegable de Tipo (en vez de que llegara ya lleno por una asignación), la pantalla no se refrescaba sola para mostrar el botón "Listo" — el dato sí se guardaba, pero visualmente no aparecía hasta hacer otra acción. Ya corregido: el selector ahora refresca la pantalla de inmediato al elegir una opción.

- **Botón "📥 Asignaciones" siempre visible para Trámite**: en vez de depender solo de la notificación (que puede pasar desapercibida), ahora aparece un botón en rojo claro en la barra superior con el número de asignaciones pendientes, visible en cualquier pantalla mientras tengas alguna. Al tocarlo se abre un panel de solo consulta con Guía, Cliente, Almacén y quién la asignó — no crea ni modifica nada, solo informa. Para agregarlas al reporte se sigue usando el botón "Traer asignaciones pendientes" dentro de la captura, como ya funcionaba.

- **Corregido: las notificaciones no llegaban al cerrar sesión y volver a entrar.** El sistema comparaba "qué es nuevo" solo en memoria, dentro de la misma sesión abierta — al recargar la página o volver a entrar más tarde, perdía ese rastro por completo. Ahora el rastro se guarda en el dispositivo (localStorage), así que las notificaciones también funcionan al iniciar sesión de nuevo, no solo mientras la app se queda abierta todo el tiempo.

- **Corregido: "Traer asignaciones" no hacía nada al editar reportes antiguos.** Si el reporte se había creado en una versión anterior de la app (antes de que existiera la sección "Pendientes"), le faltaba ese campo internamente — y el botón fallaba en silencio, sin ningún aviso. Ahora cualquier reporte (nuevo o viejo) se completa automáticamente con los campos que le falten antes de poder editarse.
- **Red de seguridad nueva**: si algún botón llegara a fallar por un error de programación no previsto, ahora se muestra un aviso en pantalla en vez de "no pasar nada" — para que siempre sepas si algo salió mal.

- **Corregido: las asignaciones no llegaban si el tramitador estaba a media captura.** Antes, mientras alguien estaba llenando un reporte, la app dejaba de refrescar las asignaciones hasta que saliera de esa pantalla — por eso una asignación nueva "no llegaba". Ahora las asignaciones se refrescan siempre, sin importar en qué pantalla esté cada quien.
- **Sistema de notificaciones**: 🔔 Trámite recibe aviso cuando le cae una asignación nueva; Administrador/Coordinación reciben aviso cuando una asignación se completa. Hay dos formas en que se muestran:
  1. **Notificación real del navegador/celular** (fuera de la pestaña) — hay que activarla una vez con el botón "Activar" que aparece en el inicio.
  2. **Campana 🔔 en la barra superior** — funciona siempre, sin necesitar permiso, muestra las últimas notificaciones dentro de la app.
- ⚠️ **Límite honesto**: esto funciona mientras la app esté abierta (aunque sea en otra pestaña o minimizada). Si alguien cierra la app por completo, no le va a llegar nada hasta que la vuelva a abrir — para notificaciones que lleguen con la app cerrada haría falta un sistema más grande (Firebase Cloud Messaging con su propio servidor), que es un proyecto aparte si algún día lo necesitan.

- **Corregido bug grave: los reportes borrados "resucitaban" solos.** Al borrar el último reporte que quedaba en una aduana, el sistema de migración interpretaba la lista vacía como "esta aduana nunca se ha usado" y volvía a copiar los datos viejos de antes de separar por aduanas — trayendo de vuelta lo que se acababa de borrar. Ahora usa una bandera permanente que se marca una sola vez y nunca se vuelve a evaluar, sin importar si después la lista de reportes queda vacía por un borrado legítimo.

- **Solo Administrador/Coordinación pueden borrar reportes ya capturados.** Trámite y Ejecutivo ya no ven el botón de eliminar (ni hoja individual ni día completo), y aunque lo intentaran por fuera de la interfaz, el sistema lo rechaza igual.
- **Actualización automática en segundo plano**: la app ahora refresca los reportes y asignaciones cada 25 segundos (y también al volver a "Inicio"), para que si alguien más borra o cambia algo, se vea reflejado sin tener que cerrar sesión y volver a entrar. No interrumpe mientras estás capturando un reporte, y respeta cualquier cambio tuyo que todavía no se haya subido.

- **Corregido**: la pantalla de Catálogos truena con "Cannot read properties of undefined (reading 'map')" si el catálogo guardado en Firebase venía de una versión anterior sin la clave "tramitadores" (u otra). Ahora cualquier catálogo cargado se completa automáticamente con las claves que falten antes de mostrarse.

- **Las asignaciones ahora caen todas a Pendientes primero**, con un campo "Tipo" (Previo/Despacho). Cada renglón de Pendientes con Tipo asignado muestra un botón **"✅ Listo"** que lo sube directo al cuadro de Previos o Despachos correspondiente y lo quita de Pendientes — sin duplicar información.
- **Nueva columna "N° previo" / "N° despacho"**: cuando se marca "Listo" y esa Guía ya tenía un previo/despacho anterior (algo normal — puede haber varios previos para la misma guía), se etiqueta automáticamente como "2do previo", "3er previo", etc. en esta columna nueva, sin bloquear ni avisar como error.
- **"Tipo de previo" ahora es un selector de dificultad (A/B/C/D)**, en vez de texto libre — refleja cómo realmente se clasifican los previos.
- El aviso de guía duplicada ya no se dispara para renglones que ya traen su "N° previo/despacho" etiquetado (porque ya se identificaron a propósito como repetición legítima).

- **Las asignaciones se traen solas al reporte**: en el formulario de captura, si el tramitador tiene asignaciones pendientes (de Previos o Despachos), aparece un botón "📥 Traer asignaciones pendientes (N)" — al presionarlo, agrega los renglones con **Guía, Cliente y Almacén ya llenos**, dejando Referencia/Tipo/Resultado en blanco para completar manualmente. Las guías que no tengan asignación se siguen agregando a mano con "+ Agregar renglón", como siempre. Al guardar, esos renglones cierran su asignación automáticamente (igual que antes).
- Corrección menor: "+ Agregar pendiente" ahora crea el renglón con los campos correctos de Pendientes (antes usaba por error los mismos campos que Despachos).

- **Tramitadores editable desde la app**: en "Catálogos" (solo Administrador/Coordinación) ya se puede agregar, quitar, o corregir el nombre de cualquier tramitador — se refleja de inmediato en la pantalla de login, sin tocar código.
- **Contraseña automática por tramitador**: cada quien tiene su propia contraseña para entrar como Trámite — se genera sola con las 3 primeras letras de su nombre + "2026" (ej. Mariana Carrillo → `mar2026`). Si dos nombres coinciden en esas 3 letras, automáticamente se usa una letra más para la segunda persona, así nunca se repite sin que el administrador tenga que hacer nada. Las contraseñas se ven junto a cada nombre en "Catálogos" para que el administrador las pueda comunicar. ⚠️ Igual que la contraseña de Administrador, esto es un filtro básico (vive en el código del navegador), no seguridad de nivel bancario — pero cumple bien para evitar que alguien entre por error a la cuenta de otro compañero.

- **Buscador global** (botón "🔍 Buscar" en la barra superior): busca por Guía, Referencia, Pedimento o Cliente en todo el historial de la aduana activa, con acceso directo a cada reporte encontrado.
- **Aviso de guía duplicada**: al guardar un reporte, si alguna Guía de previos/despachos ya existe en otro reporte, se muestra una advertencia con el detalle antes de guardar (se puede continuar si es intencional).
- **Asignaciones atrasadas resaltadas**: las asignaciones pendientes con 2+ días sin completarse se marcan en rojo intenso con "⚠️ Atrasada (N días)", y el botón "Asignaciones" en la barra superior muestra un contador si hay atrasadas.
- **Historial de cambios por reporte**: cada hoja capturada guarda quién la creó y quién la editó después, con fecha y hora — visible al final del detalle de cada reporte.
- **Respaldo completo en JSON** (botón "💾 Respaldo completo" en el inicio, solo Administrador/Coordinación): descarga un archivo con absolutamente todo (reportes, asignaciones, catálogos) de la aduana activa — un seguro extra además del Excel.

- **"Completar referencias" agrupado por cliente**: ahora el Ejecutivo ve una lista de clientes con pendientes (colapsada), y al tocar un cliente se abren sus guías pendientes de referencia — más fácil de navegar cuando hay muchas.
- **Colores de estatus**: rojo claro para "pendiente" y verde para "completada", tanto en Asignaciones como en Completar referencias.

- **Nuevos roles: Coordinación y Ejecutivo**. El login ahora tiene 4 opciones:
  - **Administrador** y **Coordinación**: misma contraseña (`ow2026`), mismos accesos completos.
  - **Trámite**: como antes, elige tu nombre de la lista de 6 personas.
  - **Ejecutivo**: escribe tu nombre libremente (no hay lista fija, cualquiera puede entrar).
- **Asignaciones** (pestaña nueva, solo Administrador/Coordinación): el coordinador crea asignaciones de previos/despachos por número de **Guía**, indicando a qué tramitador se le asigna. Cuando ese tramitador guarda un reporte con un renglón de esa misma Guía, la asignación **se cierra sola automáticamente** — sin que nadie tenga que marcarla a mano. Se ve en dos listas: Pendientes y Completadas.
- **Completar referencias** (pantalla principal del rol Ejecutivo): junta todos los renglones de cualquier reporte (previos, despachos, pendientes) que se quedaron sin número de Referencia, agrupados por Guía. El ejecutivo escribe la referencia correcta una sola vez, y se actualiza automáticamente en **todos** los renglones de **todos** los reportes que compartan esa misma Guía.

- **Dashboard de KPIs por tramitador**: nueva pestaña "KPIs" en la barra superior (solo visible para Administrador). Muestra, por cada persona: total de previos, despachos, pendientes, suma total, número de reportes capturados, y días distintos con actividad — más una gráfica comparativa de barras. Se calcula automáticamente con todo el historial de la aduana activa.
- **Aduana del reporte corregida**: el campo Aduana ahora siempre se autocompleta con la aduana elegida al iniciar sesión (igual que Tramitador para el rol Trámite) y queda bloqueado — antes se quedaba vacío.

- **Orden del login corregido**: ahora se elige primero la aduana ("¿Con qué aduana vas a trabajar?"), y después el rol (Administrador/Trámite) — como debía ser desde el inicio.
- **"Salir" corregido**: antes no borraba la sesión guardada, así que recargar la página después de cerrar sesión te volvía a meter solo. Ya queda una salida real.

- **Separación completa por aduana**: después de iniciar sesión, se elige con qué aduana se va a trabajar (por ahora solo "GDL", pero está listo para agregar más). Cada aduana tiene su propio historial, catálogos, y Excel compartido — completamente independientes entre sí en Firebase (`aduanas/{aduana}/...`). Los datos de prueba capturados antes de esta versión se migran automáticamente a "GDL" la primera vez que se usa (no se pierden, y tampoco se borran del lugar original).
- **Nueva sección "Pendientes"**: cada reporte ahora tiene una tercera área (entre Despachos y Otras Actividades) con los mismos campos que Despachos pero sin Tipo ni Resultado — para anotar operaciones que la persona de trámite deja pendientes ese día. Aparece como hoja propia en el Excel (con quién la dejó y en qué fecha) y en el CSV.
- **Para agregar más aduanas en el futuro**: edita la constante `ADUANAS` en `src/main.js` (por ahora `["GDL"]`) y agrega el nombre/siglas que necesites — el resto (separación de datos, catálogos, Excel) funciona automáticamente para cada aduana nueva sin más cambios de código.

- **Contraseña para Administrador**: al elegir "Administrador" en el login, ahora pide una contraseña (`ow2026`, definida en `src/main.js` como `ADMIN_PASSWORD` — puedes cambiarla ahí cuando quieras). ⚠️ Importante: como esta app no tiene servidor propio, esta contraseña vive en el código del navegador — es un filtro simple para evitar accesos accidentales o casuales, no una protección de nivel bancario contra alguien con conocimientos técnicos que revise el código fuente.

- **Usuarios individuales para "Trámite"**: al elegir "Trámite" en el login, se abre una segunda pantalla para elegir la persona específica (Monica Ortega, Luis Arreola, Mariana Carrillo, Mayra Romero, Javier Garcia, Julio Regalado). Esa selección queda guardada como "Capturado por" en cada reporte automáticamente, y también llena solo el campo "Tramitador" del reporte (bloqueado para edición, ya que es automático). "Administrador" sigue siendo un solo acceso general.
- **Excel restringido a Administrador**: el rol "Trámite" ya no ve los botones de descargar Excel (ni por día ni el histórico completo) — solo Administrador puede descargarlos. (Nota honesta: el link especial `?historico=1` no pasa por este control de rol, ya que no requiere iniciar sesión; si alguien tiene ese link exacto puede descargar el Excel sin importar su rol. Avísame si quieres que también se restrinja.)

- **Usuarios por tramitador**: al elegir "Trámite" en el login, aparece una segunda pantalla para elegir quién eres (Monica Ortega, Luis Arreola, Mariana Carrillo, Mayra Romero, Javier Garcia, Julio Regalado) — ese nombre queda registrado automáticamente como "capturado por" en cada reporte que subas, sin tener que escribirlo. Editables en `TRAMITADORES` dentro de `src/main.js`.
- **Trámite no puede descargar Excel**: los botones de Excel (link fijo, histórico completo, y por día) solo se muestran para "Administrador". Nota: el link especial `?historico=1` sigue funcionando sin iniciar sesión (para compartir con gente externa como el contador) — si alguien de Trámite llegara a tener ese link exacto, técnicamente podría abrirlo, ya que fue diseñado para no requerir cuenta. Avísame si prefieres que ese link también exija sesión de Administrador.
- **Funciona 100% sin conexión (lectura y escritura)**: la app guarda una copia de todos los reportes y catálogos en el teléfono cada vez que se conecta. Si abres la app sin señal, ves esa copia (con un aviso "📴 Mostrando la última copia guardada en este dispositivo"), puedes seguir capturando/editando/borrando con total normalidad, y todo se sincroniza solo en cuanto regresa la señal — incluyendo el Excel compartido y el link de descarga automática. Único límite honesto: si la otra persona hizo cambios desde su teléfono mientras el tuyo estuvo sin señal, no los verás hasta que ambos tengan conexión al mismo tiempo (no hay forma de que un teléfono sin internet reciba datos de otro).
- **Guardado confiable sin señal**: si guardas, editas o borras un reporte sin conexión, la app lo guarda en el teléfono (no lo pierde) y lo marca como "⏳ pendiente". En cuanto detecta que regresó la señal (o cada 20 segundos como respaldo), sube automáticamente todo lo pendiente a Firebase, refresca los datos con lo último del servidor, y actualiza el Excel compartido. La barra superior siempre muestra el estado: 🔴 Sin conexión, 🔄 Sincronizando…, o ⏳ N pendientes.
- **Instalable como app de celular (PWA)**: la app ahora se puede "instalar" en el teléfono, con ícono propio en la pantalla de inicio y sin la barra del navegador (pantalla completa, como una app nativa). En Android/Chrome aparece un botón "Instalar" automático; en iPhone se muestra un aviso con instrucciones (Compartir → Agregar a pantalla de inicio, ya que Apple no permite instalación automática desde el navegador).
- **Identidad visual actualizada**: logo real de la empresa (`public/logo-ow.jpg`) en el encabezado, login, e íconos de instalación; paleta de colores en tonos de azul y blanco.
- **Catálogo de aduanas**: la aduana ahora se sugiere con autocompletado igual que clientes/almacenes/tramitadores, editable en "Catálogos".
- **Nombres de usuario actualizados**: "Administrador" y "Trámite" (antes Davis / Compañero(a)).
- **Uso simultáneo confirmado**: varias personas pueden usar la app al mismo tiempo desde distintos dispositivos — cada una se conecta a la misma Realtime Database. Si dos personas editan exactamente el mismo reporte al mismo tiempo, gana el último que guarde (no se combinan cambios); para reportes distintos no hay ningún conflicto.
- **Link fijo de Excel siempre actualizado — sin costo**: cada vez que guardas, editas o borras un reporte, la app regenera el Excel histórico completo y lo guarda dentro de Realtime Database (NO usa Firebase Storage, que exige plan de pago). En la pantalla de inicio aparece "🔗 Link del Excel siempre actualizado" — ábrelo cuando quieras (funciona sin iniciar sesión) y descarga automáticamente la versión más reciente.
- **Captura manual**: si no tienes el Word ni el PDF a la mano, el botón "Llenar manualmente" abre el mismo formulario en blanco para capturar el reporte directo en la app.
- **Descarga en Excel**: además del link fijo, puedes descargar en cualquier momento un `.xlsx` con 3 hojas (Resumen, Previos, Despachos) — por día individual o del histórico completo desde el inicio.
- **Editar y eliminar reportes ya guardados**: corregido — ahora abre correctamente el formulario con los datos existentes, y hay un botón para borrar el reporte completo de un día directamente desde la pantalla de inicio.
- **Formato simplificado**: el encabezado ahora es solo Fecha + Tramitador + Aduana. Las tablas de previos y despachos ya no tienen columnas de horario.
- **Lector de PDF corregido**: se arregló un error que desalineaba columnas cuando el nombre de una columna se partía en varias palabras o líneas (ej. "Tipo de previo", "Tipo despacho").
- **Manejo de errores visible**: si algo falla, se muestra el motivo en pantalla en vez de quedarse "congelado" sin explicación.

## Notas importantes

- **Formato del documento**: la lectura automática espera exactamente 4 tablas en este orden: encabezado, "1. CONTROL DE PREVIOS", "2. CONTROL DE DESPACHOS", y la tabla de firmas (Elaboró/Revisó/Autorizó), igual que la plantilla original de Word. Si cambia el diseño, la lectura puede desalinearse — siempre revisen los datos en la pantalla de confirmación antes de guardar.
- **Word es más confiable que PDF**: Word usa la estructura real de tablas del documento (exacta). PDF reconstruye columnas por posición de texto — revisen con más cuidado los reportes cargados como PDF.
- **PDFs escaneados (imágenes) no funcionan**: si el PDF es una foto o escaneo sin texto real seleccionable, no se puede leer sin OCR/IA (que tiene costo). Usen el Word original o un PDF exportado directamente desde Word.
- **Sin usuarios ni contraseñas**: al entrar se elige un nombre de una lista fija (`const USERS` en `src/main.js`). Editen esa lista si cambian los nombres.
