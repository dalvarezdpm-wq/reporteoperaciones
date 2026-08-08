# Reporte Operativo Diario — Oñate, Willy & Cía.

Aplicación web (Vite + Firebase Realtime Database) para digitalizar el "Reporte Operativo Diario de Previos y Despachos" a partir de archivos **Word (.docx)** o **PDF** ya llenados. No usa inteligencia artificial ni servicios de pago: toda la lectura del documento ocurre en el navegador.

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
