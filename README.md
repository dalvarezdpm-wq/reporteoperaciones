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

- **Captura manual**: si no tienes el Word ni el PDF a la mano, el botón "Llenar manualmente" abre el mismo formulario en blanco para capturar el reporte directo en la app.
- **Descarga en Excel**: cada día tiene un botón "Descargar Excel" (además de CSV) que genera un `.xlsx` con 3 hojas: Resumen, Previos y Despachos.
- **Editar reportes ya guardados**: corregido — ahora abre correctamente el formulario con los datos existentes, y si cambias la fecha al editar, mueve el reporte al día correcto en Firebase.
- **Formato simplificado**: el encabezado ahora es solo Fecha + Tramitador + Aduana (ya no se capturan Coordinador, horarios, ni firmas). Las tablas de previos y despachos ya no tienen columnas de horario (Inicio/Fin), ya que no se estaban usando.
- **Manejo de errores visible**: si algo falla al guardar, editar, o dibujar una pantalla, ahora se muestra el motivo en pantalla en vez de quedarse "congelado" sin explicación.

## Notas importantes

- **Formato del documento**: la lectura automática espera exactamente 4 tablas en este orden: encabezado, "1. CONTROL DE PREVIOS", "2. CONTROL DE DESPACHOS", y la tabla de firmas (Elaboró/Revisó/Autorizó), igual que la plantilla original de Word. Si cambia el diseño, la lectura puede desalinearse — siempre revisen los datos en la pantalla de confirmación antes de guardar.
- **Word es más confiable que PDF**: Word usa la estructura real de tablas del documento (exacta). PDF reconstruye columnas por posición de texto — revisen con más cuidado los reportes cargados como PDF.
- **PDFs escaneados (imágenes) no funcionan**: si el PDF es una foto o escaneo sin texto real seleccionable, no se puede leer sin OCR/IA (que tiene costo). Usen el Word original o un PDF exportado directamente desde Word.
- **Sin usuarios ni contraseñas**: al entrar se elige un nombre de una lista fija (`const USERS` en `src/main.js`). Editen esa lista si cambian los nombres.
