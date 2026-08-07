# Reporte Operativo Diario — Oñate, Willy & Cía.

Aplicación web (Vite + Firebase Realtime Database) para digitalizar el "Reporte Operativo Diario de Previos y Despachos" a partir de archivos **Word (.docx)** o **PDF** ya llenados. No usa inteligencia artificial ni servicios de pago: toda la lectura del documento ocurre en el navegador.

## Estructura del proyecto

```
reporte-ow/
├── index.html            # HTML raíz (Vite)
├── package.json
├── vite.config.js
├── src/
│   ├── main.js            # Lógica de la app, UI y estado
│   ├── firebase.js         # Inicialización de Firebase (tu configuración real)
│   ├── storage.js          # Lectura/escritura del historial en Realtime Database
│   ├── parseDocx.js        # Lector gratuito de Word (.docx), sin IA
│   ├── parsePdf.js         # Lector gratuito de PDF, sin IA (best-effort)
│   ├── fields.js            # Definición de columnas de previos/despachos
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

## 3. Desplegar en Vercel (recomendado)

1. Sube este proyecto a un repositorio de GitHub.
2. Entra a [vercel.com](https://vercel.com) → "Add New… → Project" → selecciona el repositorio.
3. Vercel detecta Vite automáticamente (build command `vite build`, output `dist`) — no necesitas configurar nada.
4. "Deploy". Tu app quedará en una URL tipo `https://reporte-ow.vercel.app`.

## 4. Desplegar en GitHub Pages (alternativa)

1. `npm run build` genera la carpeta `dist/`.
2. La forma más simple es usar el paquete `gh-pages`:
   ```bash
   npm install -D gh-pages
   ```
   Agrega a `package.json` en `"scripts"`:
   ```json
   "deploy": "vite build && gh-pages -d dist"
   ```
   Luego:
   ```bash
   npm run deploy
   ```
3. En GitHub: **Settings → Pages → Source** → rama `gh-pages`, carpeta `/ (root)`.

## Notas importantes

- **Formato del documento**: la lectura automática espera exactamente 4 tablas en este orden: encabezado, "1. CONTROL DE PREVIOS", "2. CONTROL DE DESPACHOS", y la tabla de firmas (Elaboró/Revisó/Autorizó), igual que la plantilla original de Word. Si cambia el diseño, la lectura puede desalinearse — siempre revisen los datos en la pantalla de confirmación antes de guardar.
- **Word es más confiable que PDF**: Word usa la estructura real de tablas del documento (exacta). PDF reconstruye columnas por posición de texto — revisen con más cuidado los reportes cargados como PDF.
- **PDFs escaneados (imágenes) no funcionan**: si el PDF es una foto o escaneo sin texto real seleccionable, no se puede leer sin OCR/IA (que tiene costo). Usen el Word original o un PDF exportado directamente desde Word.
- **Sin usuarios ni contraseñas**: al entrar se elige un nombre de una lista fija (`const USERS` en `src/main.js`). Editen esa lista si cambian los nombres.
