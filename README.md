# Reporte de Stock — Instrucciones de Deploy en Railway

## Qué hace esta herramienta
Andrea abre una página web, arrastra 3 archivos (fotos de stock, PDF planificado, Excel de ventas)
y descarga automáticamente el Excel con el reporte completo de control de stock.

---

## Paso 1 — Subir el código a GitHub

1. Creá una cuenta en https://github.com si no tenés
2. Creá un repositorio nuevo (ej: `reporte-stock-mrnut`), privado
3. Subí todos los archivos de esta carpeta al repo:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `static/index.html`

---

## Paso 2 — Crear el servicio en Railway

1. Entrá a https://railway.app y creá una cuenta (podés usar tu cuenta de GitHub)
2. Clic en **"New Project"** → **"Deploy from GitHub repo"**
3. Seleccioná el repositorio que creaste
4. Railway va a detectar automáticamente que es Python y lo va a deployar

---

## Paso 3 — Configurar la API Key de Anthropic (IMPORTANTE)

La herramienta usa Claude para leer las fotos de stock. Necesitás agregar tu API key:

1. En Railway, andá a tu proyecto → pestaña **"Variables"**
2. Agregá esta variable:
   - Nombre: `ANTHROPIC_API_KEY`
   - Valor: tu API key de Anthropic (la conseguís en https://console.anthropic.com)
3. Railway va a reiniciar el servicio automáticamente

---

## Paso 4 — Compartir el link con Andrea

1. En Railway, andá a tu proyecto → pestaña **"Settings"** → **"Domains"**
2. Copiá el dominio público (algo como `reporte-stock-mrnut.up.railway.app`)
3. Mandáselo a Andrea por WhatsApp — ella solo tiene que abrir ese link en el navegador

---

## Plan gratuito de Railway
- Incluye 500 horas/mes gratis (suficiente para uso moderado)
- Si el servicio está sin uso se "duerme" y tarda ~30 segundos en despertar la primera vez

---

## Estructura del proyecto
```
reporte_stock/
├── app.py              ← Backend Python (Flask)
├── requirements.txt    ← Dependencias
├── Procfile            ← Comando de arranque para Railway
└── static/
    └── index.html      ← Página que ve Andrea
```
