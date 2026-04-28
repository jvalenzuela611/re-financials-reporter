# RE Financials Reporter — Instrucciones para Florencia

## ¿Qué es esto?
Aplicación web local que genera reportes financieros de activos inmobiliarios (Notes to Financials, Master Excel, Management Commentary) a partir de los archivos contables exportados de Appfolio.

---

## Requisitos previos (una sola vez)

### 1. Instalar Python 3.11
- Ir a: https://www.python.org/downloads/
- Descargar **Python 3.11.x** (o superior)
- Durante la instalación: **marcar la casilla "Add Python to PATH"** ✅
- Completar la instalación

### 2. Instalar Claude Code (para poder pedirle ayuda a Claude)
- Ir a: https://claude.ai/download
- Descargar e instalar Claude para Desktop
- Iniciar sesión con tu cuenta de Anthropic

---

## Instalación de la app (una sola vez)

1. Descargar la carpeta `RE_Financials_Reporter` y colocarla en el Escritorio
2. Hacer **doble clic** en `instalar.bat`
3. Se abrirá una ventana negra que instalará las dependencias automáticamente
4. Esperar hasta que diga **"Instalación completada"** y presionar cualquier tecla para cerrar

---

## Cómo usar la app (cada vez)

1. Hacer **doble clic** en `iniciar_app.bat`
2. Se abrirá una ventana negra — esperar unos segundos
3. El navegador abrirá automáticamente **http://localhost:8501**
4. Usar la app normalmente
5. Para cerrar: cerrar la ventana negra (o presionar `Ctrl+C` dentro de ella)

---

## Flujo de la app

```
Paso 1 → Subir archivos contables (.xlsm de Appfolio)
Paso 2 → Revisar y clasificar líneas contables
Paso 3 → Generar reportes (Notes to Financials, Master Excel)
```

---

## Claves API necesarias

La app usa modelos de IA (OpenAI o Anthropic). En el **Paso 3** deberás ingresar tu clave API:

- **OpenAI:** https://platform.openai.com/api-keys
- **Anthropic:** https://console.anthropic.com/settings/keys

La clave se ingresa directamente en la app — no se guarda en ningún archivo.

---

## Si algo no funciona

Abrí una sesión de Claude Code, arrastrá la carpeta del proyecto y decile:
> "La app no está corriendo. ¿Podés ayudarme a levantarla?"

Claude leerá este archivo y te guiará paso a paso.

---

## Estructura de la carpeta

```
RE_Financials_Reporter/
├── app.py                  ← Aplicación principal
├── requirements.txt        ← Dependencias Python
├── instalar.bat            ← Instalación (correr una vez)
├── iniciar_app.bat         ← Iniciar la app
├── LEEME.md                ← Este archivo
├── masters/                ← Archivos Master Excel por activo
│   ├── master_alice_house_jv_llc_2025.xlsx
│   ├── master_portfolio_oakland_2025.xlsx
│   └── master_295_29th_street_jv_llc.xlsx
├── src/                    ← Código fuente
└── .streamlit/             ← Configuración visual (tema STARS)
```
