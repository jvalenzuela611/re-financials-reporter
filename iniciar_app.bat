@echo off
chcp 65001 >nul
echo.
echo =====================================================
echo   RE Financials Reporter — Iniciando aplicacion
echo =====================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado. Correr primero "instalar.bat"
    pause
    exit /b 1
)

:: Verificar Streamlit
python -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Streamlit no instalado. Correr primero "instalar.bat"
    pause
    exit /b 1
)

echo Iniciando la app...
echo.
echo La aplicacion se abrira en el navegador en:
echo   http://localhost:8501
echo.
echo Para cerrar la app: cerrar esta ventana o presionar Ctrl+C
echo.

:: Cambiar al directorio donde está el script (por si se ejecuta desde otro lado)
cd /d "%~dp0"

:: Lanzar Streamlit y abrir navegador automaticamente
python -m streamlit run app.py --server.port 8501 --server.headless false
