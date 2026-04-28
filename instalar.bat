@echo off
chcp 65001 >nul
echo.
echo =====================================================
echo   RE Financials Reporter — Instalacion de dependencias
echo =====================================================
echo.

:: Verificar que Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no está instalado o no está en el PATH.
    echo.
    echo Por favor instalar Python 3.11 desde:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANTE: Durante la instalacion marcar "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
python --version
echo.

:: Instalar dependencias
echo Instalando dependencias (puede tardar 1-2 minutos)...
echo.
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Hubo un problema instalando las dependencias.
    echo Verificar conexion a internet e intentar nuevamente.
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   Instalacion completada exitosamente!
echo =====================================================
echo.
echo Ahora podes correr la app haciendo doble clic en:
echo   iniciar_app.bat
echo.
pause
