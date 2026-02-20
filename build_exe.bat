@echo off
setlocal ENABLEDELAYEDEXPANSION

REM ==========================
REM Configuración
REM ==========================
set APP_NAME=Office_files_Cypher
set ENTRY_POINT=src/main.py
set DIST_DIR=dist
set BUILD_DIR=build

echo.
echo ==========================
echo Building %APP_NAME%
echo ==========================
echo.

REM Comprobar que estamos en el venv correcto
where python >nul 2>nul
if errorlevel 1 (
  echo ERROR: No se encuentra "python" en PATH. Activa el venv antes de ejecutar.
  pause
  exit /b 1
)

REM Mostrar python activo (debug útil)
python -c "import sys; print('Python:', sys.executable)"

REM Asegurar pip disponible
python -m pip --version >nul 2>nul
if errorlevel 1 (
  echo ERROR: pip no disponible en este entorno.
  pause
  exit /b 1
)

REM Comprobar/instalar PyInstaller
python -c "import PyInstaller" >nul 2>nul
if errorlevel 1 (
  echo PyInstaller no esta instalado. Instalando...
  python -m pip install --upgrade pip
  python -m pip install pyinstaller
)

REM Limpieza previa
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%APP_NAME%.spec" del "%APP_NAME%.spec"

REM Build (usar python -m PyInstaller para no depender del comando pyinstaller)
python -m PyInstaller ^
  --name "%APP_NAME%" ^
  --onefile ^
  --windowed ^
  --clean ^
  --noconfirm ^
  --add-data "i18n;i18n" ^
  "%ENTRY_POINT%"

REM Resultado
if exist "%DIST_DIR%\%APP_NAME%.exe" (
    echo.
    echo ==========================
    echo Build completado OK
    echo EXE: %DIST_DIR%\%APP_NAME%.exe
    echo ==========================
) else (
    echo.
    echo ==========================
    echo ERROR: No se genero el EXE
    echo Revisa el log anterior
    echo ==========================
)

pause
