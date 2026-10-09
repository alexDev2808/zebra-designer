@echo off
setlocal
cd /d "%~dp0"
rem ---------------------------------------------------------------
rem  Genera dist\EtiquetasZebra\EtiquetasZebra.exe con PyInstaller
rem  y, si Inno Setup 6 esta instalado, el instalador en dist\.
rem ---------------------------------------------------------------

if not exist .venv\Scripts\python.exe (
    echo Creando entorno de Python...
    python -m venv .venv || goto :error
)

echo Instalando dependencias...
.venv\Scripts\python.exe -m pip install -q -r requirements.txt pyinstaller || goto :error

echo Generando ejecutable...
.venv\Scripts\pyinstaller.exe --noconfirm --clean --windowed ^
    --name EtiquetasZebra --icon assets\icono.ico app.py || goto :error
copy /y driver.json dist\EtiquetasZebra\ >nul || goto :error
echo.
echo Ejecutable listo: dist\EtiquetasZebra\EtiquetasZebra.exe

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" goto :noinno

echo Generando instalador...
"%ISCC%" /Q installer.iss || goto :error
echo Instalador listo en la carpeta dist\
goto :end

:noinno
echo.
echo Inno Setup 6 no esta instalado; se omitio el instalador.
echo Instalelo con:  winget install JRSoftware.InnoSetup
goto :end

:error
echo.
echo *** Ocurrio un error durante la compilacion. ***
pause
exit /b 1

:end
pause
