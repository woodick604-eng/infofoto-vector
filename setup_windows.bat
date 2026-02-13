@echo off
echo ===========================================
echo   Configuracio automatica per a Windows
echo ===========================================

where python >nul 2>nul
if errorlevel 1 (
    echo No s'ha trobat Python. Instal·la'l des de https://www.python.org/downloads/
    pause
    exit /b
)

if not exist .venv (
    echo Creant entorn virtual .venv...
    python -m venv .venv
) else (
    echo Ja existeix l'entorn virtual .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip wheel

set REQS=
if exist requirements.txt set REQS=requirements.txt
if exist requeriments.txt set REQS=requeriments.txt

if defined REQS (
    echo Instal·lant dependències des de %REQS% ...
    pip install -r %REQS%
) else (
    echo No s'ha trobat ni requirements.txt ni requeriments.txt. Instal·la-les manualment si cal.
)

echo Entorn preparat correctament. Executa start_windows.bat per arrencar.
pause
