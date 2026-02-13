@echo off
echo ===========================================
echo   Arrencant l'aplicacio Informe Fotografic
echo ===========================================

if not exist .venv (
    echo No existeix .venv. Executa primer setup_windows.bat
    pause
    exit /b
)

call .venv\Scripts\activate

for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5051') do taskkill /PID %%a /F >nul 2>nul

start "app" python app.py
timeout /t 2 >nul
start http://127.0.0.1:5051

echo Servidor engegat. Prem una tecla per tancar aquesta finestra.
pause >nul
