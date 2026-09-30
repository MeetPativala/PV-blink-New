@echo off
setlocal
cd /d "%~dp0"
set "DASHBOARD_PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%DASHBOARD_PYTHON%" set "DASHBOARD_PYTHON=%~dp0venv\Scripts\python.exe"
if not exist "%DASHBOARD_PYTHON%" (
    echo Create the Python environment using the Local development steps in README.md.
    exit /b 1
)
"%DASHBOARD_PYTHON%" -m streamlit run "%~dp0pv_blink_dashboard.py" %*
endlocal
