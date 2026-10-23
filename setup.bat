@echo off
setlocal
cd /d "%~dp0"
echo Setting up PCGD...
if exist "venv\Scripts\python.exe" goto install
where py >nul 2>nul
if errorlevel 1 goto python_fallback
py -3 -m venv venv
if errorlevel 1 goto python_fallback
goto install
:python_fallback
python -m venv venv
if errorlevel 1 goto failed
:install
"venv\Scripts\python.exe" -c "import sys; assert sys.version_info >= (3, 10), 'Python 3.10 or newer is required'"
if errorlevel 1 goto failed
"venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
"venv\Scripts\python.exe" scripts\check_install.py
if errorlevel 1 goto failed
echo Setup complete. Open run.bat to start PCGD.
pause
exit /b 0
:failed
echo Setup failed. Read the error above. Install Python 3.10+ with Tcl/Tk and Microsoft Excel desktop.
pause
exit /b 1
