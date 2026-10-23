@echo off
setlocal
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" call setup.bat
if errorlevel 1 exit /b 1
"venv\Scripts\python.exe" scripts\check_install.py
if errorlevel 1 goto failed
"venv\Scripts\python.exe" main.py
if errorlevel 1 goto failed
exit /b 0
:failed
echo PCGD could not start. Read the error above. Run setup.bat to check dependencies.
pause
exit /b 1
