@echo off
setlocal
cd /d "%~dp0"
echo ===========================================
echo NovaTest Quality Lab - Windows setup
echo ===========================================
where py >nul 2>nul
if errorlevel 1 (
  echo Python launcher was not found.
  echo Install Python 3.12 from https://www.python.org/downloads/windows/
  echo IMPORTANT: tick "Add python.exe to PATH" in the installer.
  pause
  exit /b 1
)
py -3.12 --version
if errorlevel 1 (
  echo Python 3.12 was not found. Install Python 3.12 and retry.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  echo Creating project virtual environment...
  py -3.12 -m venv .venv
  if errorlevel 1 goto failed
)
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
if errorlevel 1 goto failed
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
python -m playwright install chromium
if errorlevel 1 goto failed
if not exist ".env" copy ".env.example" ".env" >nul
echo.
echo Setup complete. Next run run_app.bat
pause
exit /b 0
:failed
echo.
echo Setup failed. Copy the error shown above and ask for help.
pause
exit /b 1
