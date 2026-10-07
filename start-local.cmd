@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_EXE=python"
"%PYTHON_EXE%" -c "import openpyxl,pypdf" 2>nul
if errorlevel 1 set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
"%PYTHON_EXE%" -c "import openpyxl,pypdf" 2>nul
if errorlevel 1 (
  echo Missing Python packages. See README.md for offline installation.
  pause
  exit /b 1
)
"%PYTHON_EXE%" server.py --open
pause
