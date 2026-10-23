@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "TARGET=%~1"

if "%TARGET%"=="" (
  echo.
  echo   Drag the exported .zip or .json file onto this window, then press Enter.
  echo   [ Drop file here ]
  echo.
  set /p TARGET=
)

if "%TARGET%"=="" (
  echo Nothing to do.
  pause
  exit /b 1
)

set "TARGET=%TARGET:"=%"

where python >nul 2>nul
if %errorlevel%==0 (
  set "PY=python"
) else (
  set "PY=py -3"
)

%PY% -c "import docx" >nul 2>nul
if errorlevel 1 (
  echo Installing python-docx ...
  %PY% -m pip install -q python-docx
  %PY% -c "import docx" >nul 2>nul
  if errorlevel 1 (
    echo.
    echo [ERROR] python-docx is missing. Run:  %PY% -m pip install python-docx
    pause
    exit /b 1
  )
)

set PYTHONIOENCODING=utf-8
echo.
echo Converting: "%TARGET%"
echo.
%PY% deepseek_archive.py convert "%TARGET%" --merge
echo.
pause
