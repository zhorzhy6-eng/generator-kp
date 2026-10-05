@echo off
chcp 65001 >nul
rem ^^^ Keep this line: scripts\setup.py re-encodes its own stdout to UTF-8,
rem ^^^ so the console must be UTF-8 for its Russian text to render. This .bat
rem ^^^ holds ASCII text only, so cmd.exe can never garble it, on any locale.
setlocal

rem ============================================================
rem  Build the standalone EXE for the KP generator.
rem ============================================================
rem  This script only calls scripts\setup.py, which does the work:
rem  installs dependencies, protects the repo with a hook,
rem  builds the EXE from the spec files, copies it to the Desktop
rem  and checks that the program starts.
rem
rem  Any setup.py option can be passed to this file as-is, e.g.:
rem      build.bat --skip-template
rem      build.bat --build-only
rem ============================================================

echo ============================================================
echo   BUILDING THE KP GENERATOR
echo ============================================================
echo.

rem Make sure Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python was not found in PATH.
    echo.
    echo Install Python 3.9+ and tick the checkbox
    echo "Add Python to PATH" during setup:
    echo    https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Running the installer and the build...
echo.
python "%~dp0setup.py" %*
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo [DONE] Build completed successfully.
) else (
    echo [ERROR] Build failed with exit code %RC%.
    echo Scroll up and look for the lines marked [X].
)
echo.
pause
exit /b %RC%
