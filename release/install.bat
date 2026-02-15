@echo off
setlocal EnableExtensions

REM --------------------------------------------
REM OllamaCAD - SOLIDWORKS Add-in Installer
REM Registers COM DLL (x64) + imports SolidWorks add-in registry keys
REM --------------------------------------------

set "SCRIPT_DIR=%~dp0"
set "DLL=%SCRIPT_DIR%OllamaCAD.dll"
set "REG=%SCRIPT_DIR%OllamaCAD_Addin.reg"

echo.
echo === OllamaCAD Add-in Install ===
echo Folder: %SCRIPT_DIR%
echo.

if not exist "%DLL%" (
  echo ERROR: Missing DLL: %DLL%
  echo Make sure OllamaCAD.dll is in the same folder as install.bat
  echo.
  pause
  exit /b 1
)

if not exist "%REG%" (
  echo ERROR: Missing REG file: %REG%
  echo Make sure OllamaCAD_Addin.reg is in the same folder as install.bat
  echo.
  pause
  exit /b 1
)

REM Admin check (HKLM requires admin)
net session >nul 2>&1
if errorlevel 1 (
  echo ERROR: Please right-click install.bat and choose "Run as administrator".
  echo.
  pause
  exit /b 1
)

set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
if not exist "%REGASM%" (
  echo ERROR: regasm not found at: %REGASM%
  echo Please ensure .NET Framework 4.x is installed.
  echo.
  pause
  exit /b 1
)

echo [1/2] Registering COM (regasm /codebase)...
"%REGASM%" "%DLL%" /codebase
if errorlevel 1 (
  echo ERROR: regasm failed.
  echo Command: "%REGASM%" "%DLL%" /codebase
  echo.
  pause
  exit /b 1
)

echo.
echo [2/2] Importing SolidWorks add-in registry keys...
reg import "%REG%"
if errorlevel 1 (
  echo ERROR: reg import failed.
  echo Command: reg import "%REG%"
  echo.
  pause
  exit /b 1
)

echo.
echo SUCCESS: OllamaCAD installed.
echo - Start SOLIDWORKS
echo - Go to Tools ^> Add-Ins
echo - Enable "Ollama CAD"
echo.
pause
exit /b 0
