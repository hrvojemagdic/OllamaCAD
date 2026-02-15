@echo off
setlocal EnableExtensions

REM --------------------------------------------
REM OllamaCAD - SOLIDWORKS Add-in Uninstaller
REM Unregisters COM DLL (x64) + removes registry keys
REM --------------------------------------------

set "SCRIPT_DIR=%~dp0"
set "DLL=%SCRIPT_DIR%OllamaCAD.dll"

REM Add-in GUID (must match your AddIn.cs Guid attribute)
set "ADDIN_GUID={D21CDAF8-30C5-46DE-9B44-2386572E1D43}"

echo.
echo === OllamaCAD Add-in Uninstall ===
echo Folder: %SCRIPT_DIR%
echo.

REM Check DLL exists (best-effort uninstall even if missing)
if not exist "%DLL%" (
  echo WARNING: DLL not found at: %DLL%
  echo Will still attempt to remove registry keys.
  echo.
)

REM Locate 64-bit regasm
set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
if exist "%REGASM%" (
  if exist "%DLL%" (
    echo [1/3] Unregistering COM (regasm /unregister)...
    "%REGASM%" /unregister "%DLL%"
    if errorlevel 1 (
      echo WARNING: regasm unregister failed (may require Administrator).
    )
  ) else (
    echo [1/3] Skipping regasm unregister (DLL missing).
  )
) else (
  echo WARNING: regasm not found at: %REGASM%
  echo Skipping COM unregister step.
)

echo.
echo [2/3] Removing SolidWorks add-in registry key (HKLM)...
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\SolidWorks\Addins\%ADDIN_GUID%" /f >nul 2>&1
if errorlevel 1 (
  echo WARNING: Could not delete HKLM key (Administrator rights likely required).
) else (
  echo OK: HKLM key removed.
)

echo.
echo [3/3] Removing add-in startup key (HKCU)...
reg delete "HKEY_CURRENT_USER\SOFTWARE\SolidWorks\AddInsStartup\%ADDIN_GUID%" /f >nul 2>&1
if errorlevel 1 (
  echo WARNING: Could not delete HKCU startup key (may not exist).
) else (
  echo OK: HKCU startup key removed.
)

echo.
echo Uninstall complete.
echo - Restart SOLIDWORKS if it is currently running.
echo.
pause
exit /b 0
