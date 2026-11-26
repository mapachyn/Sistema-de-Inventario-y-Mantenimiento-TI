@echo off
setlocal
cd /d "%~dp0"
fltmc >nul 2>&1
if errorlevel 1 (
    echo Se solicitaran permisos de administrador...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    if errorlevel 1 echo No se obtuvo elevacion. Ejecute manualmente como administrador.
    exit /b
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Inventario_Mantenimiento.ps1"
echo.
echo Presione una tecla para cerrar la ventana.
pause >nul
endlocal
