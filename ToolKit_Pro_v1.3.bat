@echo off
setlocal enabledelayedexpansion
:: ---------------------------------------------------------
:: NOMBRE:      Panel de Herramientas Pro
:: VERSION:     1.3.0
:: AUTOR:       %USERNAME%
:: FECHA:       09/02/2026
:: ---------------------------------------------------------
title Panel de Herramientas Pro v1.3.0 - %USERNAME%

:: Definir codigos de colores ANSI
set "esc="
set "G=%esc%[92m"  :: Verde (Exito/Menu)
set "Y=%esc%[93m"  :: Amarillo (Advertencia/Proceso)
set "R=%esc%[91m"  :: Rojo (Error)
set "W=%esc%[0m"   :: Blanco/Reset

:: Verificar privilegios de entrada
net session >nul 2>&1
if %errorLevel% == 0 ( set "admin=SI" ) else ( set "admin=NO" )

:menu
cls
echo %G%======================================================
echo           PANEL DE CONTROL v1.3.0
echo ======================================================
echo  1. Reiniciar Outlook (Cerrar y Abrir)
echo  2. Ver Estado de Bateria (Salud y Porcentaje)
echo  3. Ver Informacion del Computador (Hardware)
echo  4. Generar Informe de Bateria HTML (Escritorio)
echo  5. Limpiar Temporales (Limpieza Profunda)
echo  6. Reparar Audio (HP EliteBook/Realtek)
echo  7. Salir
echo ======================================================%W%

if %admin%==NO (
    echo %Y%[!] AVISO: Ejecute como Administrador para Audio y Limpieza.%W%
) else (
    echo %G%[OK] Modo Administrador detectado.%W%
)
echo %G%======================================================%W%
set /p opt="Seleccione una opcion (1-7): "

if "%opt%"=="1" goto outlook
if "%opt%"=="2" goto bateria
if "%opt%"=="3" goto hardware
if "%opt%"=="4" goto informe
if "%opt%"=="5" goto limpieza
if "%opt%"=="6" goto fixaudio
if "%opt%"=="7" exit
goto menu

:outlook
cls
echo %G%[>] Cerrando Outlook...%W%
taskkill /F /IM outlook.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Outlook...%W%
start outlook:
pause
goto menu

:bateria
cls
echo %G%--- ESTADO DE LA BATERIA ---%W%
powershell -Command "$b=Get-CimInstance Win32_Battery; $f=Get-CimInstance -Namespace root/WMI -ClassName BatteryFullChargedCapacity; $s=Get-CimInstance -Namespace root/WMI -ClassName BatteryStaticData; if($b){ Write-Host 'Modelo: ' $b.Name; Write-Host 'Carga Actual: ' $b.EstimatedChargeRemaining '%%'; if($s.DesignedCapacity -gt 0){ $h=[math]::Round(($f.FullChargedCapacity/$s.DesignedCapacity)*100,2); Write-Host 'Salud de Vida: ' $h '%%' -ForegroundColor Green } } else { Write-Host 'ERROR: No se detecto bateria.' -ForegroundColor Red }"
echo.
pause
goto menu

:hardware
cls
echo %G%--- INFORMACION DE HARDWARE ---%W%
powershell -Command "Write-Host 'CPU: ' (Get-CimInstance Win32_Processor).Name"
powershell -Command "Write-Host 'RAM: ' ([math]::round((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)) 'GB'"
powershell -Command "Write-Host 'GPU: ' (Get-CimInstance Win32_VideoController).Name"
for /f "tokens=2 delims==" %%a in ('wmic bios get serialnumber /value') do echo S/N: %%a
echo.
pause
goto menu

:informe
cls
echo %G%[>] Generando reporte...%W%
powercfg /batteryreport /output "%USERPROFILE%\Desktop\Reporte_Bateria.html" >nul 2>&1
if %errorLevel% == 0 (
    echo %G%[OK] Reporte creado en el escritorio.%W%
    start "" "%USERPROFILE%\Desktop\Reporte_Bateria.html"
) else (
    echo %R%[!] Error al generar el reporte de bateria.%W%
)
pause
goto menu

:limpieza
cls
echo %G%--- INICIANDO LIMPIEZA ---%W%
echo %Y%(Nota: Los archivos en uso no se pueden borrar)%W%
echo.
echo [>] Limpiando Temp de Usuario...
del /q /s /f "%temp%\*" >nul 2>&1
for /d %%x in ("%temp%\*") do rd /s /q "%%x" >nul 2>&1

if %admin%==SI (
    echo [>] Limpiando Temp de Windows...
    del /q /s /f "C:\Windows\Temp\*" >nul 2>&1
    for /d %%x in ("C:\Windows\Temp\*") do rd /s /q "%%x" >nul 2>&1
    echo [>] Limpiando Prefetch...
    del /q /s /f "C:\Windows\Prefetch\*" >nul 2>&1
    echo %G%[OK] Limpieza de sistema completada.%W%
) else (
    echo %R%[!] Saltando carpetas de sistema (Requiere Admin).%W%
)
echo.
pause
goto menu

:fixaudio
cls
echo %G%--- REPARACION DE AUDIO (HP/REALTEK) ---%W%
if %admin%==NO (
    echo %R%[ERROR] Necesitas permisos de Administrador para reiniciar servicios.%W%
    echo Por favor, cierre y vuelva a ejecutar como Administrador.
    echo.
    pause
    goto menu
)

echo %Y%[1/4] Deteniendo Audio Services...%W%
net stop "Audiosrv" /y >nul 2>&1
net stop "AudioEndpointBuilder" /y >nul 2>&1
timeout /t 2 >nul

echo %Y%[2/4] Reiniciando Audio Services...%W%
net start "AudioEndpointBuilder" >nul 2>&1
net start "Audiosrv" >nul 2>&1

echo %Y%[3/4] Verificando estado...%W%
sc query "Audiosrv" | find "RUNNING" >nul
if %errorLevel% equ 0 (
    echo %G%[EXITO] Servicios de audio restablecidos.%W%
) else (
    echo %R%[FALLO] Reinicie el equipo.%W%
)
echo.
pause
goto menu
