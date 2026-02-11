@echo off
setlocal enabledelayedexpansion
:: ---------------------------------------------------------
:: NOMBRE:      Panel de Herramientas Pro
:: VERSION:     1.5.0 (Office Suite Update)
:: AUTOR:       %USERNAME%
:: FECHA:       09/02/2026
:: ---------------------------------------------------------
title Panel de Herramientas Pro v1.5.0 - %USERNAME%

:: Asegurarnos de no estar en una carpeta temporal
cd /d "%USERPROFILE%"

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
echo           PANEL DE CONTROL v1.5.0
echo ======================================================
echo   -- REINICIAR OFIMATICA --
echo  1. Outlook
echo  2. Microsoft Word
echo  3. Microsoft Excel
echo  4. Microsoft Teams (New/Classic)
echo  5. Adobe Acrobat / Reader
echo  6. Google Chrome
echo.
echo   -- SISTEMA Y DIAGNOSTICO --
echo  7. Ver Estado de Bateria
echo  8. Ver Informacion Hardware
echo  9. Generar Informe Bateria HTML
echo  10. Limpieza Temporales
echo  11. Reparar Audio (HP EliteBook)
echo  0. Salir
echo ======================================================%W%

if %admin%==NO (
    echo %Y%[!] AVISO: Ejecute como Administrador para Audio y Limpieza.%W%
) else (
    echo %G%[OK] Modo Administrador detectado.%W%
)
echo %G%======================================================%W%
set /p opt="Seleccione una opcion (0-11): "

:: Opciones de Aplicaciones
if "%opt%"=="1" goto outlook
if "%opt%"=="2" goto word
if "%opt%"=="3" goto excel
if "%opt%"=="4" goto teams
if "%opt%"=="5" goto acrobat
if "%opt%"=="6" goto chrome

:: Opciones de Sistema
if "%opt%"=="7" goto bateria
if "%opt%"=="8" goto hardware
if "%opt%"=="9" goto informe
if "%opt%"=="10" goto limpieza
if "%opt%"=="11" goto fixaudio
if "%opt%"=="0" exit
goto menu

:: ---------------------------------------------------------
:: SECCION DE APLICACIONES
:: ---------------------------------------------------------

:outlook
cls
echo %G%[>] Cerrando Outlook...%W%
taskkill /F /IM outlook.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Outlook...%W%
start outlook:
goto menu

:word
cls
echo %G%[>] Cerrando Word...%W%
taskkill /F /IM winword.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Word...%W%
start winword
goto menu

:excel
cls
echo %G%[>] Cerrando Excel...%W%
taskkill /F /IM excel.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Excel...%W%
start excel
goto menu

:teams
cls
echo %G%[>] Cerrando Microsoft Teams...%W%
taskkill /F /IM ms-teams.exe /T >nul 2>&1
taskkill /F /IM Teams.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Teams...%W%
start msteams:
goto menu

:chrome
cls
echo %G%[>] Cerrando Google Chrome...%W%
taskkill /F /IM chrome.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Chrome...%W%
start chrome
goto menu

:acrobat
cls
echo %G%[>] Cerrando Adobe Acrobat/Reader...%W%
taskkill /F /IM Acrobat.exe /T >nul 2>&1
taskkill /F /IM AcroRd32.exe /T >nul 2>&1
timeout /t 2 >nul
echo %G%[>] Iniciando Acrobat...%W%
start acrobat
if %errorLevel% neq 0 (
    echo %Y%[INFO] No se pudo iniciar automaticamente. Abrelo manualmente.%W%
    pause
)
goto menu

:: ---------------------------------------------------------
:: SECCION DE SISTEMA
:: ---------------------------------------------------------

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

echo %G%[>] Limpiando Temp de Usuario...%W%
del /q /s /f "%temp%\*" >nul 2>&1
for /d %%x in ("%temp%\*") do rd /s /q "%%x" >nul 2>&1

if %admin%==NO goto skip_sys_clean

echo %G%[>] Limpiando Temp de Windows...%W%
del /q /s /f "C:\Windows\Temp\*" >nul 2>&1
for /d %%x in ("C:\Windows\Temp\*") do rd /s /q "%%x" >nul 2>&1

echo %G%[>] Limpiando Prefetch...%W%
del /q /s /f "C:\Windows\Prefetch\*" >nul 2>&1
echo %G%[OK] Limpieza de sistema completada.%W%
goto fin_limpieza

:skip_sys_clean
echo %R%[!] Saltando carpetas de sistema (Requiere Admin).%W%

:fin_limpieza
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
