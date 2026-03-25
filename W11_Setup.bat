@echo off
:: =============================================================
:: Start.bat - Starter fuer W11_Setup.ps1
:: Doppelklick genuegt - UAC-Dialog erscheint automatisch
:: =============================================================

:: Pfad zum Skript ermitteln (gleicher Ordner wie diese BAT)
set "SKRIPT=%~dp0W11_Setup.ps1"

:: Pruefen ob Skript vorhanden
if not exist "%SKRIPT%" (
    echo.
    echo  FEHLER: W11_Setup.ps1 nicht gefunden!
    echo  Bitte Start.bat und W11_Setup.ps1 in denselben Ordner legen.
    echo.
    pause
    exit /b 1
)

:: PowerShell als Administrator starten via UAC
:: Pfad wird als Variable uebergeben - keine Anführungszeichen-Probleme
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$skript = '%SKRIPT:\=\\%'; Start-Process PowerShell -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File \"' + $skript + '\"') -Verb RunAs -Wait"
