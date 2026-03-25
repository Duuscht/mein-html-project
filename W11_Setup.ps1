# =============================================================
# W11_Setup.ps1
# Autor: Helmut | Erstellt mit Claude
# =============================================================
# Starten: Rechtsklick auf Start.bat -> Als Administrator ausfuehren
# =============================================================

$SkriptVersion = "V19"
$SkriptDatum   = "2026-03-17"

# Fehler still behandeln - keine aufblitzenden roten Meldungen
$ErrorActionPreference = "SilentlyContinue"


# =============================================================
# HILFSFUNKTIONEN
# =============================================================
function Show-Header {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Windows11_Setup  |  Version: $SkriptVersion  |  Stand: $SkriptDatum" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Success { param($msg) Write-Host "  OK: $msg" -ForegroundColor Green }
function Show-Warning { param($msg) Write-Host "  WARNUNG: $msg" -ForegroundColor Yellow }
function Show-Error   { param($msg) Write-Host "  FEHLER: $msg" -ForegroundColor Red }
function Show-Info    { param($msg) Write-Host "  INFO: $msg" -ForegroundColor White }

function Pause-Script {
    Write-Host ""
    Write-Host "  Druecke eine beliebige Taste um fortzufahren..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# HKU PSDrive registrieren um alle Benutzer-Hives lesen zu koennen
if (-not (Get-PSDrive -Name "HKU" -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name "HKU" -PSProvider Registry -Root "HKEY_USERS" | Out-Null
}

# Hilfsfunktion: Registry-Wert aus allen Benutzer-Hives lesen
function Get-HKUValue {
    param([string]$SubKey, [string]$ValueName)
    $results = @()
    Get-ChildItem "HKU:\" -ErrorAction SilentlyContinue | ForEach-Object {
        $sid = $_.PSChildName
        if ($sid -match "^S-1-5-21-") {
            $pfad = "HKU:\$sid\$SubKey"
            $val = Get-ItemProperty -Path $pfad -Name $ValueName -ErrorAction SilentlyContinue
            if ($null -ne $val) { $results += $val.$ValueName }
        }
    }
    return $results
}

# Hilfsfunktion: Registry-Wert in alle Benutzer-Hives schreiben
function Set-HKUValue {
    param([string]$SubKey, [string]$ValueName, $Value, [string]$Type = "DWord")
    $count = 0
    Get-ChildItem "HKU:\" -ErrorAction SilentlyContinue | ForEach-Object {
        $sid = $_.PSChildName
        if ($sid -match "^S-1-5-21-") {
            $pfad = "HKU:\$sid\$SubKey"
            if (Test-Path $pfad) {
                Set-ItemProperty -Path $pfad -Name $ValueName -Value $Value -Type $Type -Force -ErrorAction SilentlyContinue
                $count++
            }
        }
    }
    return $count
}

# Prueft ob App installiert ist - durchsucht HKLM + alle Benutzer-Hives (HKU) + AppX
function Get-InstallStatus {
    param([string]$DisplayName, [string]$WingetID)

    # 1. HKLM - systemweite Installationen
    $hklmPfade = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($pfad in $hklmPfade) {
        $hit = Get-ChildItem $pfad -ErrorAction SilentlyContinue |
            Get-ItemProperty -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like "*$DisplayName*" } |
            Select-Object -First 1
        if ($hit) { return @{ Installiert = $true; Version = $hit.DisplayVersion } }
    }

    # 2. HKU - alle Benutzerprofile durchsuchen
    $uninstallKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $hkuTreffer = $null
    Get-ChildItem "HKU:\" -ErrorAction SilentlyContinue | ForEach-Object {
        if ($hkuTreffer) { return }
        $sid = $_.PSChildName
        if ($sid -match "^S-1-5-21-") {
            $pfad = "HKU:\$sid\$uninstallKey"
            $hit = Get-ChildItem $pfad -ErrorAction SilentlyContinue |
                Get-ItemProperty -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -and $_.DisplayName -like "*$DisplayName*" } |
                Select-Object -First 1
            if ($hit) { $hkuTreffer = @{ Installiert = $true; Version = $hit.DisplayVersion } }
        }
    }
    if ($hkuTreffer) { return $hkuTreffer }

    # 3. AppX Store Apps (alle Benutzer)
    $appx = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*$DisplayName*" } |
        Select-Object -First 1
    if ($appx) { return @{ Installiert = $true; Version = $appx.Version } }

    return @{ Installiert = $false; Version = "" }
}

# =============================================================
# MODUL 1: SOFTWARE INSTALLATION
# =============================================================
function Install-Software {

    $software = @(
        @{ Nr = 1;  Name = "OneCalendar";  WingetID = "9NBLGGH5S5BN";                     DisplayName = "OneCalendar";      Quelle = "msstore" },
        @{ Nr = 2;  Name = "OneCommander"; WingetID = "MilosParipovic.OneCommander";       DisplayName = "OneCommander";     Quelle = "winget"  },
        @{ Nr = 3;  Name = "Vivaldi";      WingetID = "Vivaldi.Vivaldi";                   DisplayName = "Vivaldi";          Quelle = "winget"  },
        @{ Nr = 4;  Name = "Brave";        WingetID = "Brave.Brave";                       DisplayName = "Brave";            Quelle = "winget"  },
        @{ Nr = 5;  Name = "Firefox";      WingetID = "Mozilla.Firefox";                   DisplayName = "Mozilla Firefox";   Quelle = "winget" },
        @{ Nr = 6;  Name = "Bitwarden";    WingetID = "Bitwarden.Bitwarden";               DisplayName = "Bitwarden";        Quelle = "winget"  },
        @{ Nr = 7;  Name = "TeamViewer";   WingetID = "TeamViewer.TeamViewer";             DisplayName = "TeamViewer";       Quelle = "winget"  },
        @{ Nr = 8;  Name = "Thunderbird";  WingetID = "Mozilla.Thunderbird";               DisplayName = "Thunderbird";      Quelle = "winget"  },
        @{ Nr = 9;  Name = "pCloud_Drive"; WingetID = "pCloudAG.pCloudDrive";              DisplayName = "pCloud Drive";     Quelle = "direkt"  },
        @{ Nr = 10; Name = "NAPS2";        WingetID = "Cyanfish.NAPS2";                    DisplayName = "NAPS2";            Quelle = "winget"  },
        @{ Nr = 11; Name = "VLC";          WingetID = "VideoLAN.VLC";                      DisplayName = "VLC media player";  Quelle = "winget" },
        @{ Nr = 12; Name = "LibreOffice";  WingetID = "TheDocumentFoundation.LibreOffice"; DisplayName = "LibreOffice";      Quelle = "winget"  }
    )

    Show-Header
    Write-Host "  [MODUL 1] Software-Installation" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Pruefe installierte Programme..." -ForegroundColor Gray
    Write-Host ""
    Write-Host ("  " + "Nr".PadRight(6) + "Programm".PadRight(20) + "Status".PadRight(22) + "Version") -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    foreach ($app in $software) {
        $label  = $app.Name + $(if ($app.Quelle -eq "msstore") { " [Store]" } elseif ($app.Quelle -eq "direkt") { " [Web]" } else { "" })
        $nr     = "[$($app.Nr.ToString().PadLeft(2))]"
        $status = Get-InstallStatus -DisplayName $app.DisplayName -WingetID $app.WingetID

        Write-Host ("  " + $nr.PadRight(6) + $label.PadRight(20)) -NoNewline -ForegroundColor White
        if ($status.Installiert) {
            Write-Host ("Installiert".PadRight(22)) -NoNewline -ForegroundColor Green
            Write-Host $status.Version -ForegroundColor Gray
        } else {
            Write-Host "Nicht_installiert" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "  HINWEIS: OneCalendar  -> manuell im Microsoft Store installieren." -ForegroundColor Yellow
    Write-Host "           pCloud [Web] -> wird direkt von pcloud.com geladen." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [A]  Alle installieren" -ForegroundColor Cyan
    Write-Host "  [0]  Zurueck zum Hauptmenue" -ForegroundColor Gray
    Write-Host ""
    $auswahl = Read-Host "  Auswahl (Nummer oder mehrere z.B. 3,4,6 oder A)"
    if ($auswahl -eq "0") { return }

    $zuInstallieren = if ($auswahl -eq "A" -or $auswahl -eq "a") {
        $software
    } else {
        $nummern = $auswahl -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
        $software | Where-Object { $nummern -contains $_.Nr.ToString() }
    }

    Write-Host ""
    foreach ($app in $zuInstallieren) {
        Write-Host "  Installiere $($app.Name)..." -ForegroundColor White

        if ($app.Quelle -eq "msstore") {
            winget install --id $app.WingetID --source msstore --accept-package-agreements --accept-source-agreements

        } elseif ($app.Quelle -eq "direkt") {
            $ziel = "$env:TEMP\pCloud_Installer.exe"
            Write-Host "    Lade pCloud von pcloud.com..." -ForegroundColor Gray
            try {
                (New-Object System.Net.WebClient).DownloadFile("https://p-def8.pcloud.com/cBZgCl6bZQmoifV57ZZZCsSmkZ2ZZDQmZZZZDRLqiZL17e0ZZXjZZUkZZOQZeZXTZtRZdvZHnZCFZ7XZPSZFFZSuZ5oZ79ZcWZjMZ/pCloud_Windows_4.4.0.exe", $ziel)
                Start-Process -FilePath $ziel -ArgumentList "/S" -Wait
                Show-Success "pCloud installiert"
            } catch {
                Show-Warning "Download fehlgeschlagen - bitte manuell von pcloud.com herunterladen"
            }

        } else {
            winget install --id $app.WingetID --silent --accept-package-agreements --accept-source-agreements
            if ($LASTEXITCODE -eq 0) { Show-Success "$($app.Name) installiert" }
            else { Show-Warning "$($app.Name) - Fehler oder bereits installiert" }
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 2: ADMINISTRATOR-KONTO
# =============================================================
function Setup-Administrator {
    Show-Header
    Write-Host "  [MODUL 2] Administrator-Konto" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""

    $admin = Get-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
    if ($admin) {
        $statusStr   = if ($admin.Enabled) { "Aktiv" } else { "Deaktiviert" }
        $statusFarbe = if ($admin.Enabled) { "Green" } else { "Gray" }
        Write-Host "  Aktueller Status: " -NoNewline -ForegroundColor White
        Write-Host $statusStr -ForegroundColor $statusFarbe
        Write-Host ""
    }

    Write-Host "  [1]  Administrator aktivieren + Passwort setzen" -ForegroundColor White
    Write-Host "  [2]  Administrator deaktivieren" -ForegroundColor White
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl = Read-Host "  Auswahl"

    switch ($auswahl) {
        "1" {
            Write-Host ""
            $pw1 = Read-Host "  Neues Passwort fuer Administrator" -AsSecureString
            $pw2 = Read-Host "  Passwort bestaetigen" -AsSecureString

            $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                      [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw1))
            $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                      [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw2))

            if ($p1 -ne $p2) {
                Show-Error "Passwoerter stimmen nicht ueberein - Abbruch"
            } elseif ($p1.Length -lt 6) {
                Show-Error "Passwort zu kurz (mindestens 6 Zeichen)"
            } else {
                Set-LocalUser -Name "Administrator" -Password $pw1
                Enable-LocalUser -Name "Administrator"
                Show-Success "Administrator aktiviert und Passwort gesetzt"
            }
            $p1 = $null; $p2 = $null
        }
        "2" {
            $ok = Read-Host "  Administrator wirklich deaktivieren? (ja/nein)"
            if ($ok -eq "ja") {
                Disable-LocalUser -Name "Administrator"
                Show-Success "Administrator deaktiviert"
            }
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 3: BENUTZERRECHTE
# =============================================================
function Setup-Benutzerrechte {
    Show-Header
    Write-Host "  [MODUL 3] Benutzerrechte" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""

    $interaktivUser = $null
    try {
        $queryResult = query user 2>$null
        if ($queryResult) {
            $aktivZeile = $queryResult | Where-Object { $_ -match "Active|Aktiv" } | Select-Object -First 1
            if ($aktivZeile) {
                $interaktivUser = ($aktivZeile.Trim() -replace "^>", "" -split "\s+")[0]
            }
        }
    } catch {}

    if (-not $interaktivUser) {
        try {
            $explorer = Get-WmiObject Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($explorer) {
                $owner = $explorer.GetOwner()
                $interaktivUser = $owner.User
            }
        } catch {}
    }

    if (-not $interaktivUser) { $interaktivUser = $env:USERNAME }

    $adminGruppe = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Name

    Write-Host ("  " + "Name".PadRight(25) + "Aktiv".PadRight(8) + "Rolle".PadRight(20) + "Letzter_Login") -ForegroundColor Cyan
    Write-Host ("  " + "-" * 70) -ForegroundColor Gray

    Get-LocalUser | ForEach-Object {
        $istAdmin   = ($adminGruppe -like "*\$($_.Name)") -or ($adminGruppe -contains $_.Name)
        $rolle      = if ($istAdmin) { "Administrator" } else { "Standardbenutzer" }
        $farbe      = if ($istAdmin) { "Yellow" } else { "White" }
        $aktiv      = if ($_.Enabled) { "Ja" } else { "Nein" }
        $login      = if ($_.LastLogon) { $_.LastLogon.ToString("dd.MM.yyyy_HH:mm") } else { "-" }
        $istAktuell = $_.Name -eq $interaktivUser
        $markierung = if ($istAktuell) { "  <-- du" } else { "" }

        Write-Host ("  " + $_.Name.PadRight(25) + $aktiv.PadRight(8)) -NoNewline -ForegroundColor White
        Write-Host ($rolle.PadRight(20)) -NoNewline -ForegroundColor $farbe
        if ($istAktuell) {
            Write-Host ($login) -NoNewline -ForegroundColor Gray
            Write-Host $markierung -ForegroundColor Cyan
        } else {
            Write-Host $login -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "  Eingeloggter Benutzer: $interaktivUser" -ForegroundColor Cyan
    Write-Host ""

    $zielUser = $interaktivUser
    Write-Host "  [1]  $zielUser -> Standardbenutzer" -ForegroundColor White
    Write-Host "  [2]  $zielUser -> Administrator" -ForegroundColor White
    Write-Host "  [3]  Anderen Benutzer auswaehlen" -ForegroundColor White
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl = Read-Host "  Auswahl"

    switch ($auswahl) {
        "1" {
            $adminCheck = Get-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
            if (-not $adminCheck.Enabled) {
                Show-Warning "Achtung: Administrator-Konto ist deaktiviert!"
                Show-Warning "Erst Modul 2 ausfuehren, sonst kein Admin-Zugang mehr!"
                $ok = Read-Host "  Trotzdem fortfahren? (ja/nein)"
            } else { $ok = "ja" }
            if ($ok -eq "ja") {
                Remove-LocalGroupMember -Group "Administrators" -Member $zielUser -ErrorAction SilentlyContinue
                Show-Success "$zielUser ist jetzt Standardbenutzer"
            }
        }
        "2" {
            Add-LocalGroupMember -Group "Administrators" -Member $zielUser -ErrorAction SilentlyContinue
            Show-Success "$zielUser ist jetzt Administrator"
        }
        "3" {
            Write-Host ""
            Write-Host "  Verfuegbare Benutzer:" -ForegroundColor Cyan
            Get-LocalUser | Where-Object { $_.Enabled } | ForEach-Object {
                Write-Host "    $($_.Name)" -ForegroundColor White
            }
            Write-Host ""
            $zielUser = Read-Host "  Benutzername eingeben"
            if (Get-LocalUser -Name $zielUser -ErrorAction SilentlyContinue) {
                $aktion = Read-Host "  [1] Standardbenutzer  [2] Administrator"
                if ($aktion -eq "1") {
                    Remove-LocalGroupMember -Group "Administrators" -Member $zielUser -ErrorAction SilentlyContinue
                    Show-Success "$zielUser ist jetzt Standardbenutzer"
                } elseif ($aktion -eq "2") {
                    Add-LocalGroupMember -Group "Administrators" -Member $zielUser -ErrorAction SilentlyContinue
                    Show-Success "$zielUser ist jetzt Administrator"
                }
            } else {
                Show-Error "Benutzer '$zielUser' nicht gefunden"
            }
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 4: BOOT-NACHRICHTEN
# =============================================================
function Setup-BootNachrichten {
    Show-Header
    Write-Host "  [MODUL 4] Boot-Nachrichten" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [1]  Verbose Boot aktivieren (Details beim Start/Shutdown)" -ForegroundColor White
    Write-Host "  [2]  Verbose Boot deaktivieren" -ForegroundColor White
    Write-Host "  [3]  Schnellstart deaktivieren (Dual-Boot empfohlen)" -ForegroundColor White
    Write-Host "  [4]  Schnellstart aktivieren" -ForegroundColor White
    Write-Host "  [A]  Beides: Verbose an + Schnellstart aus" -ForegroundColor Cyan
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl   = Read-Host "  Auswahl"
    $regPath   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    $powerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"

    switch ($auswahl) {
        "1" { Set-ItemProperty -Path $regPath   -Name "VerboseStatus"    -Value 1 -Type DWord; Show-Success "Verbose_Boot aktiviert" }
        "2" { Set-ItemProperty -Path $regPath   -Name "VerboseStatus"    -Value 0 -Type DWord; Show-Success "Verbose_Boot deaktiviert" }
        "3" { Set-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -Value 0 -Type DWord; Show-Success "Schnellstart deaktiviert" }
        "4" { Set-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -Value 1 -Type DWord; Show-Success "Schnellstart aktiviert" }
        { $_ -eq "A" -or $_ -eq "a" } {
            Set-ItemProperty -Path $regPath   -Name "VerboseStatus"    -Value 1 -Type DWord
            Set-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -Value 0 -Type DWord
            Show-Success "Verbose_Boot an + Schnellstart aus"
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 5: SCROLLBALKEN
# =============================================================
function Setup-Scrollbalken {
    Show-Header
    Write-Host "  [MODUL 5] Scrollbalken" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [1]  Immer sichtbar + 3x breiter (empfohlen)" -ForegroundColor White
    Write-Host "  [2]  Nur immer sichtbar" -ForegroundColor White
    Write-Host "  [3]  Nur 3x breiter" -ForegroundColor White
    Write-Host "  [4]  Windows-Standard wiederherstellen" -ForegroundColor White
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl    = Read-Host "  Auswahl"
    $uiPath     = "HKCU:\Control Panel\Accessibility"
    $metricPath = "HKCU:\Control Panel\Desktop\WindowMetrics"

    switch ($auswahl) {
        "1" {
            Set-ItemProperty -Path $uiPath     -Name "DynamicScrollbars" -Value 0     -Type DWord
            Set-ItemProperty -Path $metricPath -Name "ScrollWidth"        -Value "-51" -Type String
            Set-ItemProperty -Path $metricPath -Name "ScrollHeight"       -Value "-51" -Type String
            Show-Success "Immer sichtbar + 3x breiter"
        }
        "2" { Set-ItemProperty -Path $uiPath -Name "DynamicScrollbars" -Value 0 -Type DWord; Show-Success "Immer sichtbar" }
        "3" {
            Set-ItemProperty -Path $metricPath -Name "ScrollWidth"  -Value "-51" -Type String
            Set-ItemProperty -Path $metricPath -Name "ScrollHeight" -Value "-51" -Type String
            Show-Success "3x breiter"
        }
        "4" {
            Set-ItemProperty -Path $uiPath     -Name "DynamicScrollbars" -Value 1     -Type DWord
            Set-ItemProperty -Path $metricPath -Name "ScrollWidth"        -Value "-17" -Type String
            Set-ItemProperty -Path $metricPath -Name "ScrollHeight"       -Value "-17" -Type String
            Show-Success "Windows-Standard wiederhergestellt"
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 6: MICROSOFT BLOATWARE ENTFERNEN
# =============================================================
function Remove-Bloatware {
    $bloatware = @(
        @{ Nr = 1;  Name = "OneDrive";              Typ = "winget";   AppxID = "Microsoft.OneDrive"                         },
        @{ Nr = 2;  Name = "Outlook_neu";           Typ = "appx";     AppxID = "Microsoft.OutlookForWindows"                },
        @{ Nr = 3;  Name = "Copilot";               Typ = "appx";     AppxID = "Microsoft.Copilot"                          },
        @{ Nr = 4;  Name = "Cortana";               Typ = "appx";     AppxID = "Microsoft.549981C3F5F10"                    },
        @{ Nr = 5;  Name = "Teams_Consumer";        Typ = "appx";     AppxID = "MSTeams"                                    },
        @{ Nr = 6;  Name = "Bing_App";              Typ = "appx";     AppxID = "Microsoft.BingSearch"                       },
        @{ Nr = 7;  Name = "Bing_Suche_Taskleiste"; Typ = "reg";      AppxID = ""                                           },
        @{ Nr = 8;  Name = "Clipchamp";             Typ = "appx";     AppxID = "Clipchamp.Clipchamp"                        },
        @{ Nr = 9;  Name = "Family";                Typ = "appx";     AppxID = "MicrosoftCorporationII.MicrosoftFamily"     },
        @{ Nr = 10; Name = "Feedback_Hub";          Typ = "appx";     AppxID = "Microsoft.WindowsFeedbackHub"               },
        @{ Nr = 11; Name = "Media_Player";          Typ = "appx";     AppxID = "Microsoft.ZuneMusic"                        },
        @{ Nr = 12; Name = "To_Do";                 Typ = "appx";     AppxID = "Microsoft.Todos"                            },
        @{ Nr = 13; Name = "Xbox_Apps";             Typ = "appx";     AppxID = "Microsoft.GamingApp"                        },
        @{ Nr = 14; Name = "Solitaire_Collection";  Typ = "appx";     AppxID = "Microsoft.MicrosoftSolitaireCollection"     },
        @{ Nr = 15; Name = "Get_Help";              Typ = "appx";     AppxID = "Microsoft.GetHelp"                          },
        @{ Nr = 16; Name = "Tipps_App";             Typ = "appx";     AppxID = "Microsoft.Getstarted"                       }
    )

    Show-Header
    Write-Host "  [MODUL 6] Microsoft_Bloatware entfernen" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Pruefe Komponenten..." -ForegroundColor Gray
    Write-Host ""
    Write-Host ("  " + "Nr".PadRight(6) + "Komponente".PadRight(30) + "Status") -ForegroundColor Cyan
    Write-Host ("  " + "-" * 55) -ForegroundColor Gray

    foreach ($item in $bloatware) {
        $nr = "[$($item.Nr.ToString().PadLeft(2))]"
        $vorhanden = switch ($item.Typ) {
            "appx"     {
                $pkg = Get-AppxPackage -Name "*$($item.AppxID)*" -AllUsers -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $pkg) {
                    $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                        Where-Object { $_.PackageName -like "*$($item.AppxID)*" } |
                        Select-Object -First 1
                    $null -ne $prov
                } else { $true }
            }
            "winget"   {
                $gesperrt = (Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -ErrorAction SilentlyContinue).DisableFileSyncNGSC -eq 1
                $laeuft = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
                $datei  = Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
                ($laeuft -or $datei) -and (-not $gesperrt)
            }
            "reg"      {
                $werte = Get-HKUValue -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -ValueName "BingSearchEnabled"
                ($werte.Count -eq 0) -or ($werte | Where-Object { $_ -ne 0 })
            }
        }

        Write-Host ("  " + $nr.PadRight(6) + $item.Name.PadRight(30)) -NoNewline -ForegroundColor White
        if ($vorhanden) {
            Write-Host "Vorhanden" -ForegroundColor Red
        } else {
            Write-Host "Bereits_entfernt" -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "  [A]  Alle entfernen" -ForegroundColor Cyan
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl = Read-Host "  Auswahl (Nummer oder mehrere z.B. 1,3,6 oder A)"
    if ($auswahl -eq "0") { return }

    $zuEntfernen = if ($auswahl -eq "A" -or $auswahl -eq "a") {
        $bloatware
    } else {
        $nummern = $auswahl -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
        $bloatware | Where-Object { $nummern -contains $_.Nr.ToString() }
    }

    Write-Host ""
    foreach ($item in $zuEntfernen) {
        Write-Host "  Entferne $($item.Name)..." -ForegroundColor White

        switch ($item.Typ) {
            "appx" {
                $erfolg = $false
                $pkgs = Get-AppxPackage -Name "*$($item.AppxID)*" -AllUsers -ErrorAction SilentlyContinue
                if ($pkgs) {
                    $pkgs | ForEach-Object {
                        Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                    }
                    $erfolg = $true
                }
                $provPkgs = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                    Where-Object { $_.PackageName -like "*$($item.AppxID)*" }
                if ($provPkgs) {
                    $provPkgs | ForEach-Object {
                        Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
                    }
                    $erfolg = $true
                }
                $wgCheck = (winget list --id $item.AppxID --accept-source-agreements 2>$null) | Out-String
                if ($wgCheck -match [regex]::Escape($item.AppxID)) {
                    winget uninstall --id $item.AppxID --silent --accept-source-agreements 2>$null | Out-Null
                    $erfolg = $true
                }
                if ($erfolg) { Show-Success "$($item.Name) entfernt" }
                else { Show-Warning "$($item.Name) - nicht gefunden (moeglicherweise bereits entfernt)" }
            }
            "winget" {
                Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 1
                winget uninstall --id $item.AppxID --silent --accept-source-agreements 2>$null

                $regPfad = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
                if (-not (Test-Path $regPfad)) { New-Item -Path $regPfad -Force | Out-Null }
                Set-ItemProperty -Path $regPfad -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force

                Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -ErrorAction SilentlyContinue
                Get-ChildItem "HKU:\" -ErrorAction SilentlyContinue | ForEach-Object {
                    $sid = $_.PSChildName
                    if ($sid -match "^S-1-5-21-") {
                        Remove-ItemProperty -Path "HKU:\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -ErrorAction SilentlyContinue
                    }
                }

                Show-Success "OneDrive deinstalliert + Neuinstallation dauerhaft gesperrt"
            }
            "reg" {
                $n = Set-HKUValue -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -ValueName "BingSearchEnabled" -Value 0
                Set-HKUValue -SubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -ValueName "CortanaConsent" -Value 0 | Out-Null
                if ($n -gt 0) { Show-Success "Bing_Suche in $n Benutzerprofil(en) deaktiviert" }
                else { Show-Warning "Kein passendes Benutzerprofil gefunden" }
            }
        }
    }
    Pause-Script
}

# =============================================================
# MODUL 7: SICHERHEITS-CHECK
# =============================================================
function Check-Sicherheit {
    Show-Header
    Write-Host "  [MODUL 7] Sicherheits-Check" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""

    function Show-Check {
        param([string]$Label, [bool]$Ok, [string]$Info = "")
        Write-Host ("  " + $Label.PadRight(35)) -NoNewline -ForegroundColor White
        if ($Ok) { Write-Host "OK     " -NoNewline -ForegroundColor Green }
        else      { Write-Host "WARNUNG" -NoNewline -ForegroundColor Red }
        if ($Info) { Write-Host "  $Info" -ForegroundColor Gray }
        else { Write-Host "" }
    }

    $srPunkte = Get-ComputerRestorePoint -ErrorAction SilentlyContinue
    $srAnzahl = if ($srPunkte) { @($srPunkte).Count } else { 0 }
    $srOk     = $srAnzahl -gt 0
    $srInfo   = if ($srOk) {
        $neuester = ($srPunkte | Sort-Object CreationTime -Descending | Select-Object -First 1)
        $datum = [System.Management.ManagementDateTimeConverter]::ToDateTime($neuester.CreationTime).ToString("dd.MM.yyyy HH:mm")
        "$srAnzahl Punkt(e), letzter: $datum"
    } else { "Keine_Wiederherstellungspunkte_vorhanden!" }
    Show-Check "System_Restore" $srOk $srInfo

    $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defender) {
        Show-Check "Defender_Echtzeit" ($defender.RealTimeProtectionEnabled) $(if ($defender.RealTimeProtectionEnabled) { "Aktiv" } else { "Deaktiviert!" })
        $defAge = (Get-Date) - $defender.AntivirusSignatureLastUpdated
        Show-Check "Defender_Signaturen" ($defAge.TotalDays -lt 3) "Alter: $([int]$defAge.TotalDays) Tage"
    } else {
        Show-Check "Defender" $false "Nicht_lesbar"
    }

    $fw = Get-NetFirewallProfile -ErrorAction SilentlyContinue
    foreach ($profil in $fw) {
        Show-Check "Firewall_$($profil.Name)" ($profil.Enabled) $(if ($profil.Enabled) { "Aktiv" } else { "Deaktiviert!" })
    }

    $uac = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue).ConsentPromptBehaviorAdmin
    Show-Check "UAC_Level" ($uac -ge 2) "Level: $uac (empfohlen: 2+)"

    $wu = (Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue)
    $autoUpdate = if ($wu) { $wu.NoAutoUpdate -eq 0 } else { $true }
    Show-Check "Automatische_Updates" $autoUpdate $(if ($autoUpdate) { "Aktiv" } else { "Deaktiviert!" })

    $rdp = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -ErrorAction SilentlyContinue).fDenyTSConnections
    Show-Check "Remote_Desktop" ($rdp -eq 1) $(if ($rdp -eq 1) { "Deaktiviert_(sicher)" } else { "Aktiv_(Risiko!)" })

    $gast = Get-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
    Show-Check "Gastkonto" (-not $gast.Enabled) $(if ($gast.Enabled) { "Aktiv_(Risiko!)" } else { "Deaktiviert_(sicher)" })

    Write-Host ""
    Write-Host "  [1]  System_Restore - neuen Punkt erstellen" -ForegroundColor White
    Write-Host "  [2]  System_Restore deaktivieren + alle Punkte loeschen" -ForegroundColor White
    Write-Host "  [3]  Remote_Desktop deaktivieren" -ForegroundColor White
    Write-Host "  [4]  Gastkonto deaktivieren" -ForegroundColor White
    Write-Host "  [0]  Zurueck" -ForegroundColor Gray
    Write-Host ""
    $auswahl = Read-Host "  Auswahl"

    switch ($auswahl) {
        "1" {
            $beschr = Read-Host "  Beschreibung fuer den Wiederherstellungspunkt"
            if (-not $beschr) { $beschr = "Manuell_erstellt_$(Get-Date -Format 'yyyy-MM-dd')" }
            Checkpoint-Computer -Description $beschr -RestorePointType "MODIFY_SETTINGS" -ErrorAction SilentlyContinue
            Show-Success "Wiederherstellungspunkt erstellt: $beschr"
        }
        "2" {
            $ok = Read-Host "  Alle Punkte loeschen und System_Restore deaktivieren? (ja/nein)"
            if ($ok -eq "ja") {
                Disable-ComputerRestore -Drive "C:" -ErrorAction SilentlyContinue
                vssadmin delete shadows /all /quiet 2>$null
                Show-Success "System_Restore deaktiviert und alle Punkte geloescht"
            }
        }
        "3" {
            Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 1 -Type DWord
            Show-Success "Remote_Desktop deaktiviert"
        }
        "4" {
            Disable-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
            Show-Success "Gastkonto deaktiviert"
        }
    }
    if ($auswahl -ne "0") { Pause-Script }
}

# =============================================================
# MODUL 8: HARDWARE & NETZWERK INFO
# =============================================================
function Show-HardwareInfo {
    Show-Header
    Write-Host "  [MODUL 8] Hardware & Netzwerk-Info" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""

    # --- FESTPLATTEN ---
    Write-Host "  FESTPLATTEN" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
        Where-Object { $_.Used -ne $null -and ($_.Used + $_.Free) -gt 0 } | ForEach-Object {
        $gesamt = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
        $frei   = [math]::Round($_.Free / 1GB, 1)
        $belegt = [math]::Round($_.Used / 1GB, 1)
        $proz   = [math]::Round($_.Used / ($_.Used + $_.Free) * 100)
        $farbe  = if ($proz -gt 90) { "Red" } elseif ($proz -gt 75) { "Yellow" } else { "Green" }
        Write-Host ("  " + "[$($_.Name):]".PadRight(6) + "Gesamt:${gesamt}GB".PadRight(16) + "Belegt:${belegt}GB".PadRight(16) + "Frei:${frei}GB".PadRight(14)) -NoNewline -ForegroundColor White
        Write-Host "${proz}%" -ForegroundColor $farbe
    }

    # --- SMART-Status ---
    Write-Host ""
    Write-Host "  SMART-STATUS" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    Get-PhysicalDisk -ErrorAction SilentlyContinue | ForEach-Object {
        $farbe = switch ($_.HealthStatus) {
            "Healthy"   { "Green"  }
            "Warning"   { "Yellow" }
            "Unhealthy" { "Red"    }
            default     { "Gray"   }
        }
        $groesse = [math]::Round($_.Size / 1GB)
        Write-Host ("  " + $_.FriendlyName.PadRight(38) + "${groesse}GB".PadRight(8)) -NoNewline -ForegroundColor White
        Write-Host $_.HealthStatus -ForegroundColor $farbe
    }

    # --- RAM ---
    Write-Host ""
    Write-Host "  ARBEITSSPEICHER" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
    if ($os -and $cs) {
        $gesamtBytes = $cs.TotalPhysicalMemory
        $freiBytes   = $os.FreePhysicalMemory * 1KB
        $belegtBytes = $gesamtBytes - $freiBytes
        $gesamt = [math]::Round($gesamtBytes / 1GB, 1)
        $frei   = [math]::Round($freiBytes   / 1GB, 1)
        $belegt = [math]::Round($belegtBytes / 1GB, 1)
        $proz   = [math]::Round($belegtBytes / $gesamtBytes * 100)
        $farbe  = if ($proz -gt 85) { "Red" } elseif ($proz -gt 70) { "Yellow" } else { "Green" }
        Write-Host ("  " + "Gesamt:${gesamt}GB".PadRight(18) + "Belegt:${belegt}GB".PadRight(18) + "Frei:${frei}GB".PadRight(14)) -NoNewline -ForegroundColor White
        Write-Host "${proz}%" -ForegroundColor $farbe
    }

    # --- NETZWERKKARTEN ---
    Write-Host ""
    Write-Host "  NETZWERKKARTEN" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    Write-Host ("  " + "Name".PadRight(20) + "Typ".PadRight(10) + "Status".PadRight(18) + "MAC / Speed") -ForegroundColor Cyan

    function Get-AdapterTyp { param($desc)
        if     ($desc -match "Wi-Fi|Wireless|WLAN|802\.11")                              { "WLAN" }
        elseif ($desc -match "NordLynx|WireGuard|OpenVPN|TAP|NordVPN|Tunnel")            { "VPN"  }
        elseif ($desc -match "Bluetooth")                                                 { "BT"   }
        elseif ($desc -match "Ethernet|Realtek PCIe|Intel.*Ethernet|Family Controller")  { "LAN"  }
        else                                                                              { $null  }
    }

    $alleAdapter = Get-NetAdapter -ErrorAction SilentlyContinue
    foreach ($a in $alleAdapter) {
        $typ = Get-AdapterTyp $a.InterfaceDescription
        if ($null -eq $typ) { continue }

        $statusStr = switch ($a.Status) {
            "Up"           { "Verbunden" }
            "Disconnected" { "Nicht_verbunden" }
            "Disabled"     { "Deaktiviert" }
            default        { $a.Status }
        }

        $speedRaw = $a.LinkSpeed
        $speedStr = if ($speedRaw -match "^\d+$") {
            "$([math]::Round([long]$speedRaw / 1MB))_Mbps"
        } else { "$speedRaw" -replace "\s+", "_" }

        $macStr = if ($typ -eq "VPN" -or $a.MacAddress -eq "00-00-00-00-00-00" -or -not $a.MacAddress) {
            "---"
        } else { $a.MacAddress }

        $nameStr = $a.Name
        if ($nameStr.Length -gt 18) { $nameStr = $nameStr.Substring(0,16) + ".." }

        $istVpnHilfsadapter = ($typ -eq "VPN" -and $a.Status -ne "Up" -and
            $a.InterfaceDescription -match "TAP|OpenVPN|Local Area Connection")

        $farbe = if ($istVpnHilfsadapter) { "DarkGray" } else {
            switch ($typ) {
                "VPN"  { "Yellow" }
                "WLAN" { "Cyan"   }
                "LAN"  { "White"  }
                "BT"   { "Gray"   }
            }
        }
        $statusFarbe = if ($a.Status -eq "Up") { "Green" } else { "DarkGray" }

        Write-Host ("  " + $nameStr.PadRight(20) + $typ.PadRight(7)) -NoNewline -ForegroundColor $farbe
        Write-Host ($statusStr.PadRight(18)) -NoNewline -ForegroundColor $statusFarbe
        if ($a.Status -eq "Up") {
            Write-Host "$macStr  $speedStr" -ForegroundColor Gray
        } else {
            Write-Host $macStr -ForegroundColor DarkGray
        }
    }

    # --- AKTIVE VERBINDUNGEN ---
    Write-Host ""
    Write-Host "  AKTIVE_VERBINDUNGEN" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    $verbindungen = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
        Where-Object { $_.IPv4Address -and $_.IPv4Address.IPAddress -notmatch "^169\.254\." }

    if ($verbindungen) {
        foreach ($v in $verbindungen) {
            $adapter = Get-NetAdapter -InterfaceIndex $v.InterfaceIndex -ErrorAction SilentlyContinue
            $typ = if ($adapter) { Get-AdapterTyp $adapter.InterfaceDescription } else { "?" }
            $typLabel = "[$typ]"
            $farbe = switch ($typ) {
                "VPN"  { "Yellow" }
                "WLAN" { "Cyan"   }
                "LAN"  { "White"  }
                default{ "Gray"   }
            }
            Write-Host ("  " + $typLabel.PadRight(8) + $v.InterfaceAlias.PadRight(22)) -NoNewline -ForegroundColor $farbe
            Write-Host "IP: $($v.IPv4Address.IPAddress)" -ForegroundColor Green
            if ($v.IPv4DefaultGateway) {
                Write-Host ("  " + "".PadRight(30) + "Gateway: $($v.IPv4DefaultGateway.NextHop)") -ForegroundColor Gray
            }
            $dns = ($v.DNSServer | Where-Object { $_.AddressFamily -eq 2 } |
                Select-Object -ExpandProperty ServerAddresses -ErrorAction SilentlyContinue) -join ", "
            if ($dns) {
                Write-Host ("  " + "".PadRight(30) + "DNS:     $dns") -ForegroundColor Gray
            }
            Write-Host ""
        }
    } else {
        Write-Host "  Keine_aktiven_Verbindungen_gefunden" -ForegroundColor Gray
        Write-Host ""
    }

    # --- VPN STATUS ---
    Write-Host "  VPN_STATUS" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    $vpnAdapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
        $_.InterfaceDescription -match "NordLynx|WireGuard" -or
        ($_.InterfaceDescription -match "OpenVPN|TAP" -and $_.Status -eq "Up")
    }

    if ($vpnAdapters) {
        foreach ($vpn in $vpnAdapters) {
            $vpnIp = (Get-NetIPAddress -InterfaceIndex $vpn.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress -notmatch "^169\.254\." } |
                Select-Object -First 1).IPAddress

            $beschr = $vpn.InterfaceDescription
            if ($beschr.Length -gt 35) { $beschr = $beschr.Substring(0, 33) + ".." }

            Write-Host ("  " + $beschr.PadRight(37)) -NoNewline -ForegroundColor White
            if ($vpn.Status -eq "Up" -and $vpnIp) {
                Write-Host "VERBUNDEN  " -NoNewline -ForegroundColor Green
                Write-Host "Tunnel-IP: $vpnIp" -ForegroundColor Gray

                Write-Host "  Standort wird ermittelt..." -NoNewline -ForegroundColor DarkGray
                try {
                    $geoRaw = Invoke-WebRequest -Uri "https://ipinfo.io/json" -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
                    $geo = $geoRaw.Content | ConvertFrom-Json

                    Write-Host ("`r" + " " * 50 + "`r") -NoNewline

                    $land   = if ($geo.country) { $geo.country } else { "?" }
                    $stadt  = if ($geo.city)    { $geo.city }    else { "?" }
                    $region = if ($geo.region)  { $geo.region }  else { "" }
                    $org    = if ($geo.org)     { $geo.org }     else { "" }
                    $oeffentlicheIP = if ($geo.ip) { $geo.ip } else { "?" }

                    Write-Host ("  " + "".PadRight(37) + "Exit-IP:  ") -NoNewline -ForegroundColor Gray
                    Write-Host $oeffentlicheIP -ForegroundColor Cyan
                    Write-Host ("  " + "".PadRight(37) + "Land:     ") -NoNewline -ForegroundColor Gray
                    Write-Host "$stadt, $region ($land)" -ForegroundColor Cyan
                    if ($org) {
                        Write-Host ("  " + "".PadRight(37) + "Provider: ") -NoNewline -ForegroundColor Gray
                        Write-Host $org -ForegroundColor Gray
                    }
                } catch {
                    Write-Host ("`r" + " " * 50 + "`r") -NoNewline
                    Write-Host ("  " + "".PadRight(37) + "Standort: Nicht_ermittelbar (kein Internet?)") -ForegroundColor DarkGray
                }
            } elseif ($vpn.Status -eq "Up") {
                Write-Host "AKTIV_(kein_Tunnel)" -ForegroundColor Yellow
            } else {
                Write-Host "NICHT_VERBUNDEN" -ForegroundColor Red
            }
        }
    } else {
        $nvCheck = Get-InstallStatus -DisplayName "NordVPN" -WingetID "NordVPN.NordVPN"
        if ($nvCheck.Installiert) {
            Write-Host ("  " + "NordVPN".PadRight(37)) -NoNewline -ForegroundColor White
            Write-Host "Installiert_-_kein_aktiver_Tunnel" -ForegroundColor Yellow
            Write-Host "  Tipp: In NordVPN einloggen und Server waehlen" -ForegroundColor Gray
        } else {
            Write-Host "  Kein_VPN_installiert_oder_erkannt" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Pause-Script
}

# =============================================================
# MODUL 9: SYSTEM-DIAGNOSE (BitLocker / UEFI / Secure Boot / Zertifikate)
# =============================================================
function Show-SystemDiagnose {
    Show-Header
    Write-Host "  [MODUL 9] System-Diagnose" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------" -ForegroundColor Gray
    Write-Host ""

    # Hilfsfunktion fuer einheitliche Ausgabezeilen
    function Show-DiagLine {
        param([string]$Label, [string]$Wert, [string]$Farbe = "White", [string]$Info = "")
        Write-Host ("  " + $Label.PadRight(30)) -NoNewline -ForegroundColor Gray
        Write-Host ($Wert.PadRight(20)) -NoNewline -ForegroundColor $farbe
        if ($Info) { Write-Host $Info -ForegroundColor DarkGray }
        else { Write-Host "" }
    }

    # ----------------------------------------------------------
    # ABSCHNITT 1: UEFI / Secure Boot
    # ----------------------------------------------------------
    Write-Host "  BOOT-UMGEBUNG" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    # UEFI oder Legacy
    $firmwareTyp = $env:firmware_type
    if (-not $firmwareTyp) {
        # Fallback ueber WMI
        $firmwareTyp = try {
            $reg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -ErrorAction SilentlyContinue
            if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State") { "UEFI" } else { "Legacy" }
        } catch { "Unbekannt" }
    }
    $uefiOk    = $firmwareTyp -eq "UEFI"
    $uefiFarbe = if ($uefiOk) { "Green" } else { "Yellow" }
    Write-Host ("  " + "Boot-Modus".PadRight(30)) -NoNewline -ForegroundColor Gray
    Write-Host $firmwareTyp -ForegroundColor $uefiFarbe

    # Secure Boot
    $sbStatus = "Nicht_verfuegbar"
    $sbFarbe  = "Gray"
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($sb -eq $true)  { $sbStatus = "Aktiv";        $sbFarbe = "Green"  }
        if ($sb -eq $false) { $sbStatus = "Deaktiviert";  $sbFarbe = "Yellow" }
    } catch {
        if ($_.Exception.Message -match "not supported|nicht unterstuetzt|Cmdlet") {
            $sbStatus = "Nicht_unterstuetzt_(Legacy-BIOS)"
            $sbFarbe  = "Gray"
        } else {
            $sbStatus = "Fehler_beim_Lesen"
            $sbFarbe  = "Red"
        }
    }
    Write-Host ("  " + "Secure Boot".PadRight(30)) -NoNewline -ForegroundColor Gray
    Write-Host $sbStatus -ForegroundColor $sbFarbe

    # TPM
    Write-Host ""
    Write-Host "  TPM" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    $tpm = Get-Tpm -ErrorAction SilentlyContinue
    if ($tpm) {
        $tpmVorhanden = $tpm.TpmPresent
        $tpmAktiv     = $tpm.TpmEnabled
        $tpmFarbe     = if ($tpmVorhanden -and $tpmAktiv) { "Green" } elseif ($tpmVorhanden) { "Yellow" } else { "Red" }
        $tpmStatus    = if ($tpmVorhanden -and $tpmAktiv) { "Vorhanden_und_aktiv" } elseif ($tpmVorhanden) { "Vorhanden_aber_deaktiviert" } else { "Nicht_vorhanden" }
        Write-Host ("  " + "TPM-Status".PadRight(30)) -NoNewline -ForegroundColor Gray
        Write-Host $tpmStatus -ForegroundColor $tpmFarbe

        # TPM-Version aus Registry
        $tpmVer = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\TPM\WMI" -ErrorAction SilentlyContinue).ManufacturerVersion
        if (-not $tpmVer) {
            $tpmVer = try {
                (Get-CimInstance -Namespace "root\cimv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction Stop).SpecVersion
            } catch { $null }
        }
        if ($tpmVer) {
            Write-Host ("  " + "TPM-Version".PadRight(30)) -NoNewline -ForegroundColor Gray
            Write-Host $tpmVer -ForegroundColor Gray
        }
    } else {
        Write-Host ("  " + "TPM-Status".PadRight(30)) -NoNewline -ForegroundColor Gray
        Write-Host "Nicht_lesbar_(kein_Zugriff)" -ForegroundColor Gray
    }

    # ----------------------------------------------------------
    # ABSCHNITT 2: BitLocker
    # ----------------------------------------------------------
    Write-Host ""
    Write-Host "  BITLOCKER" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray
    Write-Host ("  " + "Laufwerk".PadRight(12) + "Status".PadRight(22) + "Schutz".PadRight(18) + "Verschluesselung") -ForegroundColor Cyan

    $blVolumes = Get-BitLockerVolume -ErrorAction SilentlyContinue
    if ($blVolumes) {
        foreach ($vol in $blVolumes) {
            $mp     = $vol.MountPoint
            $vst    = $vol.VolumeStatus
            $pst    = $vol.ProtectionStatus
            $proz   = "$($vol.EncryptionPercentage)%"

            $schutzStr  = if ($pst -eq "On")  { "Aktiv" } elseif ($pst -eq "Off") { "Aus" } else { "$pst" }
            $schutzFarbe = switch ($pst) {
                "On"      { "Green"  }
                "Off"     { "Yellow" }
                default   { "Gray"   }
            }

            $vstFarbe = switch -Wildcard ($vst) {
                "FullyEncrypted"     { "Green"  }
                "FullyDecrypted"     { "Yellow" }
                "EncryptionInProgress" { "Cyan" }
                default              { "Gray"   }
            }

            Write-Host ("  " + $mp.PadRight(12)) -NoNewline -ForegroundColor White
            Write-Host ($vst.PadRight(22)) -NoNewline -ForegroundColor $vstFarbe
            Write-Host ($schutzStr.PadRight(18)) -NoNewline -ForegroundColor $schutzFarbe
            Write-Host $proz -ForegroundColor Gray
        }
    } else {
        Write-Host "  BitLocker-Daten nicht lesbar (Modul nicht verfuegbar oder kein Laufwerk)" -ForegroundColor Gray
    }

    # ----------------------------------------------------------
    # ABSCHNITT 3: Zertifikate
    # ----------------------------------------------------------
    Write-Host ""
    Write-Host "  ZERTIFIKATE (LocalMachine\My)" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    $zertifikate = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
        Select-Object Subject, NotBefore, NotAfter,
            @{ Name = "TageNoch"; Expression = { [int](($_.NotAfter - (Get-Date)).TotalDays) } },
            Thumbprint |
        Sort-Object TageNoch

    if ($zertifikate -and @($zertifikate).Count -gt 0) {
        Write-Host ("  " + "Tage".PadRight(8) + "Gueltig_bis".PadRight(14) + "Betreff") -ForegroundColor Cyan
        Write-Host ("  " + "-" * 62) -ForegroundColor Gray

        foreach ($z in $zertifikate) {
            $tage    = $z.TageNoch
            $ablauf  = $z.NotAfter.ToString("dd.MM.yyyy")
            $betreff = $z.Subject
            if ($betreff.Length -gt 42) { $betreff = $betreff.Substring(0, 40) + ".." }

            $farbe = if    ($tage -lt 0)   { "Red"    }
                     elseif ($tage -lt 30)  { "Red"    }
                     elseif ($tage -lt 90)  { "Yellow" }
                     else                   { "Green"  }

            $tageStr = if ($tage -lt 0) { "ABGELAUFEN" } else { "$tage Tage" }

            Write-Host ("  " + $tageStr.PadRight(8)) -NoNewline -ForegroundColor $farbe
            Write-Host ($ablauf.PadRight(14)) -NoNewline -ForegroundColor Gray
            Write-Host $betreff -ForegroundColor White
        }
    } else {
        Write-Host "  Keine_Zertifikate_im_Speicher_LocalMachine\My_gefunden" -ForegroundColor Gray
    }

    # Zusaetzlich: Zertifikate die in 90 Tagen ablaufen - auch aus anderen Stores
    Write-Host ""
    Write-Host "  ZERTIFIKATE BALD ABLAUFEND (alle Stores, < 90 Tage)" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor Gray

    $alleStores = @("Cert:\LocalMachine\My", "Cert:\LocalMachine\Root", "Cert:\LocalMachine\CA", "Cert:\CurrentUser\My")
    $baldAblaufend = @()

    foreach ($store in $alleStores) {
        $certs = Get-ChildItem -Path $store -ErrorAction SilentlyContinue |
            Where-Object { $_.NotAfter -lt (Get-Date).AddDays(90) } |
            Select-Object Subject, NotAfter, Thumbprint,
                @{ Name = "Store";    Expression = { $store } },
                @{ Name = "TageNoch"; Expression = { [int](($_.NotAfter - (Get-Date)).TotalDays) } }
        if ($certs) { $baldAblaufend += $certs }
    }

    if ($baldAblaufend.Count -gt 0) {
        Write-Host ("  " + "Tage".PadRight(12) + "Ablauf".PadRight(14) + "Store".PadRight(22) + "Betreff") -ForegroundColor Cyan
        foreach ($z in ($baldAblaufend | Sort-Object TageNoch)) {
            $tage    = $z.TageNoch
            $ablauf  = $z.NotAfter.ToString("dd.MM.yyyy")
            $betreff = $z.Subject
            $store   = $z.Store -replace "Cert:\\", "" -replace "\\", "\"
            if ($betreff.Length -gt 28) { $betreff = $betreff.Substring(0, 26) + ".." }
            $farbe   = if ($tage -lt 0) { "Red" } elseif ($tage -lt 30) { "Red" } else { "Yellow" }
            $tageStr = if ($tage -lt 0) { "ABGELAUFEN" } else { "$tage Tage" }

            Write-Host ("  " + $tageStr.PadRight(12)) -NoNewline -ForegroundColor $farbe
            Write-Host ($ablauf.PadRight(14)) -NoNewline -ForegroundColor Gray
            Write-Host ($store.PadRight(22)) -NoNewline -ForegroundColor DarkGray
            Write-Host $betreff -ForegroundColor White
        }
    } else {
        Write-Host "  Keine_Zertifikate_in_den_naechsten_90_Tagen_ablaufend" -ForegroundColor Green
    }

    Write-Host ""
    Pause-Script
}

# =============================================================
# HAUPTMENUE
# =============================================================
function Show-MainMenu {
    Show-Header
    Write-Host "  Modul waehlen:" -ForegroundColor White
    Write-Host ""
    Write-Host "  [1]  Software-Installation" -ForegroundColor White
    Write-Host "       OneCalendar, OneCommander, Vivaldi, Brave, Firefox," -ForegroundColor Gray
    Write-Host "       Bitwarden, TeamViewer, Thunderbird, pCloud, NAPS2, VLC, LibreOffice" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [2]  Administrator-Konto" -ForegroundColor White
    Write-Host "       Aktivieren, Passwort setzen, deaktivieren" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [3]  Benutzerrechte" -ForegroundColor White
    Write-Host "       Uebersicht (Admin/Standard) + Rechte aendern" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [4]  Boot-Nachrichten" -ForegroundColor White
    Write-Host "       Verbose_Boot + Schnellstart konfigurieren" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [5]  Scrollbalken" -ForegroundColor White
    Write-Host "       Immer sichtbar + 3x breiter" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [6]  Microsoft_Bloatware entfernen" -ForegroundColor White
    Write-Host "       OneDrive, Outlook, Copilot, Cortana, Teams, Bing, Xbox..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [7]  Sicherheits-Check" -ForegroundColor White
    Write-Host "       Defender, Firewall, UAC, RDP, Gastkonto, System_Restore" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [8]  Hardware & Netzwerk-Info" -ForegroundColor White
    Write-Host "       Festplatten, RAM, Netzwerkkarten, IP, VPN, SMART" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [9]  System-Diagnose" -ForegroundColor White
    Write-Host "       BitLocker, UEFI/Secure Boot, TPM, Zertifikat-Laufzeiten" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [A]  Alle Module ausfuehren" -ForegroundColor Cyan
    Write-Host "  [0]  Beenden" -ForegroundColor Gray
    Write-Host ""
    return Read-Host "  Auswahl"
}

# =============================================================
# START
# =============================================================
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "FEHLER: Als Administrator ausfuehren!" -ForegroundColor Red
    Write-Host "Rechtsklick auf Start.bat -> Als_Administrator_ausfuehren" -ForegroundColor Yellow
    Pause-Script
    exit
}

do {
    $auswahl = Show-MainMenu
    switch ($auswahl) {
        "1" { Install-Software }
        "2" { Setup-Administrator }
        "3" { Setup-Benutzerrechte }
        "4" { Setup-BootNachrichten }
        "5" { Setup-Scrollbalken }
        "6" { Remove-Bloatware }
        "7" { Check-Sicherheit }
        "8" { Show-HardwareInfo }
        "9" { Show-SystemDiagnose }
        { $_ -eq "A" -or $_ -eq "a" } {
            Install-Software; Setup-Administrator; Setup-Benutzerrechte
            Setup-BootNachrichten; Setup-Scrollbalken; Remove-Bloatware
            Check-Sicherheit; Show-HardwareInfo; Show-SystemDiagnose
        }
        "0" {
            Show-Header
            Write-Host "  Neustart empfohlen!" -ForegroundColor Yellow
            $neu = Read-Host "  Jetzt neu starten? (ja/nein)"
            Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
            Show-Info "ExecutionPolicy zurueckgesetzt auf: Restricted"
            if ($neu -eq "ja") { Restart-Computer -Force }
        }
    }
} while ($auswahl -ne "0")
