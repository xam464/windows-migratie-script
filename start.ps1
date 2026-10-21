# Starter voor de Migratie-Assistent. Gebruik in PowerShell:
#
#   irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1 | iex
#
# Met opties (bijvoorbeeld -Cli voor de tekstversie, of -Doel E:\):
#
#   & ([scriptblock]::Create((irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1))) -Cli
#
# Dit bestand haalt Migratie-Assistent.ps1 op als bestand en start het met Windows PowerShell.
# Als bestand blijven de Nederlandse tekens (UTF-8 met BOM) intact, en kan het script zichzelf in
# de back-up meenemen voor het herstellen. Dit bestand zelf moet puur ASCII blijven: irm | iex
# leest het zonder tekenset.
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Opties)

# In een eigen scriptblok, zodat er niets achterblijft in de sessie van de gebruiker
& {
    param([string[]]$Opties)

    if ($env:OS -ne 'Windows_NT') {
        Write-Host 'De Migratie-Assistent draait alleen op Windows.' -ForegroundColor Red
        return
    }

    # Oudere Windows 10-installaties gebruiken standaard nog geen TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor
        [Net.SecurityProtocolType]::Tls12

    $url = 'https://github.com/xam464/windows-migratie-script/releases/latest/download/Migratie-Assistent.ps1'
    $map = Join-Path $env:TEMP 'Migratie-Assistent'
    $pad = Join-Path $map 'Migratie-Assistent.ps1'
    New-Item -ItemType Directory -Force -Path $map | Out-Null

    # Vlak na een nieuwe release hangt het script er soms nog niet aan: dan een paar keer opnieuw proberen
    Write-Host 'Migratie-Assistent downloaden...' -ForegroundColor Cyan
    for ($poging = 1; $true; $poging++) {
        try {
            (New-Object Net.WebClient).DownloadFile($url, $pad)
            break
        } catch {
            if ($poging -ge 4) {
                Write-Host "Downloaden mislukt: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
            Write-Host 'Downloaden mislukt; over 10 seconden opnieuw...' -ForegroundColor Yellow
            Start-Sleep -Seconds 10
        }
    }

    # Een eigen Windows PowerShell-proces in hetzelfde venster: de execution policy speelt dan
    # geen rol, en het script draait als bestand in plaats van als losse tekst
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & $ps -NoProfile -ExecutionPolicy Bypass -File $pad @Opties
} $Opties
