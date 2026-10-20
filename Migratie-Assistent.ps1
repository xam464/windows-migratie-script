<#
.SYNOPSIS
    Migratie-assistent: inventariseert en back-upt een Windows 10-pc om over te zetten naar een nieuwe laptop.
.DESCRIPTION
    - Geïnstalleerde programma's (register, Store-apps, winget export)
    - Browsers: bladwijzers (ook als HTML), instellingen, extensielijst
      (Chrome, Edge, Brave, Vivaldi, Opera, Firefox)
    - Begeleide wachtwoord-export uit de browser, daarna desgewenst versleuteld met 7-Zip (AES-256)
    - Programma-instellingen uit AppData (Outlook, Office, VS Code, Notepad++, Thunderbird, PuTTY, ...)
    - Home-map scannen (grootte, aantal bestanden, grootste bestanden) en kopiëren met robocopy,
      met logging en controle achteraf
    Alles komt in één map: <Doel>\Migratie_<PC>_<datum>\ met rapport.html en HERSTELLEN.txt.
    Er wordt alleen gekopieerd; op de oude pc wordt niets verwijderd.

    Herstelmodus (op de nieuwe laptop): installeert de programma's met winget en zet instellingen,
    bladwijzers, het Firefox-profiel en bestanden terug vanuit zo'n back-upmap. Het script kopieert
    zichzelf in de back-upmap; start het daar en het biedt vanzelf aan om te herstellen.

    Wordt de nieuwe computer een Linux-pc? Stap L (of -NaarLinux) vertaalt de programma's naar
    Linux-apps op Flathub, zet PuTTY/WinSCP-sessies om naar ~/.ssh/config en schrijft
    Linux\herstel-linux.sh in de back-up. Op Linux: bash Linux/herstel-linux.sh
.PARAMETER Doel
    Doelmap: externe schijf of NAS-share. Zonder parameter wordt erom gevraagd.
.PARAMETER Herstel
    Back-upmap (Migratie_<PC>_<datum>) om op deze laptop terug te zetten.
.PARAMETER Gui
    Grafische dialogen gebruiken (mapkeuze, Out-GridView voor selecties).
.PARAMETER Alles
    Alle stappen achter elkaar doorlopen, zonder menu.
.PARAMETER NaarLinux
    De nieuwe computer wordt een Linux-pc: met -Alles komt de Linux-voorbereiding er aan het eind bij.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Migratie-Assistent.ps1 -Gui
.EXAMPLE
    .\Migratie-Assistent.ps1 -Doel E:\ -Alles
.EXAMPLE
    .\Migratie-Assistent.ps1 -Herstel E:\Migratie_OUDEPC_20261014_0930
.EXAMPLE
    .\Migratie-Assistent.ps1 -Doel E:\ -Alles -NaarLinux
#>
[CmdletBinding()]
param(
    [string]$Doel,
    [string]$Herstel,
    [switch]$Gui,
    [switch]$Alles,
    [switch]$NaarLinux
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Configuratie
# ---------------------------------------------------------------------------

# Chromium-browsers: waar de gebruikersdata staat en waar de wachtwoord-export zit
$script:Chromium = @(
    @{ Naam = 'Chrome';  Proces = 'chrome';  Data = "$env:LOCALAPPDATA\Google\Chrome\User Data";               Ww = 'chrome://password-manager/settings' }
    @{ Naam = 'Edge';    Proces = 'msedge';  Data = "$env:LOCALAPPDATA\Microsoft\Edge\User Data";              Ww = 'edge://settings/passwords' }
    @{ Naam = 'Brave';   Proces = 'brave';   Data = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"; Ww = 'brave://password-manager/settings' }
    @{ Naam = 'Vivaldi'; Proces = 'vivaldi'; Data = "$env:LOCALAPPDATA\Vivaldi\User Data";                     Ww = 'vivaldi://settings/passwords' }
    @{ Naam = 'Opera';   Proces = 'opera';   Data = "$env:APPDATA\Opera Software\Opera Stable";                Ww = 'opera://settings/passwords' }
)

# Bestanden per Chromium-profiel die de moeite waard zijn. 'Login Data' (wachtwoorden) en
# 'Secure Preferences' zijn aan deze Windows-installatie gebonden en hebben elders geen zin.
$script:ChromiumBestanden = @('Bookmarks', 'Bookmarks.bak', 'Preferences', 'History', 'Favicons',
                              'Top Sites', 'Shortcuts', 'Web Data', 'Custom Dictionary.txt')

# Mappen in de home-map die standaard niet gekopieerd worden
$script:HomeOverslaan = @('AppData', 'MicrosoftEdgeBackups', 'IntelGraphicsProfiles')

# Programma-instellingen buiten de home-map. Paden blijven met %VARIABELEN% bewaard, zodat de
# herstelstap ze op de nieuwe laptop (mogelijk andere gebruikersnaam) op de juiste plek terugzet.
# Soort: Map (robocopy, eventueel met Filter), Register (reg export/import), Lettertypen.
$script:Instellingen = @(
    @{ Naam = 'Outlook-handtekeningen'; Soort = 'Map';         Pad = '%APPDATA%\Microsoft\Signatures';  Proces = 'outlook' }
    @{ Naam = 'Outlook PST-bestanden';  Soort = 'Map';         Pad = '%LOCALAPPDATA%\Microsoft\Outlook'; Proces = 'outlook'; Filter = '*.pst' }
    @{ Naam = 'Office-sjablonen';       Soort = 'Map';         Pad = '%APPDATA%\Microsoft\Templates';   Proces = 'winword' }
    @{ Naam = 'Office-woordenlijsten';  Soort = 'Map';         Pad = '%APPDATA%\Microsoft\UProof';      Proces = 'winword' }
    @{ Naam = 'Notepad++';              Soort = 'Map';         Pad = '%APPDATA%\Notepad++';             Proces = 'notepad++' }
    @{ Naam = 'VS Code';                Soort = 'Map';         Pad = '%APPDATA%\Code\User';             Proces = 'code'; Uitsluiten = @('workspaceStorage', 'History') }
    @{ Naam = 'Thunderbird';            Soort = 'Map';         Pad = '%APPDATA%\Thunderbird';           Proces = 'thunderbird' }
    # eM Client: de hele database. Accounts, serverinstellingen en de (versleutelde) wachtwoorden
    # staan daarin en werken na terugzetten op een andere pc. Zoek: een eigen databasemap.
    @{ Naam = 'eM Client';              Soort = 'Map';         Pad = '%APPDATA%\eM Client';             Proces = 'MailClient'
       Zoek = { Get-EmClientMap }; Vervangen = $true; MoetDicht = $true }
    @{ Naam = 'FileZilla';              Soort = 'Map';         Pad = '%APPDATA%\FileZilla';             Proces = 'filezilla' }
    @{ Naam = 'Windows Terminal';       Soort = 'Map';         Pad = '%LOCALAPPDATA%\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState' }
    @{ Naam = 'Sticky Notes';           Soort = 'Map';         Pad = '%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState'; Proces = 'Microsoft.Notes' }
    @{ Naam = 'Eigen lettertypen';      Soort = 'Lettertypen'; Pad = '%LOCALAPPDATA%\Microsoft\Windows\Fonts' }
    @{ Naam = 'PuTTY-sessies';          Soort = 'Register';    Pad = 'HKCU\Software\SimonTatham\PuTTY';  Proces = 'putty' }
    @{ Naam = 'WinSCP-sessies';         Soort = 'Register';    Pad = 'HKCU\Software\Martin Prikryl\WinSCP 2'; Proces = 'winscp' }
)

# Mappen in de back-up (Home\<naam>) die op de nieuwe laptop naar een bekende map horen;
# die kan omgeleid zijn, bijvoorbeeld naar OneDrive
# (ook onder de Nederlandse naam, voor wie een map zelf in de back-up heeft gezet)
$script:BekendeMappen = @{
    Desktop = 'Desktop'; Documents = 'MyDocuments'; Pictures = 'MyPictures'
    Music = 'MyMusic'; Videos = 'MyVideos'; Favorites = 'Favorites'
    Bureaublad = 'Desktop'; Documenten = 'MyDocuments'; Afbeeldingen = 'MyPictures'
    Muziek = 'MyMusic'; "Video's" = 'MyVideos'
}

# Windows-programma -> Linux. Patroon is een regex op de programmanaam; Id is de Flatpak-ID op
# Flathub. Zonder Id is er alleen een advies (Opm). Cmd: staat dit commando al op de Linux-pc
# (meegeleverd door de distributie), dan installeert het herstelscript niets. Uit: niet aangevinkt.
$script:LinuxApps = @(
    # Browsers en communicatie
    @{ Patroon = 'Mozilla Firefox';                Id = 'org.mozilla.firefox';          Naam = 'Firefox';           Cmd = 'firefox' }
    @{ Patroon = 'Google Chrome';                  Id = 'com.google.Chrome';            Naam = 'Google Chrome';     Cmd = 'google-chrome' }
    @{ Patroon = '^Microsoft Edge$';               Id = 'com.microsoft.Edge';           Naam = 'Microsoft Edge';    Cmd = 'microsoft-edge'; Uit = $true }
    @{ Patroon = '^Brave';                         Id = 'com.brave.Browser';            Naam = 'Brave';             Cmd = 'brave-browser' }
    @{ Patroon = '^Vivaldi';                       Id = 'com.vivaldi.Vivaldi';          Naam = 'Vivaldi';           Cmd = 'vivaldi' }
    @{ Patroon = '^Opera';                         Id = 'com.opera.Opera';              Naam = 'Opera';             Cmd = 'opera' }
    @{ Patroon = 'Thunderbird';                    Id = 'org.mozilla.Thunderbird';      Naam = 'Thunderbird';       Cmd = 'thunderbird' }
    @{ Patroon = '^eM Client';                     Id = 'org.mozilla.Thunderbird';      Naam = 'Thunderbird (alternatief)'; Cmd = 'thunderbird'; Opm = 'eM Client bestaat niet voor Linux' }
    @{ Patroon = 'Microsoft (Office|365)';         Id = 'org.libreoffice.LibreOffice';  Naam = 'LibreOffice (alternatief)'; Cmd = 'libreoffice'; Opm = 'Of Office online via office.com' }
    @{ Patroon = 'Teams';                          Id = 'com.github.IsmaelMartinez.teams_for_linux'; Naam = 'Teams for Linux (onofficieel)'; Opm = 'Of teams.microsoft.com in de browser' }
    @{ Patroon = '^Zoom';                          Id = 'us.zoom.Zoom';                 Naam = 'Zoom' }
    @{ Patroon = 'Slack';                          Id = 'com.slack.Slack';              Naam = 'Slack' }
    @{ Patroon = 'Discord';                        Id = 'com.discordapp.Discord';       Naam = 'Discord' }
    @{ Patroon = 'Signal';                         Id = 'org.signal.Signal';            Naam = 'Signal' }
    @{ Patroon = 'Telegram';                       Id = 'org.telegram.desktop';         Naam = 'Telegram' }
    @{ Patroon = '^Element';                       Id = 'im.riot.Riot';                 Naam = 'Element' }
    @{ Patroon = 'WhatsApp';                       Opm = 'Geen Linux-app: gebruik web.whatsapp.com' }
    # Media en creatief
    @{ Patroon = 'VLC';                            Id = 'org.videolan.VLC';             Naam = 'VLC';               Cmd = 'vlc' }
    @{ Patroon = 'Spotify';                        Id = 'com.spotify.Client';           Naam = 'Spotify' }
    @{ Patroon = 'Kodi';                           Id = 'tv.kodi.Kodi';                 Naam = 'Kodi' }
    @{ Patroon = '^Plex';                          Id = 'tv.plex.PlexDesktop';          Naam = 'Plex' }
    @{ Patroon = 'iTunes|Apple Music';             Opm = 'Gebruik Rhythmbox of Lollypop; je muziek staat in de map Music' }
    @{ Patroon = 'Audacity';                       Id = 'org.audacityteam.Audacity';    Naam = 'Audacity' }
    @{ Patroon = 'OBS Studio';                     Id = 'com.obsproject.Studio';        Naam = 'OBS Studio' }
    @{ Patroon = 'HandBrake';                      Id = 'fr.handbrake.ghb';             Naam = 'HandBrake' }
    @{ Patroon = 'GIMP';                           Id = 'org.gimp.GIMP';                Naam = 'GIMP';              Cmd = 'gimp' }
    @{ Patroon = 'Inkscape';                       Id = 'org.inkscape.Inkscape';        Naam = 'Inkscape' }
    @{ Patroon = 'Krita';                          Id = 'org.kde.krita';                Naam = 'Krita' }
    @{ Patroon = 'darktable';                      Id = 'org.darktable.Darktable';      Naam = 'darktable' }
    @{ Patroon = 'Blender';                        Id = 'org.blender.Blender';          Naam = 'Blender' }
    @{ Patroon = 'Shotcut';                        Id = 'org.shotcut.Shotcut';          Naam = 'Shotcut' }
    @{ Patroon = 'Kdenlive';                       Id = 'org.kde.kdenlive';             Naam = 'Kdenlive' }
    @{ Patroon = 'paint\.net';                     Id = 'com.github.PintaProject.Pinta'; Naam = 'Pinta (alternatief)' }
    @{ Patroon = 'Photoshop|Lightroom';            Opm = 'Geen Linux-versie; alternatief: GIMP, Krita of darktable' }
    @{ Patroon = 'Acrobat|Adobe Reader|Foxit|PDF-XChange'; Opm = 'Linux heeft een ingebouwde pdf-lezer; bewerken kan met Okular of Xournal++' }
    @{ Patroon = 'calibre';                        Id = 'com.calibre_ebook.calibre';    Naam = 'calibre' }
    # Hulpmiddelen
    @{ Patroon = '^7-Zip|WinRAR|PeaZip';           Opm = 'Archieven openen zit in Linux; het herstelscript installeert 7-Zip zelf als het nodig is' }
    @{ Patroon = 'Notepad\+\+';                    Id = 'org.kde.kate';                 Naam = 'Kate (alternatief)' }
    @{ Patroon = 'Visual Studio Code';             Id = 'com.visualstudio.code';        Naam = 'VS Code';           Cmd = 'code' }
    @{ Patroon = 'Visual Studio (Community|Professional|Enterprise)'; Opm = 'Geen Linux-versie; gebruik VS Code of JetBrains Rider' }
    @{ Patroon = 'KeePass';                        Id = 'org.keepassxc.KeePassXC';      Naam = 'KeePassXC';         Cmd = 'keepassxc' }
    @{ Patroon = 'Bitwarden';                      Id = 'com.bitwarden.desktop';        Naam = 'Bitwarden' }
    @{ Patroon = '1Password';                      Opm = 'Officiële Linux-app via 1password.com (deb/rpm)' }
    @{ Patroon = 'FileZilla';                      Id = 'org.filezillaproject.Filezilla'; Naam = 'FileZilla';       Cmd = 'filezilla' }
    @{ Patroon = 'WinSCP';                         Id = 'org.filezillaproject.Filezilla'; Naam = 'FileZilla (alternatief)'; Cmd = 'filezilla'; Opm = 'Of sftp:// in je bestandsbeheer' }
    @{ Patroon = 'PuTTY';                          Opm = 'ssh zit in Linux; je sessies staan in Linux\ssh-config.txt' }
    @{ Patroon = 'ShareX|Greenshot|Lightshot';     Id = 'org.flameshot.Flameshot';      Naam = 'Flameshot (alternatief)' }
    @{ Patroon = '^Everything';                    Id = 'io.github.cboxdoerfer.FSearch'; Naam = 'FSearch (alternatief)' }
    @{ Patroon = 'Rufus|balenaEtcher';             Id = 'io.gitlab.adhami3310.Impression'; Naam = 'Impression (alternatief)' }
    @{ Patroon = 'Wireshark';                      Id = 'org.wireshark.Wireshark';      Naam = 'Wireshark' }
    @{ Patroon = 'AnyDesk';                        Id = 'com.anydesk.Anydesk';          Naam = 'AnyDesk' }
    @{ Patroon = 'RustDesk';                       Id = 'com.rustdesk.RustDesk';        Naam = 'RustDesk' }
    @{ Patroon = 'TeamViewer';                     Opm = 'Officieel pakket (deb/rpm) via teamviewer.com' }
    @{ Patroon = 'Remote Desktop';                 Id = 'org.remmina.Remmina';          Naam = 'Remmina (alternatief)' }
    @{ Patroon = 'Dropbox';                        Id = 'com.dropbox.Client';           Naam = 'Dropbox' }
    @{ Patroon = 'Nextcloud';                      Id = 'com.nextcloud.desktopclient.nextcloud'; Naam = 'Nextcloud' }
    @{ Patroon = 'OneDrive';                       Opm = 'Geen officiële client: via onedrive.com, of onofficieel "onedriver"' }
    @{ Patroon = 'Google Drive';                   Opm = 'Geen officiële client: via de browser of GNOME Online-accounts' }
    @{ Patroon = 'Synology Drive';                 Opm = 'Officieel pakket (deb) via synology.com' }
    @{ Patroon = 'Tailscale';                      Opm = 'curl -fsSL https://tailscale.com/install.sh | sh' }
    @{ Patroon = 'WireGuard|OpenVPN';              Opm = 'Zit in NetworkManager (Instellingen > Netwerk > VPN)' }
    # Spellen
    @{ Patroon = '^Steam';                         Id = 'com.valvesoftware.Steam';      Naam = 'Steam';             Cmd = 'steam' }
    @{ Patroon = 'Epic Games|GOG Galaxy|Ubisoft|EA app|Battle\.net'; Id = 'com.heroicgameslauncher.hgl'; Naam = 'Heroic (alternatief)'; Opm = 'Niet elk spel werkt op Linux' }
    @{ Patroon = 'Minecraft';                      Id = 'org.prismlauncher.PrismLauncher'; Naam = 'Prism Launcher' }
    # Ontwikkelen
    @{ Patroon = '^Git( |$)';                      Opm = 'Pakket "git" van je distributie' }
    @{ Patroon = '^Python \d';                     Opm = 'Python zit in Linux' }
    @{ Patroon = 'Docker Desktop';                 Opm = 'Docker Engine via je distributie (docker.io of docker-ce)' }
    @{ Patroon = 'Node\.js';                       Opm = 'Pakket "nodejs" van je distributie, of nvm' }
    @{ Patroon = 'JetBrains Toolbox';              Opm = 'Toolbox voor Linux via jetbrains.com' }
    @{ Patroon = 'IntelliJ IDEA';                  Id = 'com.jetbrains.IntelliJ-IDEA-Community'; Naam = 'IntelliJ IDEA Community' }
    @{ Patroon = 'PyCharm';                        Id = 'com.jetbrains.PyCharm-Community'; Naam = 'PyCharm Community' }
    @{ Patroon = 'Android Studio';                 Id = 'com.google.AndroidStudio';     Naam = 'Android Studio' }
    @{ Patroon = 'DBeaver';                        Id = 'io.dbeaver.DBeaverCommunity';  Naam = 'DBeaver' }
    @{ Patroon = 'Postman';                        Id = 'com.getpostman.Postman';       Naam = 'Postman' }
    @{ Patroon = 'Arduino';                        Id = 'cc.arduino.IDE2';              Naam = 'Arduino IDE' }
    @{ Patroon = 'Obsidian';                       Id = 'md.obsidian.Obsidian';         Naam = 'Obsidian' }
    @{ Patroon = 'Joplin';                         Id = 'net.cozic.joplin_desktop';     Naam = 'Joplin' }
    # 3D-printen
    @{ Patroon = '\bCura\b';                      Id = 'com.ultimaker.cura';           Naam = 'UltiMaker Cura' }
    @{ Patroon = 'PrusaSlicer';                    Id = 'com.prusa3d.PrusaSlicer';      Naam = 'PrusaSlicer' }
    @{ Patroon = 'Bambu Studio';                   Id = 'com.bambulab.BambuStudio';     Naam = 'Bambu Studio' }
    @{ Patroon = 'OrcaSlicer';                     Id = 'io.github.softfever.OrcaSlicer'; Naam = 'OrcaSlicer' }
)

# Stuurprogramma's, runtimes en hulpprogramma's van de fabrikant: op Linux niet nodig
$script:LinuxOverslaan = 'Visual C\+\+|\.NET|Runtime|Redistributable|Driver|Stuurprogramma|Realtek|Intel\(R\)|' +
    'NVIDIA|AMD Software|Synaptics|Windows SDK|WebView2|Update Health|Edge Update|Machine-Wide|' +
    'Click-to-Run|Bonjour|Dolby|Chipset|Thunderbolt|PC Health|GameInput|Teams Meeting Add-in'

$script:Run = $null
$script:Scan = $null
$script:Bron = $null
$script:HerstelLog = $null
$script:Rapport = [ordered]@{
    Computer   = $env:COMPUTERNAME
    Gebruiker  = $env:USERNAME
    Datum      = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    Windows    = ''
    Programmas = @()
    Browsers   = @()
    Home       = @()
    Grootste   = @()
    Kopieen    = @()
    Instellingen = @()
    Linux      = @()
}

$script:WachtwoordUitleg = @'
  Browsers versleutelen opgeslagen wachtwoorden met een sleutel die aan deze Windows-installatie
  gebonden is (DPAPI, en bij Chrome/Edge sinds 2024 ook 'app-bound encryption'). Alleen de
  bestanden kopiëren werkt op de nieuwe laptop dus niet: je exporteert via de browser zelf naar CSV.

  Het script probeert dat exporteren voor je te automatiseren: het opent de pagina en klikt zelf
  door. Jij hoeft alleen nog je pincode of wachtwoord in te vullen bij de Windows-prompt. Lukt het
  automatisch klikken niet (andere browserversie of taal), dan verschijnen de handmatige stappen.

  Let op: zo'n CSV bevat al je wachtwoorden in leesbare tekst. Het script vraagt daarna of je hem
  met 7-Zip (AES-256) wilt versleutelen; zo ja, dan wordt het origineel verwijderd.
  Tip: een wachtwoordmanager (Bitwarden/Vaultwarden, KeePassXC) kan zo'n CSV direct importeren.

'@

$script:HerstelTekst = @'
HERSTELLEN OP DE NIEUWE LAPTOP
==============================

Nieuwe computer met Linux? Zie LINUX.txt (als je stap L 'Voorbereiden op Linux' gedaan hebt).

Snelste weg: start het script uit deze map op de nieuwe laptop en kies 'herstellen':
   powershell -ExecutionPolicy Bypass -File .\Migratie-Assistent.ps1
Het installeert de programma's en zet instellingen, bladwijzers, het Firefox-profiel en je
bestanden terug. Hieronder staat hoe je alles met de hand doet.

1. Programma's
   - winget-apps.json:  winget import -i winget-apps.json --accept-package-agreements
   - winget-niet-beschikbaar.txt: deze programma's staan niet in winget en moet je zelf installeren.
   - programmas.csv: de volledige lijst met versies en uitgevers (opent in Excel).

2. Bladwijzers
   - Browsers\<browser>\<profiel>\bladwijzers.html: in elke browser te importeren via
     Bladwijzers > Bladwijzers en instellingen importeren > HTML-bestand.
   - Zelfde browser? Sluit hem af en kopieer het bestand 'Bookmarks' naar het nieuwe profiel
     (bijv. %LOCALAPPDATA%\Google\Chrome\User Data\Default).

3. Extensies
   - extensies.csv per profiel: naam en ID. Installeer ze opnieuw uit de webwinkel.

4. Firefox
   - Start Firefox één keer en sluit hem weer. Open about:profiles, klik bij het actieve profiel
     op 'Map openen' (Hoofdmap), sluit Firefox en kopieer de inhoud van Browsers\Firefox\<profiel>
     daaroverheen. Bladwijzers, wachtwoorden, extensies en instellingen komen dan mee.

5. Wachtwoorden
   - Pak Wachtwoorden\wachtwoorden_*.7z uit met 7-Zip. Koos je bij de back-up voor niet
     versleutelen, dan staan de CSV-bestanden direct in de map Wachtwoorden.
   - Importeer de CSV in de browser (chrome://password-manager/settings > Importeren,
     edge://settings/passwords > Importeren) of in je wachtwoordmanager.
   - Verwijder de uitgepakte CSV daarna meteen.

6. Bestanden
   - Home\<map> terugzetten, bijvoorbeeld:
     robocopy "<back-up>\Home\Documents" "%USERPROFILE%\Documents" /E
   - robocopy.log bevat welke bestanden gekopieerd zijn en welke fouten er waren.

7. Programma-instellingen
   - Instellingen\instellingen.json zegt per onderdeel waar het vandaan kwam (bijv.
     %APPDATA%\Notepad++). Sluit het programma en kopieer de map terug naar die plek.
   - .reg-bestanden (PuTTY, WinSCP): dubbelklikken om te importeren.
   - Lettertypen: selecteer de bestanden, rechtermuisknop > Installeren.
   - vscode-extensies.txt: per regel  code --install-extension <naam>
'@

# ---------------------------------------------------------------------------
# Hulpfuncties
# ---------------------------------------------------------------------------

function Write-Kop([string]$Tekst) { Write-Host ''; Write-Host "== $Tekst ==" -ForegroundColor Cyan }
function Write-Ok([string]$Tekst)   { Write-Host "  [OK] $Tekst" -ForegroundColor Green }
function Write-Let([string]$Tekst)  { Write-Host "  [!]  $Tekst" -ForegroundColor Yellow }
function Write-Fout([string]$Tekst) { Write-Host "  [X]  $Tekst" -ForegroundColor Red }

function Format-Grootte([double]$Bytes) {
    if ($Bytes -ge 1TB) { return '{0:N1} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Read-JaNee([string]$Vraag) {
    return (Read-Host "  $Vraag (j/N)") -match '^(j|ja|y|yes)$'
}

# Maakt (eenmalig) de back-upmap aan op de doellocatie
function Get-RunMap {
    if ($script:Run) { return $script:Run }

    if (-not $script:Doel) {
        if ($Gui) {
            Add-Type -AssemblyName System.Windows.Forms
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Kies de doellocatie voor de migratie-back-up (externe schijf of NAS)'
            $dlg.ShowNewFolderButton = $true
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'Geen doelmap gekozen.' }
            $script:Doel = $dlg.SelectedPath
        } else {
            $script:Doel = Read-Host '  Doelmap (bijv. E:\ of \\nas\backup)'
        }
    }
    if (-not $script:Doel -or -not (Test-Path -LiteralPath $script:Doel)) { throw "Doelmap '$script:Doel' bestaat niet." }

    $volledig = (Resolve-Path -LiteralPath $script:Doel).ProviderPath
    if ([IO.Path]::GetPathRoot($volledig) -eq "$env:SystemDrive\") {
        Write-Let "De doelmap staat op dezelfde schijf als Windows ($env:SystemDrive). Dat is geen veilige plek als deze laptop het begeeft."
        if (-not (Read-JaNee 'Toch doorgaan?')) { $script:Doel = $null; throw 'Afgebroken: kies een externe schijf of NAS.' }
    }

    $naam = 'Migratie_{0}_{1}' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmm')
    $script:Run = (New-Item -ItemType Directory -Force -Path (Join-Path $volledig $naam)).FullName
    try { Start-Transcript -Path (Join-Path $script:Run 'transcript.log') | Out-Null } catch { }
    Set-Content -LiteralPath (Join-Path $script:Run 'HERSTELLEN.txt') -Value $script:HerstelTekst -Encoding UTF8
    # Het script zelf meenemen, zodat het op de nieuwe laptop de herstelmodus kan draaien
    if ($PSCommandPath) { Copy-Item -LiteralPath $PSCommandPath -Destination $script:Run -Force }
    Write-Ok "Back-upmap: $script:Run"
    return $script:Run
}

# Laat de gebruiker items kiezen: Out-GridView met -Gui, anders een genummerde lijst
function Select-Items {
    param(
        [object[]]$Items,
        [string]$Titel,
        [string[]]$Kolommen,
        [scriptblock]$Voorstel
    )
    if (-not $Items) { return @() }
    if ($Gui) { return @($Items | Out-GridView -Title $Titel -OutputMode Multiple) }

    $nr = 0
    foreach ($item in $Items) { $nr++; $item | Add-Member -NotePropertyName Nr -NotePropertyValue $nr -Force }
    $Items | Format-Table -Property (@('Nr') + $Kolommen) -AutoSize | Out-Host

    $hint = "nummers (bijv. 1,3,5-7), 'a' = alles"
    if ($Voorstel) { $hint += ", 'v' = voorstel" }
    $antw = Read-Host "  $Titel - $hint, leeg = niets"
    if (-not $antw) { return @() }
    if ($antw -eq 'a') { return $Items }
    if ($antw -eq 'v' -and $Voorstel) { return @($Items | Where-Object $Voorstel) }

    $nrs = foreach ($deel in $antw -split ',') {
        $deel = $deel.Trim()
        if ($deel -match '^(\d+)-(\d+)$') { [int]$Matches[1]..[int]$Matches[2] }
        elseif ($deel -match '^\d+$') { [int]$deel }
    }
    return @($Items | Where-Object { $nrs -contains $_.Nr })
}

# Telt bestanden en bytes recursief. Junctions en symbolische koppelingen worden overgeslagen (anders
# kan de telling rondjes lopen); OneDrive-mappen en -bestanden zijn ook reparse points, maar tellen
# wel mee. OnlineBytes: bestanden die alleen in de cloud staan en bij het kopiëren gedownload worden.
function Measure-Map([string]$Pad, [long]$GrootDrempel = 100MB) {
    $bytes = 0L
    $aantal = 0L
    $online = 0L
    $groot = New-Object System.Collections.Generic.List[object]
    $stapel = New-Object System.Collections.Generic.Stack[string]
    $stapel.Push($Pad)
    while ($stapel.Count -gt 0) {
        $dir = $stapel.Pop()
        try {
            foreach ($i in ([IO.DirectoryInfo]::new($dir)).EnumerateFileSystemInfos()) {
                if ($i -is [IO.DirectoryInfo]) {
                    if (($i.Attributes -band [IO.FileAttributes]::ReparsePoint) -and
                        (Get-Item -LiteralPath $i.FullName -Force -ErrorAction SilentlyContinue).LinkType) { continue }
                    $stapel.Push($i.FullName)
                    continue
                }
                $aantal++
                $bytes += $i.Length
                # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (0x400000): alleen online
                if ([int]$i.Attributes -band 0x400000) { $online += $i.Length }
                if ($i.Length -ge $GrootDrempel) {
                    $groot.Add([pscustomobject]@{ Pad = $i.FullName; Bytes = $i.Length; Grootte = (Format-Grootte $i.Length) })
                }
            }
        } catch {
            # Geen toegang of map verdwenen: overslaan
        }
    }
    return [pscustomobject]@{ Bytes = $bytes; Bestanden = $aantal; OnlineBytes = $online; Groot = $groot }
}

# Vraagt om een programma af te sluiten als het draait. Geeft $true als het (nu) dicht is.
function Confirm-Gesloten([string]$Naam, [string]$Proces) {
    if (-not $Proces) { return $true }
    $p = @(Get-Process -Name $Proces -ErrorAction SilentlyContinue)
    if (-not $p) { return $true }
    Write-Let "$Naam is nog open; de bestanden zijn dan (deels) vergrendeld."
    if (-not (Read-JaNee "$Naam nu afsluiten? Bij een browser komen je tabbladen terug bij de volgende start.")) { return $false }
    $p | ForEach-Object { [void]$_.CloseMainWindow() }
    Start-Sleep -Seconds 3
    Get-Process -Name $Proces -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    return -not (Get-Process -Name $Proces -ErrorAction SilentlyContinue)
}

# robocopy met vaste opties. Extra komt direct na bron en doel, dus een bestandsfilter (*.pst) mag
# vooraan staan. Geeft de exitcode terug: 0-7 is gelukt, 8 en hoger betekent fouten.
function Invoke-Robocopy([string]$Bron, [string]$Doel, [string[]]$Extra = @(), [string]$Log) {
    $a = @($Bron, $Doel) + $Extra + @('/R:1', '/W:1', '/XJD', '/NP', '/NDL')
    if ($Log) { $a += "/LOG+:$Log" } else { $a += @('/NFL', '/NJH', '/NJS') }
    & robocopy @a | Out-Null
    return $LASTEXITCODE
}

# Na een winget-installatie staat het nieuwe programma pas in PATH van een nieuwe sessie
function Update-Pad {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

# ---------------------------------------------------------------------------
# Programma's
# ---------------------------------------------------------------------------

function Get-Programmas {
    $sleutels = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty -Path $sleutels -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and -not $_.ParentKeyName -and $_.ReleaseType -notmatch 'Update|Hotfix' } |
        ForEach-Object {
            [pscustomobject]@{
                Naam          = $_.DisplayName.Trim()
                Versie        = $_.DisplayVersion
                Uitgever      = $_.Publisher
                Geinstalleerd = $_.InstallDate
                Bron          = if ($_.PSPath -like '*HKEY_CURRENT_USER*') { 'Gebruiker' } else { 'Machine' }
            }
        } |
        Sort-Object Naam -Unique
}

function Get-StoreApps {
    # Microsoft.*-pakketten zijn vrijwel allemaal standaard-apps van Windows; die laten we weg
    Get-AppxPackage -PackageTypeFilter Main -ErrorAction SilentlyContinue |
        Where-Object { -not $_.IsFramework -and -not $_.NonRemovable -and $_.SignatureKind -eq 'Store' -and $_.Name -notlike 'Microsoft.*' } |
        ForEach-Object {
            [pscustomobject]@{
                Naam          = $_.Name
                Versie        = $_.Version
                Uitgever      = $_.Publisher -replace '^CN=([^,]+).*', '$1'
                Geinstalleerd = ''
                Bron          = 'Store'
            }
        }
}

function Export-Programmas {
    $run = Get-RunMap
    Write-Kop "Geïnstalleerde programma's"

    $alle = @(Get-Programmas) + @(Get-StoreApps)
    $csv = Join-Path $run 'programmas.csv'
    $alle | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $script:Rapport.Programmas = $alle
    Write-Ok "$($alle.Count) programma's -> programmas.csv"

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        $json = Join-Path $run 'winget-apps.json'
        # In Windows PowerShell 5.1 wordt omgeleide stderr met 'Stop' een afbrekende fout
        $uitvoer = & { $ErrorActionPreference = 'Continue'; winget export -o $json --accept-source-agreements 2>&1 }
        $niet = @($uitvoer | Where-Object { "$_" -match 'not available|niet beschikbaar' })
        $niet | Set-Content -LiteralPath (Join-Path $run 'winget-niet-beschikbaar.txt') -Encoding UTF8
        if (Test-Path -LiteralPath $json) {
            Write-Ok "winget-export -> winget-apps.json ($($niet.Count) programma's niet via winget te herinstalleren)"
        } else {
            Write-Let 'winget export is mislukt; programmas.csv is er wel.'
        }
    } else {
        Write-Let 'winget is niet geïnstalleerd (App Installer uit de Store). Alleen programmas.csv gemaakt.'
    }

    if ($Gui) { $alle | Out-GridView -Title "Geïnstalleerde programma's ($($alle.Count))" }
}

# ---------------------------------------------------------------------------
# Browsers
# ---------------------------------------------------------------------------

function Get-ChromiumProfielNamen([string]$Data) {
    $namen = @{}
    $ls = Join-Path $Data 'Local State'
    if (Test-Path -LiteralPath $ls) {
        try {
            $j = Get-Content -LiteralPath $ls -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.profile.info_cache.PSObject.Properties) { $namen[$p.Name] = $p.Value.name }
        } catch { }
    }
    return $namen
}

function Get-ChromiumExtensies([string]$ProfielPad) {
    $ext = Join-Path $ProfielPad 'Extensions'
    if (-not (Test-Path -LiteralPath $ext)) { return }
    foreach ($idMap in Get-ChildItem -LiteralPath $ext -Directory | Where-Object { $_.Name -match '^[a-p]{32}$' }) {
        $versie = Get-ChildItem -LiteralPath $idMap.FullName -Directory | Sort-Object Name -Descending | Select-Object -First 1
        if (-not $versie) { continue }
        try { $m = Get-Content -LiteralPath (Join-Path $versie.FullName 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }

        # Gelokaliseerde naam ("__MSG_appName__") opzoeken in _locales
        $naam = $m.name
        if ($naam -like '__MSG_*') {
            $sleutel = $naam -replace '^__MSG_(.+)__$', '$1'
            $talen = @('nl', $m.default_locale, 'en') | Where-Object { $_ }
            foreach ($taal in $talen) {
                $msgPad = Join-Path $versie.FullName "_locales\$taal\messages.json"
                if (-not (Test-Path -LiteralPath $msgPad)) { continue }
                # Geen ConvertFrom-Json: dat weigert in 5.1 bestanden met sleutels die alleen in
                # hoofdletters verschillen, en die komen in extensies van Google voor
                $tekst = Get-Content -LiteralPath $msgPad -Raw -Encoding UTF8
                $patroon = '"' + [regex]::Escape($sleutel) + '"\s*:\s*\{[^}]*?"message"\s*:\s*"((?:[^"\\]|\\.)*)"'
                $hit = [regex]::Match($tekst, $patroon, 'IgnoreCase')
                if ($hit.Success) { $naam = [regex]::Unescape($hit.Groups[1].Value); break }
            }
        }
        [pscustomobject]@{
            Naam   = $naam
            Versie = $m.version
            Id     = $idMap.Name
            Winkel = "https://chromewebstore.google.com/detail/$($idMap.Name)"
        }
    }
}

function Add-BladwijzerKnoop($Knoop, [System.Text.StringBuilder]$Sb, [int]$Diepte) {
    $ins = '    ' * $Diepte
    $datum = 0
    # Chromium telt microseconden sinds 1601-01-01; het HTML-formaat wil Unix-seconden
    if ($Knoop.date_added) { $datum = [long][math]::Max(0, [double]$Knoop.date_added / 1e6 - 11644473600) }
    $naam = [Net.WebUtility]::HtmlEncode($Knoop.name)
    if ($Knoop.type -eq 'folder') {
        [void]$Sb.AppendLine("$ins<DT><H3 ADD_DATE=""$datum"">$naam</H3>")
        [void]$Sb.AppendLine("$ins<DL><p>")
        foreach ($kind in $Knoop.children) { Add-BladwijzerKnoop $kind $Sb ($Diepte + 1) }
        [void]$Sb.AppendLine("$ins</DL><p>")
    } elseif ($Knoop.type -eq 'url') {
        $url = [Net.WebUtility]::HtmlEncode($Knoop.url)
        [void]$Sb.AppendLine("$ins<DT><A HREF=""$url"" ADD_DATE=""$datum"">$naam</A>")
    }
}

# Zet een Chromium 'Bookmarks'-bestand om naar het Netscape-HTML-formaat dat elke browser importeert.
# Geeft het aantal bladwijzers terug.
function ConvertTo-BladwijzerHtml([string]$BookmarksJson, [string]$Uit) {
    $data = Get-Content -LiteralPath $BookmarksJson -Raw -Encoding UTF8 | ConvertFrom-Json
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE NETSCAPE-Bookmark-file-1>')
    [void]$sb.AppendLine('<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">')
    [void]$sb.AppendLine('<TITLE>Bookmarks</TITLE>')
    [void]$sb.AppendLine('<H1>Bookmarks</H1>')
    [void]$sb.AppendLine('<DL><p>')
    foreach ($rootNaam in 'bookmark_bar', 'other', 'synced') {
        $r = $data.roots.$rootNaam
        if (-not $r -or -not $r.children) { continue }
        $extra = if ($rootNaam -eq 'bookmark_bar') { ' PERSONAL_TOOLBAR_FOLDER="true"' } else { '' }
        [void]$sb.AppendLine("    <DT><H3$extra>$([Net.WebUtility]::HtmlEncode($r.name))</H3>")
        [void]$sb.AppendLine('    <DL><p>')
        foreach ($k in $r.children) { Add-BladwijzerKnoop $k $sb 2 }
        [void]$sb.AppendLine('    </DL><p>')
    }
    [void]$sb.AppendLine('</DL><p>')
    $html = $sb.ToString()
    [IO.File]::WriteAllText($Uit, $html, (New-Object System.Text.UTF8Encoding $false))
    return [regex]::Matches($html, '<DT><A ').Count
}

function Backup-Browsers {
    $run = Get-RunMap
    Write-Kop 'Browsers: bladwijzers, instellingen, extensies'
    $basis = Join-Path $run 'Browsers'
    $overzicht = New-Object System.Collections.Generic.List[object]

    foreach ($b in $script:Chromium) {
        if (-not (Test-Path -LiteralPath $b.Data)) { continue }
        [void](Confirm-Gesloten $b.Naam $b.Proces)
        $namen = Get-ChromiumProfielNamen $b.Data

        # Opera gebruikt de datamap zelf als profiel
        $profielen = if ($b.Naam -eq 'Opera') { @(Get-Item -LiteralPath $b.Data) } else {
            @(Get-ChildItem -LiteralPath $b.Data -Directory |
                Where-Object { ($_.Name -eq 'Default' -or $_.Name -like 'Profile *') -and (Test-Path -LiteralPath (Join-Path $_.FullName 'Preferences')) })
        }

        foreach ($p in $profielen) {
            $label = if ($namen[$p.Name]) { $namen[$p.Name] } else { $p.Name }
            $doelMap = Join-Path $basis "$($b.Naam)\$($p.Name)"
            New-Item -ItemType Directory -Force -Path $doelMap | Out-Null

            foreach ($f in $script:ChromiumBestanden) {
                $bron = Join-Path $p.FullName $f
                if (-not (Test-Path -LiteralPath $bron)) { continue }
                try { Copy-Item -LiteralPath $bron -Destination $doelMap -Force }
                catch { Write-Let "$($b.Naam) / ${label}: '$f' niet gekopieerd (in gebruik?)" }
            }

            $aantalBw = 0
            $bw = Join-Path $p.FullName 'Bookmarks'
            if (Test-Path -LiteralPath $bw) {
                try { $aantalBw = ConvertTo-BladwijzerHtml $bw (Join-Path $doelMap 'bladwijzers.html') }
                catch { Write-Let "$($b.Naam) / ${label}: bladwijzers niet om te zetten naar HTML ($($_.Exception.Message))" }
            }

            $ext = @(Get-ChromiumExtensies $p.FullName)
            if ($ext) { $ext | Export-Csv -LiteralPath (Join-Path $doelMap 'extensies.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';' }

            Write-Ok "$($b.Naam) / ${label}: $aantalBw bladwijzers, $($ext.Count) extensies"
            $overzicht.Add([pscustomobject]@{
                Browser     = $b.Naam
                Profiel     = $label
                Map         = $p.Name
                Bladwijzers = $aantalBw
                Extensies   = ($ext | ForEach-Object { $_.Naam }) -join ', '
            })
        }
    }

    # Firefox: het profiel is draagbaar, inclusief wachtwoorden (logins.json + key4.db)
    $ffRoot = Join-Path $env:APPDATA 'Mozilla\Firefox'
    $ffProfielen = Join-Path $ffRoot 'Profiles'
    if (Test-Path -LiteralPath $ffProfielen) {
        [void](Confirm-Gesloten 'Firefox' 'firefox')
        $ffDoel = Join-Path $basis 'Firefox'
        New-Item -ItemType Directory -Force -Path $ffDoel | Out-Null
        Copy-Item -LiteralPath (Join-Path $ffRoot 'profiles.ini') -Destination $ffDoel -ErrorAction SilentlyContinue

        foreach ($p in Get-ChildItem -LiteralPath $ffProfielen -Directory) {
            $doelMap = Join-Path $ffDoel $p.Name
            & robocopy $p.FullName $doelMap /E /R:1 /W:1 /XJD /NP /NFL /NDL /NJH /NJS `
                /XD cache2 crashes minidumps datareporting saved-telemetry-pings | Out-Null
            if ($LASTEXITCODE -ge 8) { Write-Let "Firefox / $($p.Name): niet alles gekopieerd (robocopy-code $LASTEXITCODE)" }

            $ext = @()
            $extJson = Join-Path $p.FullName 'extensions.json'
            if (Test-Path -LiteralPath $extJson) {
                try {
                    $ext = @((Get-Content -LiteralPath $extJson -Raw -Encoding UTF8 | ConvertFrom-Json).addons |
                        Where-Object { $_.location -eq 'app-profile' } |
                        ForEach-Object { $_.defaultLocale.name })
                } catch { }
            }
            $heeftWw = Test-Path -LiteralPath (Join-Path $p.FullName 'logins.json')
            Write-Ok "Firefox / $($p.Name): profiel gekopieerd, $($ext.Count) extensies, wachtwoorden: $(if ($heeftWw) { 'ja' } else { 'nee' })"
            $overzicht.Add([pscustomobject]@{
                Browser     = 'Firefox'
                Profiel     = $p.Name
                Map         = $p.Name
                Bladwijzers = '(in profiel)'
                Extensies   = $ext -join ', '
            })
        }
    }

    if ($overzicht.Count -eq 0) { Write-Let 'Geen ondersteunde browsers gevonden.' }
    $script:Rapport.Browsers = $overzicht.ToArray()
}

# ---------------------------------------------------------------------------
# Wachtwoorden
# ---------------------------------------------------------------------------

function Protect-Wachtwoorden([string]$Map) {
    $csvs = @(Get-ChildItem -LiteralPath $Map -Filter *.csv -File -ErrorAction SilentlyContinue)
    if (-not $csvs) { return }

    Write-Host "  Gevonden: $($csvs.Name -join ', ')"
    if ((Read-Host '  CSV-bestanden versleutelen met een wachtwoord? (J/n)') -match '^(n|nee|no)$') {
        Write-Let 'Niet versleuteld: de wachtwoorden staan in LEESBARE TEKST in de back-up. Bewaar de schijf goed en verwijder de CSV-bestanden na het importeren.'
        return
    }

    $7z = @((Get-Command 7z.exe -ErrorAction SilentlyContinue).Source,
            "$env:ProgramFiles\7-Zip\7z.exe",
            "${env:ProgramFiles(x86)}\7-Zip\7z.exe") |
        Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

    if (-not $7z) {
        Write-Let '7-Zip niet gevonden. De CSV-bestanden staan ONVERSLEUTELD in de back-up!'
        if ((Get-Command winget -ErrorAction SilentlyContinue) -and (Read-JaNee '7-Zip nu installeren via winget?')) {
            winget install --id 7zip.7zip -e --accept-source-agreements --accept-package-agreements
            $7z = "$env:ProgramFiles\7-Zip\7z.exe"
        }
        if (-not $7z -or -not (Test-Path -LiteralPath $7z)) { return }
    }

    while ($true) {
        $pw  = [Net.NetworkCredential]::new('', (Read-Host '  Wachtwoord voor het archief (min. 12 tekens)' -AsSecureString)).Password
        $pw2 = [Net.NetworkCredential]::new('', (Read-Host '  Herhaal het wachtwoord' -AsSecureString)).Password
        if ($pw.Length -lt 12) { Write-Let 'Te kort.'; continue }
        # Windows PowerShell 5.1 geeft " en een \ aan het eind verminkt door aan 7-Zip
        if ($pw -match '"|\\$') { Write-Let 'Gebruik geen " en laat het wachtwoord niet op \ eindigen.'; continue }
        if ($pw -ne $pw2) { Write-Let 'Komt niet overeen.'; continue }
        break
    }

    $archief = Join-Path $Map ('wachtwoorden_{0}.7z' -f (Get-Date -Format 'HHmmss'))
    & $7z a -t7z -mhe=on "-p$pw" $archief $csvs.FullName | Out-Null
    $ok = $LASTEXITCODE -eq 0
    if ($ok) { & $7z t "-p$pw" $archief | Out-Null; $ok = $LASTEXITCODE -eq 0 }
    $pw = $null; $pw2 = $null

    if ($ok) {
        $csvs | Remove-Item -Force
        Write-Ok "Versleuteld naar $(Split-Path $archief -Leaf); de CSV-bestanden zijn verwijderd."
        Write-Let 'Bewaar het archiefwachtwoord NIET bij de back-up (bijvoorbeeld op papier of in je wachtwoordmanager).'
    } else {
        Write-Fout '7-Zip is mislukt; de CSV-bestanden zijn NIET verwijderd.'
    }
}

# UI Automation-namen van de knoppen in de Chromium-wachtwoordpagina. Deze verschillen per
# browserversie en per taal; pas ze hier aan als het automatisch klikken niet aanslaat. Lukt het
# niet, dan valt Export-Wachtwoorden vanzelf terug op de handmatige stappen.
$script:UiaNamen = @{
    Meer         = '(?i)meer acties|more actions'
    Export       = '(?i)wachtwoorden exporteren|export passwords'
    Opslaan      = '(?i)^\s*(opslaan|save|bewaren)\s*$'
    Bestandsnaam = '(?i)bestandsnaam|file name'
}

# Zoekt binnen $Root naar het eerste zichtbare element van een van de gevraagde controletypen
# waarvan de naam op $Regex past. Blijft tot $TimeoutMs proberen, zodat trage pagina's meekomen.
function script:Uia-Find {
    param($Root, [System.Windows.Automation.ControlType[]]$Types, [string]$Regex, [int]$TimeoutMs = 8000)
    $conds = foreach ($t in $Types) {
        New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $t)
    }
    $cond = if (@($conds).Count -eq 1) { $conds } else { New-Object System.Windows.Automation.OrCondition($conds) }
    $deadline = [DateTime]::Now.AddMilliseconds($TimeoutMs)
    do {
        try {
            foreach ($e in $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)) {
                if ($e.Current.Name -match $Regex) { return $e }
            }
        } catch {}
        Start-Sleep -Milliseconds 350
    } while ([DateTime]::Now -lt $deadline)
    return $null
}

# Klikt/opent een element: eerst via Invoke (knoppen), anders via Expand (uitklapmenu's).
function script:Uia-Klik($E) {
    if (-not $E) { return $false }
    try { $E.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); return $true } catch {}
    try { $E.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand(); return $true } catch {}
    return $false
}

# Alle open browservensters (Chromium deelt deze vensterklasse).
function script:Uia-Vensters {
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty, 'Chrome_WidgetWin_1')
    return @($desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $c))
}

# Kiest de browservensters waarin gezocht mag worden, zodat een ander open Chromium-venster niet
# per ongeluk aangeklikt wordt. Eerst het venster van het proces dat we net startten; anders het
# voorgrondvenster (als dat een Chromium-venster is); als laatste redmiddel alle Chromium-vensters.
function script:Uia-DoelVensters($Proc) {
    $alle = Uia-Vensters
    try {
        if ($Proc -and -not $Proc.HasExited) {
            $Proc.Refresh()
            $h = $Proc.MainWindowHandle
            if ($h -ne [IntPtr]::Zero) {
                $el = [System.Windows.Automation.AutomationElement]::FromHandle($h)
                if ($el) { return @($el) }
            }
        }
    } catch {}
    try {
        if (-not ('Uia.Win' -as [type])) {
            Add-Type -Namespace Uia -Name Win -MemberDefinition `
                '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern System.IntPtr GetForegroundWindow();'
        }
        $fg = [Uia.Win]::GetForegroundWindow()
        foreach ($w in $alle) {
            if ([IntPtr]$w.Current.NativeWindowHandle -eq $fg) { return @($w) }
        }
    } catch {}
    return $alle
}

# Wacht tot een dialoogvenster met de gevraagde vensterklasse verschijnt (bijv. het opslagvenster).
function script:Uia-Dialoog([string]$ClassName, [int]$TimeoutMs) {
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty, $ClassName)
    $deadline = [DateTime]::Now.AddMilliseconds($TimeoutMs)
    do {
        try { $d = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $c); if ($d) { return $d } } catch {}
        Start-Sleep -Milliseconds 400
    } while ([DateTime]::Now -lt $deadline)
    return $null
}

# Stuurt via UI Automation de wachtwoord-export van een Chromium-browser aan. De gebruiker hoeft
# alleen nog zijn pincode in te vullen bij de Windows-beveiligingsprompt. Geeft $true terug als de
# CSV er daadwerkelijk staat; $false betekent: terugvallen op de handmatige stappen.
# Welke CSV-bestanden er sinds $Bestaand in $Map bij zijn gekomen. Zo maakt de standaard-
# bestandsnaam van de browser niet uit (Edge slaat in het NL op als 'Microsoft Edge-wachtwoorden.csv').
function script:Get-NieuweCsv([string]$Map, [string[]]$Bestaand) {
    @(Get-ChildItem -LiteralPath $Map -Filter *.csv -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notin $Bestaand })
}

function Invoke-ChromiumExport {
    param([hashtable]$B, [string]$Csv, $Proc, [string]$Map, [string[]]$Bestaand)
    try { Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes -ErrorAction Stop } catch { return $false }
    $btnMi = @([System.Windows.Automation.ControlType]::Button, [System.Windows.Automation.ControlType]::MenuItem)
    $doel = Uia-DoelVensters $Proc

    # 1. Het 'Exporteren'-menu-item vinden (eventueel eerst het 'meer acties'-menu uitklappen).
    $export = $null
    foreach ($w in $doel) { $export = Uia-Find $w $btnMi $script:UiaNamen.Export 3000; if ($export) { break } }
    if (-not $export) {
        foreach ($w in $doel) {
            $meer = Uia-Find $w $btnMi $script:UiaNamen.Meer 2500
            if ($meer) { [void](Uia-Klik $meer); Start-Sleep -Milliseconds 600; break }
        }
        foreach ($w in $doel) { $export = Uia-Find $w $btnMi $script:UiaNamen.Export 3000; if ($export) { break } }
    }
    if (-not $export) { return $false }
    [void](Uia-Klik $export)
    Start-Sleep -Milliseconds 800

    # 2. De bevestigingsknop 'Exporteren' in de dialoog (zelfde naam) aanklikken.
    foreach ($w in $doel) {
        $bevestig = Uia-Find $w @([System.Windows.Automation.ControlType]::Button) $script:UiaNamen.Export 3000
        if ($bevestig) { [void](Uia-Klik $bevestig); break }
    }

    # 3. Windows Hello: hier vult de gebruiker zelf zijn pincode of wachtwoord in.
    Write-Host '  >> Vul nu je pincode of wachtwoord in bij de Windows-beveiligingsprompt.' -ForegroundColor Yellow

    # 4. Wachten op het opslagvenster en het volledige pad invullen, daarna opslaan.
    $dlg = Uia-Dialoog '#32770' 120000
    if (-not $dlg) { return ((Get-NieuweCsv $Map $Bestaand).Count -gt 0) }
    $edit = Uia-Find $dlg @([System.Windows.Automation.ControlType]::Edit) $script:UiaNamen.Bestandsnaam 5000
    if (-not $edit) { $edit = Uia-Find $dlg @([System.Windows.Automation.ControlType]::Edit) '.' 3000 }
    if ($edit) {
        try { $edit.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Csv) } catch {}
    }
    Start-Sleep -Milliseconds 300
    $opslaan = Uia-Find $dlg @([System.Windows.Automation.ControlType]::Button) $script:UiaNamen.Opslaan 4000
    if ($opslaan) { [void](Uia-Klik $opslaan) }

    # 5. Wachten tot er een nieuw CSV-bestand verschijnt (welke naam de browser ook kiest).
    $deadline = [DateTime]::Now.AddSeconds(20)
    while ([DateTime]::Now -lt $deadline -and (Get-NieuweCsv $Map $Bestaand).Count -eq 0) { Start-Sleep -Milliseconds 400 }
    return ((Get-NieuweCsv $Map $Bestaand).Count -gt 0)
}

function Export-Wachtwoorden {
    $run = Get-RunMap
    $map = Join-Path $run 'Wachtwoorden'
    New-Item -ItemType Directory -Force -Path $map | Out-Null
    Write-Kop 'Wachtwoorden uit browsers'
    Write-Host $script:WachtwoordUitleg

    foreach ($b in $script:Chromium) {
        if (-not (Test-Path -LiteralPath $b.Data)) { continue }
        if (-not (Read-JaNee "Wachtwoorden exporteren uit $($b.Naam)?")) { continue }

        $csv = Join-Path $map "$($b.Naam)-wachtwoorden.csv"
        # Momentopname vooraf: alles wat hierna aan CSV's verschijnt, is deze export (ongeacht de naam
        # die de browser kiest, zoals 'Microsoft Edge-wachtwoorden.csv' in het Nederlands).
        $bestaand = @(Get-ChildItem -LiteralPath $map -Filter *.csv -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })

        $proc = $null
        try { $proc = Start-Process -FilePath $b.Proces -ArgumentList $b.Ww -PassThru }
        catch { Write-Let "Kon $($b.Naam) niet starten; open hem zelf."; Set-Clipboard -Value $b.Ww }
        Write-Host "  $($b.Naam) wordt geopend. Ik probeer het exporteren automatisch te starten..."
        Start-Sleep -Seconds 4

        $gelukt = $false
        try { $gelukt = Invoke-ChromiumExport -B $b -Csv $csv -Proc $proc -Map $map -Bestaand $bestaand } catch { $gelukt = $false }

        if (-not $gelukt -and (Get-NieuweCsv $map $bestaand).Count -eq 0) {
            # Automatisch klikken sloeg niet aan (andere versie/taal): terug naar de handmatige stappen.
            Write-Let 'Automatisch starten lukte niet; doe de export even met de hand:'
            Set-Clipboard -Value $b.Ww
            Write-Host "  1. Het adres $($b.Ww) staat op je klembord. Plak het in de adresbalk als de pagina niet vanzelf opent."
            Write-Host "  2. Kies 'Wachtwoorden exporteren/downloaden' en bevestig met je Windows-pincode of -wachtwoord."
            [void](Read-Host '  3. Druk op Enter zodra het venster om op te slaan open is')
            Write-Host "  4. Sla het op in deze map (de standaard bestandsnaam van de browser is prima):"
            Write-Host "     $map" -ForegroundColor White
            Set-Clipboard -Value $map
            Write-Host '     Het mappad staat op je klembord; plak het bovenin de adresbalk van het opslagvenster.'
            [void](Read-Host '  5. Druk op Enter als het bestand is opgeslagen')
        }

        $nieuw = Get-NieuweCsv $map $bestaand
        if ($nieuw.Count -gt 0) { Write-Ok "$($b.Naam): opgeslagen ($($nieuw.Name -join ', '))" }
        else { Write-Let "$($b.Naam): geen nieuw CSV-bestand gevonden in $map. Heb je het ergens anders opgeslagen? Zet het dan zelf in deze map." }
    }
    Set-Clipboard -Value ' '

    if (Test-Path -LiteralPath (Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles')) {
        Write-Host '  Firefox: de wachtwoorden gaan mee met de profielkopie (stap Browsers). Heb je een hoofdwachtwoord ingesteld? Dan heb je dat op de nieuwe laptop weer nodig.'
    }
    Protect-Wachtwoorden $map
}

# ---------------------------------------------------------------------------
# Programma-instellingen (AppData en register)
# ---------------------------------------------------------------------------

# Databasemap van eM Client: een eigen locatie (Instellingen > Algemeen > Opslag) staat in het
# register; de Microsoft Store-versie leest en schrijft in zijn eigen pakketmap
function Get-EmClientMap {
    $eigen = (Get-ItemProperty -LiteralPath 'HKCU:\SOFTWARE\eM Client' -ErrorAction SilentlyContinue).DBLocation
    $kandidaten = @($eigen, (Join-Path $env:APPDATA 'eM Client')) +
        @(Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Packages\*eM*Client*\LocalCache\Roaming\eM Client') `
            -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    return $kandidaten | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
}

function Backup-Instellingen {
    $run = Get-RunMap
    Write-Kop 'Programma-instellingen'
    $basis = Join-Path $run 'Instellingen'

    # Eerst kijken wat er op deze pc staat
    $gevonden = foreach ($d in $script:Instellingen) {
        if ($d.Soort -eq 'Register') {
            if (-not (Test-Path -LiteralPath ($d.Pad -replace '^HKCU\\', 'HKCU:\'))) { continue }
            $grootte = '-'
            $bytes = 0L
        } else {
            $pad = if ($d.Zoek) { & $d.Zoek } else { [Environment]::ExpandEnvironmentVariables($d.Pad) }
            if (-not $pad -or -not (Test-Path -LiteralPath $pad)) { continue }
            if ($d.Filter) {
                $bytes = [long]((Get-ChildItem -LiteralPath $pad -Filter $d.Filter -File -ErrorAction SilentlyContinue |
                    Measure-Object Length -Sum).Sum)
            } else {
                $bytes = (Measure-Map $pad).Bytes
            }
            if ($bytes -eq 0) { continue }
            $grootte = Format-Grootte $bytes
        }
        $bronPad = if ($d.Soort -eq 'Register') { $d.Pad } else { $pad }
        [pscustomobject]@{ Naam = $d.Naam; Grootte = $grootte; Bytes = $bytes; Pad = $bronPad; Definitie = $d }
    }
    $gevonden = @($gevonden)
    if (-not $gevonden) { Write-Let 'Geen bekende programma-instellingen gevonden.'; return }

    $keuze = @(Select-Items -Items $gevonden -Titel 'Welke instellingen meenemen?' `
        -Kolommen @('Naam', 'Grootte', 'Pad') -Voorstel { $true })
    if (-not $keuze) { Write-Let 'Niets gekozen.'; return }

    $log = Join-Path $run 'robocopy.log'
    $manifest = New-Object System.Collections.Generic.List[object]
    foreach ($k in $keuze) {
        $d = $k.Definitie
        $doelMap = Join-Path $basis $d.Naam
        if (-not (Confirm-Gesloten $d.Naam $d.Proces)) {
            # Een database die tijdens het kopiëren verandert, is op de nieuwe pc mogelijk kapot
            if ($d.MoetDicht) { Write-Fout "$($d.Naam): overgeslagen, het programma draait nog."; continue }
            Write-Let "$($d.Naam): wordt toch gekopieerd; vergrendelde bestanden ontbreken misschien."
        }

        $ok = $true
        switch ($d.Soort) {
            'Register' {
                New-Item -ItemType Directory -Force -Path $doelMap | Out-Null
                & reg.exe export $d.Pad (Join-Path $doelMap 'instellingen.reg') /y | Out-Null
                $ok = $LASTEXITCODE -eq 0
            }
            default {
                $bron = $k.Pad
                $extra = @(if ($d.Filter) { $d.Filter } else { '/E' })
                if ($d.Uitsluiten) { $extra += @('/XD') + $d.Uitsluiten }
                $ok = (Invoke-Robocopy $bron $doelMap $extra $log) -lt 8
            }
        }

        # VS Code: extensies staan niet in de User-map, dus de lijst apart bewaren
        if ($d.Naam -eq 'VS Code' -and (Get-Command code -ErrorAction SilentlyContinue)) {
            $ext = @(& code --list-extensions)
            $ext | Set-Content -LiteralPath (Join-Path $basis 'vscode-extensies.txt') -Encoding UTF8
            Write-Ok "VS Code: $($ext.Count) extensies -> vscode-extensies.txt"
        }

        if ($ok) { Write-Ok "$($d.Naam): $($k.Grootte)" } else { Write-Fout "$($d.Naam): niet (volledig) gekopieerd" }
        $manifest.Add([pscustomobject]@{
            Naam   = $d.Naam
            Soort  = $d.Soort
            Pad    = $d.Pad
            Filter = $d.Filter
            Proces = $d.Proces
            Vervangen = [bool]$d.Vervangen
            Status = if ($ok) { 'OK' } else { 'Fouten' }
        })
    }

    if ($manifest.Count -eq 0) { Write-Let 'Er is niets gekopieerd.'; return }
    New-Item -ItemType Directory -Force -Path $basis | Out-Null
    # -InputObject: anders wordt een lijst met één item in 5.1 geen JSON-array
    ConvertTo-Json -InputObject @($manifest.ToArray()) -Depth 3 |
        Set-Content -LiteralPath (Join-Path $basis 'instellingen.json') -Encoding UTF8
    $script:Rapport.Instellingen = @($manifest | Select-Object Naam, Pad, Status)

    if ($keuze | Where-Object { $_.Naam -eq 'FileZilla' }) {
        Write-Let 'FileZilla bewaart wachtwoorden van sites vrijwel leesbaar (base64) in sitemanager.xml. Bewaar de back-up veilig.'
    }
    if ($manifest | Where-Object { $_.Naam -eq 'eM Client' -and $_.Status -eq 'OK' }) {
        Write-Let 'eM Client: je mail en de opgeslagen accountwachtwoorden staan nu in de back-up. Bewaar die veilig (bijv. een BitLocker-schijf).'
        Write-Host '       Op de nieuwe pc: eerst het herstel doen, daarna pas eM Client starten. Activeer daar je licentie opnieuw (Menu > Help > Licentie).'
    }
}

# ---------------------------------------------------------------------------
# Home-map
# ---------------------------------------------------------------------------

# Het echte pad van een vaste map (Desktop, MyDocuments, ...), ook als die omgeleid is
function Get-BekendePad([string]$Naam) { return [Environment]::GetFolderPath($Naam) }

function Format-Online([long]$Bytes) { if ($Bytes -gt 0) { "$(Format-Grootte $Bytes) alleen online" } else { '' } }

function Invoke-HomeScan {
    Write-Kop "Home-map scannen: $env:USERPROFILE"
    $cloud = @($env:OneDrive, $env:OneDriveConsumer, $env:OneDriveCommercial) | Where-Object { $_ }
    $resultaat = New-Object System.Collections.Generic.List[object]
    $grootste = New-Object System.Collections.Generic.List[object]

    $mappen = @(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { -not $_.LinkType })   # junctions zoals 'Application Data' niet; OneDrive (ook een reparse point) wel

    $i = 0
    foreach ($d in $mappen) {
        Write-Progress -Activity 'Home-map scannen' -Status $d.Name -PercentComplete ($i / $mappen.Count * 100)
        $i++
        $meting = Measure-Map $d.FullName
        foreach ($g in $meting.Groot) { $grootste.Add($g) }

        $type = 'Gebruiker'
        $advies = 'Kopiëren'
        if ($cloud -contains $d.FullName -or $d.Name -like 'OneDrive*') {
            $type = 'Cloud'; $advies = 'Overslaan (staat al in OneDrive; kopiëren downloadt alles)'
        } elseif ($script:HomeOverslaan -contains $d.Name) {
            $type = 'Systeem'; $advies = 'Overslaan (programma-instellingen)'
        } elseif ($meting.Bestanden -eq 0) {
            $advies = 'Overslaan (leeg)'
        }
        $resultaat.Add([pscustomobject]@{
            Map          = $d.Name
            GrootteTekst = Format-Grootte $meting.Bytes
            Bestanden    = $meting.Bestanden
            Online       = Format-Online $meting.OnlineBytes
            Advies       = $advies
            Type         = $type
            Bytes        = $meting.Bytes
            OnlineBytes  = $meting.OnlineBytes
            Pad          = $d.FullName
        })
    }
    Write-Progress -Activity 'Home-map scannen' -Completed

    # Vaste mappen die omgeleid zijn, meestal naar OneDrive door 'OneDrive-back-up'. Die staan dan niet
    # in de home-map maar horen er wel bij. Ze krijgen de standaardnaam (Desktop, Documents, ...), zodat
    # het herstellen ze op de nieuwe pc naar het bureaublad of Documenten terugzet.
    foreach ($naam in 'Desktop', 'Documents', 'Pictures', 'Music', 'Videos') {
        $pad = Get-BekendePad $script:BekendeMappen[$naam]
        if (-not $pad -or $pad -eq (Join-Path $env:USERPROFILE $naam) -or -not (Test-Path -LiteralPath $pad)) { continue }

        $meting = Measure-Map $pad
        foreach ($g in $meting.Groot) { $grootste.Add($g) }
        Write-Ok "$naam staat in $pad; neem ik mee als '$naam'"

        # Wat er nog in de home-map onder dezelfde naam staat, is meestal een lege restmap
        $rest = $resultaat | Where-Object { $_.Map -eq $naam } | Select-Object -First 1
        if ($rest) {
            if ($rest.Bestanden -eq 0) { [void]$resultaat.Remove($rest) } else { $rest.Map = "$naam (home-map)" }
        }
        $resultaat.Add([pscustomobject]@{
            Map          = $naam
            GrootteTekst = Format-Grootte $meting.Bytes
            Bestanden    = $meting.Bestanden
            Online       = Format-Online $meting.OnlineBytes
            Advies       = if ($meting.Bestanden) { 'Kopiëren' } else { 'Overslaan (leeg)' }
            Type         = 'Omgeleid'
            Bytes        = $meting.Bytes
            OnlineBytes  = $meting.OnlineBytes
            Pad          = $pad
        })
    }

    # Losse bestanden direct in de home-map (bijv. .gitconfig), zonder het gebruikersregister
    $los = @(Get-ChildItem -LiteralPath $env:USERPROFILE -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike 'ntuser*' -and $_.Name -notlike 'UsrClass*' })
    if ($los) {
        $som = ($los | Measure-Object Length -Sum).Sum
        $resultaat.Add([pscustomobject]@{
            Map          = '(losse bestanden)'
            GrootteTekst = Format-Grootte $som
            Bestanden    = $los.Count
            Advies       = 'Kopiëren'
            Type         = 'Losse bestanden'
            Bytes        = [long]$som
            OnlineBytes  = 0L
            Pad          = $env:USERPROFILE
        })
    }

    $script:Scan = @($resultaat | Sort-Object Bytes -Descending)
    $script:Rapport.Home = $script:Scan
    $script:Rapport.Grootste = @($grootste | Sort-Object Bytes -Descending | Select-Object -First 20)

    $script:Scan | Format-Table Map, GrootteTekst, Bestanden, Online, Advies -AutoSize | Out-Host
    $totaal = ($script:Scan | Measure-Object Bytes -Sum).Sum
    $voorstel = ($script:Scan | Where-Object { $_.Advies -eq 'Kopiëren' } | Measure-Object Bytes -Sum).Sum
    Write-Host "  Totaal: $(Format-Grootte $totaal), waarvan voorgesteld om te kopiëren: $(Format-Grootte $voorstel)"

    if ($script:Rapport.Grootste) {
        Write-Host ''
        Write-Host '  Grootste bestanden (>= 100 MB):'
        $script:Rapport.Grootste | Select-Object -First 10 | Format-Table Grootte, Pad -AutoSize | Out-Host
    }
}

function Copy-HomeMappen {
    $run = Get-RunMap
    if (-not $script:Scan) { Invoke-HomeScan }
    Write-Kop 'Home-map kopiëren'

    $keuze = @(Select-Items -Items $script:Scan -Titel 'Welke mappen kopiëren?' `
        -Kolommen @('Map', 'GrootteTekst', 'Bestanden', 'Online', 'Advies') -Voorstel { $_.Advies -eq 'Kopiëren' })
    if (-not $keuze) { Write-Let 'Niets gekozen.'; return }

    if ($keuze | Where-Object { $_.Type -eq 'Cloud' }) {
        Write-Let 'Je hebt OneDrive gekozen: bestanden die alleen in de cloud staan worden nu eerst gedownload.'
    }
    $online = [long](($keuze | Measure-Object OnlineBytes -Sum).Sum)
    if ($online -gt 0) {
        Write-Let "$(Format-Grootte $online) staat alleen online (OneDrive) en wordt nu eerst gedownload; dat kan even duren."
    }

    $nodig = ($keuze | Measure-Object Bytes -Sum).Sum
    Write-Host "  Te kopiëren: $(Format-Grootte $nodig) in $($keuze.Count) map(pen)"
    try {
        $schijf = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($run))
        if ($schijf.AvailableFreeSpace -lt $nodig * 1.05) {
            Write-Fout "Te weinig ruimte op $($schijf.Name): $(Format-Grootte $schijf.AvailableFreeSpace) vrij."
            if (-not (Read-JaNee 'Toch beginnen?')) { return }
        }
    } catch {
        Write-Let 'De vrije ruimte op een netwerklocatie is niet te bepalen; controleer zelf of er genoeg is.'
    }

    $log = Join-Path $run 'robocopy.log'
    $doelBasis = Join-Path $run 'Home'
    $i = 0
    foreach ($m in $keuze) {
        Write-Progress -Activity 'Home-map kopiëren' -Status "$($m.Map) ($($m.GrootteTekst))" -PercentComplete ($i / $keuze.Count * 100)
        $i++

        if ($m.Type -eq 'Losse bestanden') {
            $dst = Join-Path $doelBasis '_losse_bestanden'
            $rcArgs = @($m.Pad, $dst, '/XF', 'ntuser*', 'UsrClass*')   # zonder /E: alleen de bovenste laag
        } else {
            $dst = Join-Path $doelBasis $m.Map
            $rcArgs = @($m.Pad, $dst, '/E')
        }
        # /XJD i.p.v. /XJ: alleen gekoppelde mappen overslaan; OneDrive-bestanden zijn ook reparse points
        $rcArgs += @('/COPY:DAT', '/DCOPY:T', '/R:1', '/W:1', '/XJD', '/MT:8', '/NP', '/NDL', "/LOG+:$log")
        & robocopy @rcArgs | Out-Null
        $code = $LASTEXITCODE

        # Controle: aantal bestanden en bytes op de bestemming
        if ($m.Type -eq 'Losse bestanden') {
            $los = @(Get-ChildItem -LiteralPath $dst -File -Force -ErrorAction SilentlyContinue)
            $controle = [pscustomobject]@{ Bestanden = $los.Count; Bytes = [long](($los | Measure-Object Length -Sum).Sum) }
        } else {
            $controle = Measure-Map $dst
        }

        $status = if ($code -ge 8) { "Fouten (robocopy $code)" }
                  elseif ($controle.Bestanden -lt $m.Bestanden) { 'Onvolledig' }
                  else { 'OK' }
        if ($status -eq 'OK') { Write-Ok "$($m.Map): $($controle.Bestanden) bestanden, $(Format-Grootte $controle.Bytes)" }
        else { Write-Fout "$($m.Map): $status - bron $($m.Bestanden) / doel $($controle.Bestanden) bestanden. Zie robocopy.log." }

        $script:Rapport.Kopieen += [pscustomobject]@{
            Map           = $m.Map
            Status        = $status
            BronBestanden = $m.Bestanden
            DoelBestanden = $controle.Bestanden
            BronGrootte   = Format-Grootte $m.Bytes
            DoelGrootte   = Format-Grootte $controle.Bytes
        }
    }
    Write-Progress -Activity 'Home-map kopiëren' -Completed
}

# ---------------------------------------------------------------------------
# Rapport
# ---------------------------------------------------------------------------

function Save-Rapport {
    if (-not $script:Run) { return }
    $R = $script:Rapport
    try { $R.Windows = (Get-CimInstance Win32_OperatingSystem).Caption } catch { }

    $R | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $script:Run 'rapport.json') -Encoding UTF8

    $css = '<style>body{font-family:"Segoe UI",sans-serif;margin:2em;color:#222}table{border-collapse:collapse;margin-bottom:1em}' +
           'td,th{border:1px solid #ccc;padding:4px 8px;text-align:left;vertical-align:top}th{background:#eee}h2{margin-top:1.5em}</style>'
    $delen = @("<h1>Migratierapport $($R.Computer)</h1><p>$($R.Gebruiker) &middot; $($R.Windows) &middot; $($R.Datum)</p>")
    if ($R.Kopieen)    { $delen += $R.Kopieen | ConvertTo-Html -Fragment -PreContent '<h2>Gekopieerd</h2>' }
    if ($R.Home)       { $delen += $R.Home | Select-Object Map, GrootteTekst, Bestanden, Online, Advies, Pad | ConvertTo-Html -Fragment -PreContent '<h2>Home-map</h2>' }
    if ($R.Grootste)   { $delen += $R.Grootste | Select-Object Grootte, Pad | ConvertTo-Html -Fragment -PreContent '<h2>Grootste bestanden</h2>' }
    if ($R.Browsers)   { $delen += $R.Browsers | ConvertTo-Html -Fragment -PreContent '<h2>Browsers</h2>' }
    if ($R.Instellingen) { $delen += $R.Instellingen | ConvertTo-Html -Fragment -PreContent '<h2>Programma-instellingen</h2>' }
    if ($R.Linux)      { $delen += $R.Linux | ConvertTo-Html -Fragment -PreContent "<h2>Programma's op Linux</h2>" }
    if ($R.Programmas) { $delen += $R.Programmas | ConvertTo-Html -Fragment -PreContent "<h2>Programma's ($($R.Programmas.Count))</h2>" }

    ConvertTo-Html -Head $css -Body ($delen -join "`n") -Title "Migratie $($R.Computer)" |
        Set-Content -LiteralPath (Join-Path $script:Run 'rapport.html') -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Naar Linux
# ---------------------------------------------------------------------------

$script:LinuxUitleg = @'
NIEUWE COMPUTER MET LINUX
=========================

1. Sluit deze schijf aan op de Linux-pc (NTFS en exFAT kan Linux gewoon lezen) of koppel de NAS-share.
2. Open een terminal in deze map en start:
       bash Linux/herstel-linux.sh
   Of geef de map mee:  bash herstel-linux.sh /pad/naar/deze/map
   Niet met sudo; het script vraagt zelf om je wachtwoord als het iets moet installeren.
3. Het script:
   - installeert de Linux-versies van je programma's via Flathub (Linux/apps.tsv);
   - zet instellingen terug: VS Code, Thunderbird, FileZilla, lettertypen; Outlook-PST wordt mbox;
   - zet bladwijzers van Chrome/Edge/Brave/Vivaldi/Opera en het hele Firefox-profiel terug;
   - voegt je PuTTY/WinSCP-sessies toe aan ~/.ssh/config (Linux/ssh-config.txt);
   - kopieert je bestanden naar de Linux-mappen (Documents -> Documenten enz.);
   - pakt de wachtwoorden uit in het geheugen, laat je ze importeren en ruimt ze daarna op.

Linux/programmas-linux.csv: per Windows-programma de Linux-vervanger of een advies.
Linux/zelf-zoeken.txt: programma's waarvoor geen vaste vervanger bekend is.
'@

# Het herstelscript voor Linux. Wordt met LF-regeleinden en zonder BOM weggeschreven.
$script:LinuxHerstelScript = @'
#!/usr/bin/env bash
# Migratie-assistent: herstellen op Linux.
# Gemaakt door Migratie-Assistent.ps1 op de oude Windows-pc. Start het vanuit de back-upmap:
#     bash Linux/herstel-linux.sh
# of geef de back-upmap mee:  bash herstel-linux.sh /media/usb/Migratie_OUDEPC_20261014_0930
# Werkt op Ubuntu/Mint/Debian, Fedora, openSUSE en Arch. Programma's komen van Flathub (Flatpak),
# zodat het niet uitmaakt welke distributie je gebruikt.
set -u
shopt -s dotglob    # ook verborgen mappen zoals Home/.ssh

HIER="$(cd -- "$(dirname "${BASH_SOURCE[0]}")" >/dev/null && pwd)"
BACKUP=''    # wordt bij de start bepaald, zie zoek_backup
LINUX=''
LOGMAP="${XDG_STATE_HOME:-$HOME/.local/state}/migratie-assistent"
mkdir -p "$LOGMAP"
LOG="$LOGMAP/herstel_$(date +%Y%m%d_%H%M).log"
KEUZE=()

kop()     { printf '\n\e[36m== %s ==\e[0m\n' "$1"; echo "== $1 ==" >>"$LOG"; }
ok()      { printf '  \e[32m[OK]\e[0m %s\n' "$1"; echo "[OK] $1" >>"$LOG"; }
waarsch() { printf '  \e[33m[!]\e[0m  %s\n' "$1"; echo "[!] $1" >>"$LOG"; }
fout()    { printf '  \e[31m[X]\e[0m  %s\n' "$1"; echo "[X] $1" >>"$LOG"; }
ja_nee()  { local a; read -r -p "  $1 (j/N) " a; [[ "$a" =~ ^([jJyY]|ja|Ja|yes)$ ]]; }

# kies TITEL TAG OMSCHRIJVING ON|OFF ...  ->  gekozen tags in KEUZE
# Met whiptail een lijst met vinkjes, anders een genummerde lijst (zoals in de Windows-versie)
kies() {
    local titel="$1"; shift
    KEUZE=()
    local tags=() oms=() aan=()
    while (($# >= 3)); do tags+=("$1"); oms+=("$2"); aan+=("$3"); shift 3; done
    ((${#tags[@]})) || return 0
    local i
    if command -v whiptail >/dev/null && [[ -t 0 && -t 1 ]]; then
        local args=() uit
        for i in "${!tags[@]}"; do args+=("${tags[$i]}" "${oms[$i]}" "${aan[$i]}"); done
        uit=$(whiptail --title 'Migratie-assistent' --separate-output --checklist \
            "$titel (spatie = aan/uit, Enter = verder)" 22 100 14 "${args[@]}" 3>&1 1>&2 2>&3) || return 0
        [[ -n "$uit" ]] && mapfile -t KEUZE <<<"$uit"
        return 0
    fi
    for i in "${!tags[@]}"; do
        printf '  %3d. [%s] %-35s %s\n' $((i + 1)) "$([[ ${aan[$i]} == ON ]] && echo x || echo ' ')" "${tags[$i]}" "${oms[$i]}"
    done
    local antw deel van tot n
    read -r -p "  $titel - nummers (bijv. 1,3,5-7), 'a' = alles, 'v' = aangevinkte, leeg = niets: " antw
    case "$antw" in
        '') return 0 ;;
        a) KEUZE=("${tags[@]}"); return 0 ;;
        v) for i in "${!tags[@]}"; do [[ ${aan[$i]} == ON ]] && KEUZE+=("${tags[$i]}"); done; return 0 ;;
    esac
    IFS=',' read -ra delen <<<"$antw"
    for deel in "${delen[@]}"; do
        deel="${deel// /}"
        if [[ "$deel" =~ ^([0-9]+)-([0-9]+)$ ]]; then van=${BASH_REMATCH[1]}; tot=${BASH_REMATCH[2]}
        elif [[ "$deel" =~ ^[0-9]+$ ]]; then van=$deel; tot=$deel
        else continue; fi
        for ((n = van; n <= tot; n++)); do
            ((n >= 1 && n <= ${#tags[@]})) && KEUZE+=("${tags[$((n - 1))]}")
        done
    done
    return 0
}

# Installeert een pakket met de pakketbeheerder van deze distributie; probeert de namen op volgorde
pakket() {
    local p
    for p in "$@"; do
        if command -v apt-get >/dev/null; then sudo apt-get install -y "$p" && return 0
        elif command -v dnf >/dev/null; then sudo dnf install -y "$p" && return 0
        elif command -v zypper >/dev/null; then sudo zypper --non-interactive install "$p" && return 0
        elif command -v pacman >/dev/null; then sudo pacman -S --needed --noconfirm "$p" && return 0
        else fout "Geen bekende pakketbeheerder; installeer zelf: $*"; return 1
        fi
    done
    return 1
}

# Een programma dat nog draait overschrijft bij het afsluiten wat we terugzetten
wacht_gesloten() {  # NAAM PROCES-REGEX
    while pgrep -x "$2" >/dev/null 2>&1; do
        waarsch "$1 draait nog; sluit het eerst af."
        ja_nee 'Gesloten? Opnieuw controleren (nee = overslaan)' || return 1
    done
    return 0
}

# kopieer BRON DOEL [rsync-opties]: overschrijft. Met --update blijven nieuwere bestanden in DOEL staan.
kopieer() {
    local bron="$1" doel="$2"; shift 2
    mkdir -p "$doel" || return 1
    if command -v rsync >/dev/null; then
        rsync -rt --chmod=u+rwX --exclude desktop.ini --exclude Thumbs.db --log-file="$LOG" "$@" "$bron/" "$doel/"
        return
    fi
    # Geen rsync (bijv. Arch). De rechten van de bron niet overnemen: een alleen-lezen back-upschijf
    # zou anders alleen-lezen mappen in je home-map opleveren.
    if [[ " $* " == *' --update '* ]]; then
        cp -ru --preserve=timestamps --no-preserve=mode "$bron/." "$doel/" 2>>"$LOG"
        return
    fi
    local opties=(--exclude=desktop.ini --exclude=Thumbs.db --mode=u+rwX)
    while (($#)); do
        [[ "$1" == --exclude ]] && { opties+=("--exclude=$2"); shift; }
        shift
    done
    tar -C "$bron" "${opties[@]}" -cf - . | tar -C "$doel" -xf - 2>>"$LOG"
}

# Waar een map uit Home/ hoort. Linux kent dezelfde vaste mappen, vaak met een Nederlandse naam.
xdg_doel() {
    local k p=''
    case "$1" in
        Desktop | Bureaublad) k=DESKTOP ;; Documents | Documenten) k=DOCUMENTS ;; Downloads) k=DOWNLOAD ;;
        Music | Muziek) k=MUSIC ;; Pictures | Afbeeldingen) k=PICTURES ;; Videos | "Video's") k=VIDEOS ;;
        _losse_bestanden) echo "$HOME"; return ;;
        *) echo "$HOME/$1"; return ;;
    esac
    command -v xdg-user-dir >/dev/null && p="$(xdg-user-dir "$k")"
    if [[ -n "$p" && "$p" != "$HOME" ]]; then echo "$p"; else echo "$HOME/$1"; fi
}

# Pad van het standaardprofiel uit een Firefox/Thunderbird profiles.ini
mozilla_standaard() {
    tr -d '\r' <"$1" | awk -F= '
        /^\[/ { sectie = $0; next }
        sectie ~ /^\[Install/ && $1 == "Default" && !inst { inst = $2 }
        sectie ~ /^\[Profile/ && $1 == "Path" { pad[sectie] = $2; if (!eerste) eerste = sectie }
        sectie ~ /^\[Profile/ && $1 == "Default" && $2 == "1" { std = sectie }
        END { if (inst) print inst; else if (std) print pad[std]; else if (eerste) print pad[eerste] }'
}

# mozilla_herstel NAAM PROCES BACKUPMAP OUDE_PROFILES_INI MOGELIJKE_MAPPEN...
# Zet een profiel uit de back-up over het (nieuwe) standaardprofiel van deze computer heen
mozilla_herstel() {
    local naam="$1" proces="$2" bronmap="$3" oude_ini="$4"; shift 4
    [[ -d "$bronmap" ]] || return 0
    local profielen=() p
    for p in "$bronmap"/*/; do [[ -d "$p" ]] && profielen+=("${p%/}"); done
    ((${#profielen[@]})) || return 0

    local root='' r
    for r in "$@"; do [[ -f "$r/profiles.ini" ]] && { root="$r"; break; }; done
    if [[ -z "$root" ]]; then
        waarsch "$naam: installeer $naam, start het één keer en sluit het weer; draai daarna deze stap opnieuw."
        return 0
    fi
    local nieuw
    nieuw="$(mozilla_standaard "$root/profiles.ini")"
    [[ -z "$nieuw" ]] && { waarsch "$naam: geen profiel gevonden in $root/profiles.ini"; return 0; }
    [[ "$nieuw" == /* ]] || nieuw="$root/$nieuw"

    local gekozen="${profielen[0]}"
    if ((${#profielen[@]} > 1)); then
        local oud='' args=() b
        [[ -f "$oude_ini" ]] && oud="$(basename "$(mozilla_standaard "$oude_ini" | tr '\\' '/')")"
        for p in "${profielen[@]}"; do
            b="$(basename "$p")"
            args+=("$b" "$(du -sh "$p" 2>/dev/null | cut -f1)" "$([[ "$b" == "$oud" ]] && echo ON || echo OFF)")
        done
        kies "Welk $naam-profiel terugzetten? (kies er één)" "${args[@]}"
        ((${#KEUZE[@]})) || return 0
        gekozen="$bronmap/${KEUZE[0]}"
    fi

    ja_nee "$naam-profiel $(basename "$gekozen") terugzetten naar $nieuw? Wat daar nu staat wordt overschreven." || return 0
    wacht_gesloten "$naam" "$proces" || return 0
    if kopieer "$gekozen" "$nieuw" --exclude compatibility.ini --exclude parent.lock --exclude lock; then
        # Windows-paden (downloadmap, mailmappen) werken hier niet. Zonder die regels valt het
        # programma terug op de standaard of op de relatieve variant (*-rel) die er ook in staat.
        [[ -f "$nieuw/prefs.js" ]] && sed -i '/user_pref("[^"]*", "[A-Za-z]:\\\\/d' "$nieuw/prefs.js"
        ok "$naam: profiel teruggezet naar $nieuw"
    else
        fout "$naam: fouten bij kopiëren, zie $LOG"
    fi
}

# ---------------------------------------------------------------------------
# Stappen
# ---------------------------------------------------------------------------

stap_apps() {
    kop "Programma's installeren (Flathub)"
    local tsv="$LINUX/apps.tsv"
    if [[ -s "$tsv" ]]; then
        if ! command -v flatpak >/dev/null; then
            waarsch 'Flatpak is niet geïnstalleerd.'
            ja_nee 'Flatpak nu installeren (vraagt om je wachtwoord)?' && pakket flatpak
        fi
        if command -v flatpak >/dev/null; then
            flatpak remote-add --user --if-not-exists flathub https://dl.flathub.org/repo/flathub.flatpakrepo
            local args=() id naam van aan cmd
            while IFS=$'\t' read -r id naam van aan cmd; do
                [[ -z "$id" ]] && continue
                flatpak info "$id" >/dev/null 2>&1 && continue
                # Meegeleverd door de distributie (bijv. Firefox, LibreOffice)? Dan niet dubbel.
                [[ -n "${cmd:-}" ]] && command -v "$cmd" >/dev/null && continue
                args+=("$id" "$naam (was: $van)" "${aan:-ON}")
            done <"$tsv"

            if ((${#args[@]})); then
                kies "Welke programma's installeren?" "${args[@]}"
                local n=0
                for id in "${KEUZE[@]}"; do
                    echo "  $id ..."
                    if flatpak install --user -y --noninteractive flathub "$id" >>"$LOG" 2>&1; then n=$((n + 1))
                    else fout "$id: installatie mislukt (zie $LOG)"; fi
                done
                ((n)) && ok "$n programma's geïnstalleerd. Ze staan soms pas na opnieuw aanmelden in het menu."
            else
                ok "Alle programma's uit de lijst staan er al."
            fi
        fi
    else
        waarsch 'Geen apps.tsv in de back-up.'
    fi
    if [[ -s "$LINUX/zelf-zoeken.txt" ]]; then
        echo
        echo "  Hiervoor is geen vaste Linux-versie bekend:"
        sed 's/^/    /' "$LINUX/zelf-zoeken.txt"
    fi
}

# instelling MODUS NAAM: 'uitleg' geeft een korte omschrijving (begint met '-' als er niets
# terug te zetten valt), 'herstel' zet het terug
instelling() {
    local modus="$1" naam="$2" bron="$BACKUP/Instellingen/$2" doel
    case "$naam" in
        'VS Code')
            [[ $modus == uitleg ]] && { echo 'instellingen en extensies -> VS Code'; return; }
            local code=(code)
            doel="$HOME/.config/Code/User"
            if flatpak info com.visualstudio.code >/dev/null 2>&1; then
                doel="$HOME/.var/app/com.visualstudio.code/config/Code/User"
                code=(flatpak run com.visualstudio.code)
            elif ! command -v code >/dev/null; then
                code=()
            fi
            wacht_gesloten 'VS Code' code || return
            kopieer "$bron" "$doel" && ok "VS Code: instellingen -> $doel"
            local lijst="$BACKUP/Instellingen/vscode-extensies.txt" e n=0
            [[ -f "$lijst" ]] || return
            if ((${#code[@]})); then
                while IFS= read -r e; do
                    e="${e%$'\r'}"; e="${e#$'\xEF\xBB\xBF'}"
                    [[ -n "$e" ]] && "${code[@]}" --install-extension "$e" </dev/null >>"$LOG" 2>&1 && n=$((n + 1))
                done <"$lijst"
                ok "VS Code: $n extensies geïnstalleerd"
            else
                waarsch 'VS Code is niet geïnstalleerd; draai deze stap opnieuw na de installatie voor de extensies.'
            fi
            ;;
        Thunderbird)
            [[ $modus == uitleg ]] && { echo 'profiel (mail, adresboek, accounts) -> Thunderbird'; return; }
            mozilla_herstel Thunderbird 'thunderbird|thunderbird-bin' "$bron/Profiles" "$bron/profiles.ini" \
                "$HOME/.thunderbird" "$HOME/snap/thunderbird/common/.thunderbird" \
                "$HOME/.var/app/org.mozilla.Thunderbird/.thunderbird"
            ;;
        FileZilla)
            [[ $modus == uitleg ]] && { echo 'sites en instellingen -> FileZilla'; return; }
            doel="$HOME/.config/filezilla"
            flatpak info org.filezillaproject.Filezilla >/dev/null 2>&1 &&
                doel="$HOME/.var/app/org.filezillaproject.Filezilla/config/filezilla"
            wacht_gesloten FileZilla filezilla || return
            kopieer "$bron" "$doel" && ok "FileZilla -> $doel"
            ;;
        'Eigen lettertypen')
            [[ $modus == uitleg ]] && { echo 'lettertypen -> ~/.local/share/fonts'; return; }
            doel="$HOME/.local/share/fonts/migratie"
            kopieer "$bron" "$doel" && { fc-cache -f >/dev/null 2>&1; ok "Lettertypen geïnstalleerd in $doel"; }
            ;;
        'Outlook PST-bestanden')
            [[ $modus == uitleg ]] && { echo 'omzetten naar mbox (voor Thunderbird)'; return; }
            if ! command -v readpst >/dev/null; then
                ja_nee 'Hiervoor is readpst nodig (pst-utils). Nu installeren?' && pakket pst-utils libpst
            fi
            command -v readpst >/dev/null || { waarsch 'readpst ontbreekt; de PST-bestanden blijven in de back-up.'; return; }
            local f sub
            doel="$(xdg_doel Documents)/Outlook-mbox"
            for f in "$bron"/*.pst; do
                [[ -f "$f" ]] || continue
                sub="$doel/$(basename "$f" .pst)"
                mkdir -p "$sub"
                if readpst -q -r -o "$sub" "$f" >>"$LOG" 2>&1; then ok "$(basename "$f") -> $sub"
                else fout "$(basename "$f"): omzetten mislukt"; fi
            done
            echo '  In Thunderbird: installeer de add-on ImportExportTools NG en kies Extra > Importeren > mbox.'
            ;;
        'Outlook-handtekeningen' | 'Office-sjablonen' | 'Office-woordenlijsten')
            [[ $modus == uitleg ]] && { echo "bestanden -> Documenten/$naam"; return; }
            doel="$(xdg_doel Documents)/$naam"
            kopieer "$bron" "$doel" && ok "$naam -> $doel"
            case "$naam" in
                Outlook-handtekeningen) echo '  Thunderbird: Accountinstellingen > Handtekening uit bestand (.htm).' ;;
                Office-sjablonen) echo '  LibreOffice opent .dotx-sjablonen gewoon.' ;;
                Office-woordenlijsten) echo '  CUSTOM.DIC is een tekstbestand met je eigen woorden.' ;;
            esac
            ;;
        'eM Client')
            [[ $modus == uitleg ]] && { echo '- geen Linux-versie; zie uitleg bij Thunderbird'; return; }
            ;;
        'PuTTY-sessies' | 'WinSCP-sessies')
            [[ $modus == uitleg ]] && { echo '- staat omgezet in de stap SSH-sessies'; return; }
            ;;
        *)
            [[ $modus == uitleg ]] && { echo '- geen Linux-versie; blijft in de back-up'; return; }
            ;;
    esac
}

stap_instellingen() {
    kop 'Programma-instellingen terugzetten'
    local basis="$BACKUP/Instellingen"
    [[ -d "$basis" ]] || { waarsch 'Geen programma-instellingen in deze back-up.'; return; }
    local args=() d naam uitleg
    for d in "$basis"/*/; do
        [[ -d "$d" ]] || continue
        naam="$(basename "$d")"
        uitleg="$(instelling uitleg "$naam")"
        args+=("$naam" "$uitleg" "$([[ $uitleg == -* ]] && echo OFF || echo ON)")
    done
    ((${#args[@]})) || { waarsch 'Geen programma-instellingen in deze back-up.'; return; }
    echo '  Installeer de programma'"'"'s eerst (stap 2) en start Thunderbird één keer.'
    kies 'Welke instellingen terugzetten?' "${args[@]}"
    for naam in "${KEUZE[@]}"; do instelling herstel "$naam"; done

    if [[ -d "$basis/eM Client" ]]; then
        echo
        waarsch 'eM Client bestaat niet voor Linux; de database blijft in de back-up (voor een Windows- of Mac-pc).'
        cat <<'UITLEG'
  Overstappen naar Thunderbird:
   - Voeg je accounts toe met je e-mailadres en wachtwoord; Thunderbird zoekt de servers zelf op.
     Je mail via IMAP of Exchange staat op de server en komt vanzelf terug.
   - Mail in lokale mappen van eM Client: exporteer die op de oude pc in eM Client
     (Menu > Exporteren > E-mails naar .eml) en importeer ze in Thunderbird met de add-on ImportExportTools NG.
UITLEG
    fi
}

chromium_bladwijzers() {  # BROWSER PROFIEL
    local browser="$1" profiel="$2" proces data='' d doel
    case "$browser" in
        Chrome) proces=chrome
            set -- "$HOME/.config/google-chrome" "$HOME/.var/app/com.google.Chrome/config/google-chrome" ;;
        Edge) proces=msedge
            set -- "$HOME/.config/microsoft-edge" "$HOME/.var/app/com.microsoft.Edge/config/microsoft-edge" ;;
        Brave) proces=brave
            set -- "$HOME/.config/BraveSoftware/Brave-Browser" "$HOME/.var/app/com.brave.Browser/config/BraveSoftware/Brave-Browser" ;;
        Vivaldi) proces='vivaldi|vivaldi-bin'
            set -- "$HOME/.config/vivaldi" "$HOME/.var/app/com.vivaldi.Vivaldi/config/vivaldi" ;;
        Opera) proces=opera
            set -- "$HOME/.config/opera" "$HOME/.var/app/com.opera.Opera/config/opera" ;;
        *) return ;;
    esac
    for d in "$@"; do
        # Opera gebruikt de datamap zelf als profiel
        if [[ $browser == Opera ]]; then doel="$d"; else doel="$d/$profiel"; fi
        [[ -f "$doel/Preferences" ]] && { data="$doel"; break; }
    done
    if [[ -z "$data" ]]; then
        waarsch "$browser/$profiel: niet gevonden. Installeer $browser, start het één keer en sluit het, of importeer Browsers/$browser/$profiel/bladwijzers.html in een andere browser."
        return
    fi
    wacht_gesloten "$browser" "$proces" || return
    [[ -f "$data/Bookmarks" ]] && cp -f "$data/Bookmarks" "$data/Bookmarks.voor-migratie"
    cp -f "$BACKUP/Browsers/$browser/$profiel/Bookmarks" "$data/Bookmarks" &&
        ok "$browser/$profiel: bladwijzers teruggezet ($data)"
}

stap_browsers() {
    kop 'Bladwijzers en Firefox-profiel terugzetten'
    local basis="$BACKUP/Browsers"
    [[ -d "$basis" ]] || { waarsch 'Geen browsers in deze back-up.'; return; }

    # Chromium-browsers: alleen 'Bookmarks'; instellingen zijn aan de oude pc gebonden
    local args=() bw profiel browser k
    for bw in "$basis"/*/*/Bookmarks; do
        [[ -f "$bw" ]] || continue
        profiel="$(basename "$(dirname "$bw")")"
        browser="$(basename "$(dirname "$(dirname "$bw")")")"
        args+=("$browser/$profiel" 'bladwijzers' ON)
    done
    if ((${#args[@]})); then
        echo '  Log je in de browser in met sync (Google- of Microsoft-account)? Dan komen de bladwijzers vanzelf terug.'
        kies 'Welke bladwijzers terugzetten?' "${args[@]}"
        for k in "${KEUZE[@]}"; do chromium_bladwijzers "${k%%/*}" "${k#*/}"; done
    fi

    # Firefox: het profiel is draagbaar tussen Windows en Linux, inclusief wachtwoorden
    mozilla_herstel Firefox 'firefox|firefox-bin' "$basis/Firefox" "$basis/Firefox/profiles.ini" \
        "$HOME/.config/mozilla/firefox" "$HOME/.mozilla/firefox" \
        "$HOME/snap/firefox/common/.mozilla/firefox" "$HOME/.var/app/org.mozilla.firefox/.mozilla/firefox"
}

stap_ssh() {
    kop 'SSH-sessies (uit PuTTY en WinSCP)'
    local cfg="$LINUX/ssh-config.txt" doel="$HOME/.ssh/config"
    [[ -s "$cfg" ]] || { waarsch 'Geen PuTTY- of WinSCP-sessies in deze back-up.'; return; }
    sed 's/^/    /' "$cfg"
    if [[ -f "$doel" ]] && grep -q '^# --- Migratie-assistent' "$doel"; then
        waarsch "Deze sessies staan al in $doel (eerder toegevoegd)."
        ja_nee 'Toch nog een keer toevoegen?' || return
    else
        ja_nee "Toevoegen aan $doel?" || return
    fi
    mkdir -p "$HOME/.ssh" && chmod 700 "$HOME/.ssh"
    { echo; echo "# --- Migratie-assistent $(date +%F) ---"; tr -d '\r' <"$cfg"; } >>"$doel"
    chmod 600 "$doel"
    ok "Toegevoegd aan $doel"
}

stap_bestanden() {
    kop 'Bestanden terugzetten'
    local thuis
    thuis="$(find "$BACKUP" -maxdepth 1 -type d -iname home -print -quit)"
    if [[ -z "$thuis" ]]; then
        waarsch "Geen map Home in $BACKUP. Is de stap 'Home-map kopiëren' op de oude pc gedaan?"
        return
    fi
    local args=() d naam aan k doel
    for d in "$thuis"/*/; do
        [[ -d "$d" ]] || continue
        naam="$(basename "$d")"
        # Typische Windows-mappen hebben op Linux geen nut
        case "$naam" in Favorites | Links | Searches | Contacts | 'Saved Games' | '3D Objects') aan=OFF ;; *) aan=ON ;; esac
        args+=("$naam" "-> $(xdg_doel "$naam")" "$aan")
    done
    ((${#args[@]})) || { waarsch "De map $thuis bevat geen mappen om terug te zetten."; return; }
    echo '  Er wordt niets verwijderd. Bestanden die hier al staan en nieuwer zijn, blijven staan.'
    kies 'Welke mappen terugzetten?' "${args[@]}"
    ((${#KEUZE[@]})) || return
    echo "  Vrij in je home-map: $(df -h --output=avail "$HOME" 2>/dev/null | tail -1 | tr -d ' ')"

    for k in "${KEUZE[@]}"; do
        doel="$(xdg_doel "$k")"
        echo "  $k -> $doel"
        if kopieer "$thuis/$k" "$doel" --update --info=progress2; then ok "$k -> $doel"
        else fout "$k: fouten bij kopiëren, zie $LOG"; fi
    done

    # ssh weigert sleutels die voor anderen leesbaar zijn; van een Windows-schijf is alles leesbaar
    if [[ -d "$HOME/.ssh" ]]; then
        chmod 700 "$HOME/.ssh"
        find "$HOME/.ssh" -type f ! -name '*.pub' ! -name 'known_hosts*' -exec chmod 600 {} +
        grep -qsE 'IdentityFile[[:space:]]+"?[A-Za-z]:' "$HOME/.ssh/config" &&
            waarsch '~/.ssh/config bevat Windows-paden (C:\...). Pas die aan naar ~/.ssh/...'
    fi
}

stap_wachtwoorden() {
    kop 'Wachtwoorden'
    local map="$BACKUP/Wachtwoorden" a z='' tmp
    local archieven=() csvs=()
    for a in "$map"/*.7z; do [[ -f "$a" ]] && archieven+=("$a"); done
    for a in "$map"/*.csv; do [[ -f "$a" ]] && csvs+=("$a"); done
    if ((${#csvs[@]})); then
        waarsch "Er staan ONVERSLEUTELDE CSV-bestanden in $map. Importeer ze en verwijder ze daarna."
        echo '  Importeren: Firefox via about:logins > menu (...) > Importeren uit een bestand,'
        echo '  Chrome/Brave/Edge via de wachtwoordinstellingen, of in je wachtwoordmanager.'
    fi
    if ! ((${#archieven[@]})); then
        ((${#csvs[@]})) ||
            waarsch 'Geen geëxporteerde wachtwoorden in deze back-up (Firefox-wachtwoorden zitten in het profiel).'
        return
    fi
    for a in 7z 7zz 7za; do command -v "$a" >/dev/null && { z="$a"; break; }; done
    if [[ -z "$z" ]]; then
        ja_nee 'Hiervoor is 7-Zip nodig. Nu installeren?' && pakket 7zip p7zip-full p7zip
        for a in 7z 7zz 7za; do command -v "$a" >/dev/null && { z="$a"; break; }; done
    fi
    [[ -n "$z" ]] || { fout '7-Zip ontbreekt; pak het archief later zelf uit.'; return; }

    # Uitpakken in het geheugen (XDG_RUNTIME_DIR is een tmpfs), zodat de CSV nooit op schijf komt
    tmp="$(mktemp -d "${XDG_RUNTIME_DIR:-/tmp}/wachtwoorden.XXXXXX")"
    chmod 700 "$tmp"
    for a in "${archieven[@]}"; do
        echo "  $(basename "$a"): geef het archiefwachtwoord"
        "$z" x -o"$tmp" "$a" >>"$LOG" || fout "Uitpakken van $(basename "$a") mislukt (verkeerd wachtwoord?)"
    done
    echo "  De CSV-bestanden staan in: $tmp"
    cat <<'UITLEG'
  Importeren:
   - Firefox: about:logins > menu (...) > Importeren uit een bestand
   - Chrome/Brave/Edge: wachtwoordinstellingen > Wachtwoorden importeren
   - of in je wachtwoordmanager (Bitwarden, KeePassXC)
UITLEG
    command -v xdg-open >/dev/null && ja_nee 'De map nu openen?' && xdg-open "$tmp" >/dev/null 2>&1
    read -r -p '  Druk op Enter als je klaar bent; de uitgepakte CSV-bestanden worden dan verwijderd. ' _
    find "$tmp" -type f -exec shred -u {} + 2>/dev/null
    rm -rf "$tmp"
    ok 'Uitgepakte wachtwoorden verwijderd.'
}

alles() {
    stap_apps
    echo
    echo '  Start nu Firefox en Thunderbird (als je die gebruikt) één keer en sluit ze weer,'
    read -r -p '  zodat er een profiel is om in terug te zetten. Druk daarna op Enter. ' _
    stap_instellingen
    stap_browsers
    stap_bestanden      # eerst: daarin zit mogelijk je oude ~/.ssh/config, waar stap_ssh dan aan toevoegt
    stap_ssh
    stap_wachtwoorden
}

# ---------------------------------------------------------------------------
# Start
# ---------------------------------------------------------------------------

# De back-upmap: een map met HERSTELLEN.txt of rapport.json. Gezocht wordt in het meegegeven pad, de
# map boven Linux/ waar dit script normaal staat, en de huidige map; anders vraagt het script ernaar.
is_backup() { [[ -n "$1" && ( -f "$1/HERSTELLEN.txt" || -f "$1/rapport.json" ) ]]; }
zoek_backup() {
    local k
    for k in "${1:-}" "$(dirname "$HIER")" "$HIER" "$PWD" "$(dirname "$PWD")"; do
        is_backup "$k" && { BACKUP="$(cd -- "$k" >/dev/null && pwd)"; break; }
    done
    while [[ -z "$BACKUP" ]]; do
        waarsch "Geen back-upmap gevonden (gezocht bij $HIER en $PWD)."
        read -r -e -p '  Pad naar de back-upmap (Migratie_<pc>_<datum>), leeg = stoppen: ' k || exit 1
        [[ -z "$k" ]] && exit 1
        k="${k/#\~/$HOME}"
        if is_backup "$k"; then BACKUP="$(cd -- "$k" >/dev/null && pwd)"
        else fout "In $k staat geen HERSTELLEN.txt of rapport.json."; fi
    done
    LINUX="$BACKUP/Linux"
}

[[ "$(uname -s)" == Linux ]] || waarsch 'Dit script is bedoeld voor Linux.'
if ((EUID == 0)); then
    fout 'Start dit script als jezelf, niet met sudo: anders komt alles in de map van root terecht.'
    exit 1
fi

zoek_backup "${1:-}"
pc="$(sed -n 's/.*"Computer": *"\([^"]*\)".*/\1/p' "$BACKUP/rapport.json" 2>/dev/null | head -1)"
while true; do
    printf '\n\e[36m  ==== Migratie-assistent: herstellen op Linux ====\e[0m\n'
    echo "  Back-up: $BACKUP${pc:+ (van $pc)}"
    echo '  1. Alles achter elkaar'
    echo "  2. Programma's installeren (Flathub)"
    echo '  3. Programma-instellingen terugzetten'
    echo '  4. Bladwijzers en Firefox-profiel terugzetten'
    echo '  5. Bestanden terugzetten'
    echo '  6. SSH-sessies (uit PuTTY/WinSCP) toevoegen'
    echo '  7. Wachtwoorden uitpakken en importeren'
    echo '  Q. Stoppen'
    read -r -p '  Keuze: ' keuze || break
    case "$keuze" in
        1) alles ;; 2) stap_apps ;; 3) stap_instellingen ;; 4) stap_browsers ;;
        5) stap_bestanden ;; 6) stap_ssh ;; 7) stap_wachtwoorden ;;
        q | Q) break ;;
    esac
done
echo "  Log: $LOG"
'@

function Write-LinuxBestand([string]$Pad, [string]$Tekst) {
    # bash struikelt over CR's en een BOM
    [IO.File]::WriteAllText($Pad, ($Tekst -replace "`r`n", "`n"), (New-Object System.Text.UTF8Encoding $false))
}

# PuTTY- en WinSCP-sessies als blokken voor ~/.ssh/config
function Get-SshSessies {
    $sessies = New-Object System.Collections.Generic.List[object]

    $putty = 'HKCU:\Software\SimonTatham\PuTTY\Sessions'
    if (Test-Path -LiteralPath $putty) {
        foreach ($k in Get-ChildItem -LiteralPath $putty) {
            $p = Get-ItemProperty -LiteralPath $k.PSPath
            if (-not $p.HostName -or ($p.Protocol -and $p.Protocol -ne 'ssh')) { continue }
            $hostNaam = $p.HostName
            $gebruiker = $p.UserName
            if ($hostNaam -match '^(.+)@(.+)$') { $gebruiker = $Matches[1]; $hostNaam = $Matches[2] }
            $sessies.Add([pscustomobject]@{ Bron = 'PuTTY'; Naam = [uri]::UnescapeDataString($k.PSChildName)
                Host = $hostNaam; User = $gebruiker; Port = $p.PortNumber; Sleutel = $p.PublicKeyFile })
        }
    }

    $winscp = 'HKCU:\Software\Martin Prikryl\WinSCP 2\Sessions'
    if (Test-Path -LiteralPath $winscp) {
        foreach ($k in Get-ChildItem -LiteralPath $winscp) {
            $p = Get-ItemProperty -LiteralPath $k.PSPath
            # FSProtocol 0 = SCP, 1/2 = SFTP; FTP, WebDAV en S3 hebben niets met ssh te maken
            if (-not $p.HostName -or ($null -ne $p.FSProtocol -and $p.FSProtocol -gt 2)) { continue }
            $sleutel = if ($p.PublicKeyFile) { [uri]::UnescapeDataString($p.PublicKeyFile) } else { $null }
            $sessies.Add([pscustomobject]@{ Bron = 'WinSCP'; Naam = [uri]::UnescapeDataString($k.PSChildName)
                Host = $p.HostName; User = $p.UserName; Port = $p.PortNumber; Sleutel = $sleutel })
        }
    }

    $sb = New-Object System.Text.StringBuilder
    $gezien = @{}
    foreach ($s in $sessies) {
        $alias = ($s.Naam -replace '[^A-Za-z0-9._-]+', '-').Trim('-')
        if (-not $alias -or $gezien[$alias]) { continue }
        $gezien[$alias] = $true
        [void]$sb.AppendLine("# Uit $($s.Bron): $($s.Naam)")
        [void]$sb.AppendLine("Host $alias")
        [void]$sb.AppendLine("    HostName $($s.Host)")
        if ($s.User) { [void]$sb.AppendLine("    User $($s.User)") }
        if ($s.Port -and $s.Port -ne 22) { [void]$sb.AppendLine("    Port $($s.Port)") }
        if ($s.Sleutel) {
            [void]$sb.AppendLine("    # Sleutel op Windows: $($s.Sleutel)")
            [void]$sb.AppendLine("    # Omzetten: puttygen <bestand>.ppk -O private-openssh -o ~/.ssh/$alias  (pakket putty-tools)")
            [void]$sb.AppendLine("    # IdentityFile ~/.ssh/$alias")
        }
        [void]$sb.AppendLine('')
    }
    return $sb.ToString()
}

function Export-LinuxPlan {
    $run = Get-RunMap
    Write-Kop 'Voorbereiden op Linux'
    $map = Join-Path $run 'Linux'
    New-Item -ItemType Directory -Force -Path $map | Out-Null

    $programmas = @($script:Rapport.Programmas)
    if (-not $programmas) { $programmas = @(Get-Programmas) + @(Get-StoreApps) }

    $vertaling = @(foreach ($p in $programmas) {
        if ($p.Naam -match $script:LinuxOverslaan) { continue }
        $hit = $script:LinuxApps | Where-Object { $p.Naam -match $_.Patroon } | Select-Object -First 1
        [pscustomobject]@{
            Windows = $p.Naam
            Linux   = if ($hit) { $hit.Naam } else { '' }
            Flatpak = if ($hit) { $hit.Id } else { '' }
            Advies  = if (-not $hit) { 'Zoek een alternatief op flathub.org of alternativeto.net' } else { $hit.Opm }
            Def     = $hit
        }
    })

    # apps.tsv voor het herstelscript: id, naam, was, ON/OFF, commando (leeg veld als laatste)
    $regels = @($vertaling | Where-Object { $_.Flatpak } | Group-Object Flatpak | ForEach-Object {
        $eerste = $_.Group[0]
        $van = (@($_.Group | ForEach-Object { $_.Windows } | Select-Object -Unique -First 3) -join ', ') -replace '\t', ' '
        $aan = if ($eerste.Def.Uit) { 'OFF' } else { 'ON' }
        "{0}`t{1}`t{2}`t{3}`t{4}" -f $_.Name, $eerste.Linux, $van, $aan, $eerste.Def.Cmd
    })
    Write-LinuxBestand (Join-Path $map 'apps.tsv') (($regels -join "`n") + "`n")

    $zelf = @($vertaling | Where-Object { -not $_.Flatpak } | Sort-Object Windows -Unique |
        ForEach-Object { "$($_.Windows): $($_.Advies)" })
    Write-LinuxBestand (Join-Path $map 'zelf-zoeken.txt') (($zelf -join "`n") + "`n")

    $overzicht = @($vertaling | Select-Object Windows, Linux, Flatpak, Advies)
    $overzicht | Export-Csv -LiteralPath (Join-Path $map 'programmas-linux.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $script:Rapport.Linux = $overzicht
    Write-Ok "$($regels.Count) Linux-apps klaargezet, $($zelf.Count) programma's met alleen een advies -> Linux\programmas-linux.csv"

    $ssh = Get-SshSessies
    if ($ssh) {
        Write-LinuxBestand (Join-Path $map 'ssh-config.txt') $ssh
        Write-Ok "PuTTY/WinSCP: $(([regex]::Matches($ssh, '(?m)^Host ')).Count) sessies -> Linux\ssh-config.txt"
    }

    Write-LinuxBestand (Join-Path $map 'herstel-linux.sh') ($script:LinuxHerstelScript + "`n")
    Write-LinuxBestand (Join-Path $run 'LINUX.txt') $script:LinuxUitleg
    Write-Ok 'Herstelscript: Linux\herstel-linux.sh (uitleg in LINUX.txt)'

    if (-not (Test-Path -LiteralPath (Join-Path $run 'Home'))) {
        Write-Let 'Vergeet de stappen Browsers, Programma-instellingen en Home-map kopiëren niet; het herstelscript gebruikt die.'
    }
}

# ---------------------------------------------------------------------------
# Herstellen op de nieuwe laptop
# ---------------------------------------------------------------------------

# Kiest (eenmalig) de back-upmap om uit te herstellen en start de logging
function Open-Backup {
    if ($script:Bron) { return $script:Bron }

    $pad = $script:Herstel
    if (-not $pad) {
        if ($Gui) {
            Add-Type -AssemblyName System.Windows.Forms
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Kies de back-upmap (Migratie_<pc>_<datum>) om terug te zetten'
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'Geen back-upmap gekozen.' }
            $pad = $dlg.SelectedPath
        } else {
            $pad = Read-Host '  Back-upmap (bijv. E:\Migratie_OUDEPC_20261014_0930)'
        }
    }
    if (-not $pad -or -not (Test-Path -LiteralPath (Join-Path $pad 'HERSTELLEN.txt'))) {
        $script:Herstel = $null
        throw "'$pad' is geen back-upmap van de migratie-assistent (HERSTELLEN.txt ontbreekt)."
    }
    $pad = (Resolve-Path -LiteralPath $pad).ProviderPath

    $rapportPad = Join-Path $pad 'rapport.json'
    if (Test-Path -LiteralPath $rapportPad) {
        try { $r = Get-Content -LiteralPath $rapportPad -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $r = $null }
        if ($r) {
            Write-Host "  Back-up van $($r.Computer) ($($r.Gebruiker), $($r.Datum))"
            if ($r.Computer -eq $env:COMPUTERNAME) {
                Write-Let 'Dit is dezelfde pc als waar de back-up gemaakt is. Herstellen hoort op de NIEUWE laptop.'
                if (-not (Read-JaNee 'Toch doorgaan?')) { $script:Herstel = $null; throw 'Afgebroken.' }
            }
        }
    }

    # De back-up kan op een alleen-lezen NAS-share staan; de logs gaan daarom naar deze laptop
    $logMap = Join-Path $env:LOCALAPPDATA 'Migratie-Assistent'
    New-Item -ItemType Directory -Force -Path $logMap | Out-Null
    $stempel = Get-Date -Format 'yyyyMMdd_HHmm'
    $script:HerstelLog = Join-Path $logMap "herstel_$stempel.log"
    try { Start-Transcript -Path (Join-Path $logMap "transcript_$stempel.log") | Out-Null } catch { }

    $script:Bron = $pad
    Write-Ok "Back-up: $pad"
    return $pad
}

function Restore-Programmas {
    $bron = Open-Backup
    Write-Kop "Programma's installeren"
    $json = Join-Path $bron 'winget-apps.json'

    if (-not (Test-Path -LiteralPath $json)) {
        Write-Let 'Geen winget-apps.json in de back-up. Gebruik programmas.csv als lijst om zelf te installeren.'
    } elseif (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Let 'winget ontbreekt. Installeer of update App Installer uit de Microsoft Store en draai deze stap opnieuw.'
    } else {
        $data = Get-Content -LiteralPath $json -Raw -Encoding UTF8 | ConvertFrom-Json
        $pakketten = @(foreach ($s in $data.Sources) {
            foreach ($p in $s.Packages) { [pscustomobject]@{ Id = $p.PackageIdentifier; Bron = $s.SourceDetails.Name } }
        })
        $keuze = @(Select-Items -Items $pakketten -Titel "Welke programma's installeren?" -Kolommen @('Id', 'Bron') -Voorstel { $true })

        if ($keuze) {
            # Alleen de gekozen pakketten in een eigen importbestand zetten
            $ids = @($keuze | ForEach-Object { $_.Id })
            foreach ($s in $data.Sources) { $s.Packages = @($s.Packages | Where-Object { $ids -contains $_.PackageIdentifier }) }
            $data.Sources = @($data.Sources | Where-Object { $_.Packages })
            $tmp = Join-Path $env:TEMP 'migratie-winget.json'
            [IO.File]::WriteAllText($tmp, ($data | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding $false))

            Write-Host "  $($ids.Count) programma's installeren. Dit kan even duren; wat er al op staat wordt overgeslagen."
            & winget import -i $tmp --accept-package-agreements --accept-source-agreements --ignore-unavailable
            if ($LASTEXITCODE -eq 0) { Write-Ok 'winget import klaar' }
            else { Write-Let "winget meldt fouten (code $LASTEXITCODE); zie de uitvoer hierboven." }
            Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
            Update-Pad
        }
    }

    $niet = Join-Path $bron 'winget-niet-beschikbaar.txt'
    if ((Test-Path -LiteralPath $niet) -and (Get-Item -LiteralPath $niet).Length -gt 3) {
        Write-Host ''
        Write-Host "  Deze programma's stonden niet in winget; installeer ze zelf:"
        Get-Content -LiteralPath $niet -Encoding UTF8 | ForEach-Object { Write-Host "    $_" }
    }
}

function Restore-Instellingen {
    $bron = Open-Backup
    Write-Kop 'Programma-instellingen terugzetten'
    $basis = Join-Path $bron 'Instellingen'
    $manifestPad = Join-Path $basis 'instellingen.json'
    if (-not (Test-Path -LiteralPath $manifestPad)) { Write-Let 'Geen programma-instellingen in deze back-up.'; return }

    $items = Get-Content -LiteralPath $manifestPad -Raw -Encoding UTF8 | ConvertFrom-Json
    $lijst = @(foreach ($i in $items) {
        $map = Join-Path $basis $i.Naam
        if (-not (Test-Path -LiteralPath $map)) { continue }
        $doel = if ($i.Soort -eq 'Register') { $i.Pad } else { [Environment]::ExpandEnvironmentVariables($i.Pad) }
        if ($i.Naam -eq 'eM Client') {
            # Heeft eM Client op deze pc een eigen databasemap, dan leest het alleen daaruit
            $eigen = (Get-ItemProperty -LiteralPath 'HKCU:\SOFTWARE\eM Client' -ErrorAction SilentlyContinue).DBLocation
            if ($eigen) { $doel = $eigen }
        }
        [pscustomobject]@{ Naam = $i.Naam; Naar = $doel; Soort = $i.Soort; Proces = $i.Proces; Filter = $i.Filter
            Vervangen = [bool]$i.Vervangen; Map = $map }
    })
    if (-not $lijst) { Write-Let 'Geen programma-instellingen in deze back-up.'; return }

    Write-Host '  Installeer de programma''s eerst; bestaande instellingen met dezelfde naam worden overschreven.'
    $keuze = @(Select-Items -Items $lijst -Titel 'Welke instellingen terugzetten?' -Kolommen @('Naam', 'Naar') -Voorstel { $true })

    foreach ($k in $keuze) {
        if (-not (Confirm-Gesloten $k.Naam $k.Proces)) { Write-Let "$($k.Naam): overgeslagen, het programma draait nog."; continue }
        switch ($k.Soort) {
            'Register' {
                # reg.exe meldt ook succes via stderr
                & { $ErrorActionPreference = 'Continue'; reg.exe import (Join-Path $k.Map 'instellingen.reg') 2>&1 } | Out-Null
                if ($LASTEXITCODE -eq 0) { Write-Ok "$($k.Naam): geïmporteerd in het register" }
                else { Write-Fout "$($k.Naam): reg import mislukt (code $LASTEXITCODE)" }
            }
            'Lettertypen' {
                # Via de map Lettertypen van de shell, net als 'Installeren' in Verkenner
                $fonts = (New-Object -ComObject Shell.Application).Namespace(0x14)
                $n = 0
                foreach ($f in Get-ChildItem -LiteralPath $k.Map -File) {
                    $alAanwezig = (Test-Path -LiteralPath (Join-Path $k.Naar $f.Name)) -or
                                  (Test-Path -LiteralPath (Join-Path "$env:windir\Fonts" $f.Name))
                    if (-not $alAanwezig) { $fonts.CopyHere($f.FullName, 0x14); $n++ }
                }
                Write-Ok "$($k.Naam): $n lettertypen geïnstalleerd"
            }
            default {
                # Een database niet mengen met een (lege) database die hier al is aangemaakt
                if ($k.Vervangen -and (Get-ChildItem -LiteralPath $k.Naar -Force -ErrorAction SilentlyContinue | Select-Object -First 1)) {
                    $opzij = '{0}.voor-migratie-{1}' -f (Split-Path $k.Naar -Leaf), (Get-Date -Format 'yyyyMMdd_HHmm')
                    Rename-Item -LiteralPath $k.Naar -NewName $opzij
                    Write-Let "$($k.Naam): wat hier al stond is hernoemd naar $opzij"
                }
                $extra = @(if ($k.Filter) { $k.Filter } else { '/E' })
                $code = Invoke-Robocopy $k.Map $k.Naar $extra $script:HerstelLog
                if ($code -lt 8) { Write-Ok "$($k.Naam): teruggezet naar $($k.Naar)" }
                else { Write-Fout "$($k.Naam): fouten bij kopiëren (robocopy $code), zie $script:HerstelLog" }
                if ($k.Naam -eq 'eM Client' -and $code -lt 8) {
                    Write-Host '       Start eM Client en activeer je licentie opnieuw (Menu > Help > Licentie).'
                }
            }
        }
    }

    # VS Code-extensies opnieuw installeren
    $extLijst = Join-Path $basis 'vscode-extensies.txt'
    if (($keuze | Where-Object { $_.Naam -eq 'VS Code' }) -and (Test-Path -LiteralPath $extLijst)) {
        Update-Pad
        $code = @((Get-Command code -ErrorAction SilentlyContinue).Source,
                  "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd",
                  "$env:ProgramFiles\Microsoft VS Code\bin\code.cmd") |
            Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
        if (-not $code) {
            Write-Let 'VS Code is (nog) niet geïnstalleerd; extensies overgeslagen. Draai deze stap opnieuw na de installatie.'
        } else {
            $ext = @(Get-Content -LiteralPath $extLijst -Encoding UTF8 | Where-Object { $_.Trim() })
            foreach ($e in $ext) { & $code --install-extension $e.Trim() | Out-Null }
            Write-Ok "VS Code: $($ext.Count) extensies geïnstalleerd"
        }
    }
}

# Geeft het pad van het standaardprofiel uit een Firefox profiles.ini (relatief of absoluut)
function Get-FirefoxStandaard([string]$Ini) {
    if (-not (Test-Path -LiteralPath $Ini)) { return $null }
    $secties = @{}
    $huidig = $null
    foreach ($regel in Get-Content -LiteralPath $Ini) {
        if ($regel -match '^\[(.+)\]\s*$') { $huidig = $Matches[1]; $secties[$huidig] = @{}; continue }
        if ($huidig -and $regel -match '^([^=]+)=(.*)$') { $secties[$huidig][$Matches[1]] = $Matches[2].Trim() }
    }
    # Sinds Firefox 67 staat het standaardprofiel per installatie in een [Install...]-sectie
    $install = $secties.Keys | Where-Object { $_ -like 'Install*' -and $secties[$_].Default } | Select-Object -First 1
    if ($install) { return $secties[$install].Default }
    $profielen = @($secties.Keys | Where-Object { $_ -like 'Profile*' } | Sort-Object)
    $standaard = $profielen | Where-Object { $secties[$_].Default -eq '1' } | Select-Object -First 1
    if (-not $standaard) { $standaard = $profielen | Select-Object -First 1 }
    if ($standaard) { return $secties[$standaard].Path }
    return $null
}

function Restore-ChromiumBladwijzers($K) {
    $b = $K.Def
    $wat = "$($b.Naam) / $($K.Profiel)"
    if (-not (Test-Path -LiteralPath $b.Data)) {
        Write-Let "${wat}: $($b.Naam) staat (nog) niet op deze laptop. Importeer anders bladwijzers.html uit de back-up in een andere browser."
        return
    }
    # Opera gebruikt de datamap zelf als profiel
    $doel = if ($b.Naam -eq 'Opera') { $b.Data } else { Join-Path $b.Data $K.Profiel }
    if (-not (Test-Path -LiteralPath (Join-Path $doel 'Preferences'))) {
        if ($K.Profiel -eq 'Default' -or $b.Naam -eq 'Opera') {
            Write-Let "${wat}: start $($b.Naam) eerst één keer en sluit hem weer, en draai deze stap opnieuw."
        } else {
            Write-Let "${wat}: dit profiel bestaat hier niet. Maak in $($b.Naam) een extra profiel aan of importeer bladwijzers.html."
        }
        return
    }
    # Een draaiende browser overschrijft de bladwijzers bij het afsluiten
    if (-not (Confirm-Gesloten $b.Naam $b.Proces)) { Write-Let "${wat}: overgeslagen, de browser draait nog."; return }

    $bestaand = Join-Path $doel 'Bookmarks'
    if (Test-Path -LiteralPath $bestaand) { Copy-Item -LiteralPath $bestaand -Destination "$bestaand.voor-migratie" -Force }
    Copy-Item -LiteralPath $K.Bestand -Destination $bestaand -Force
    Write-Ok "${wat}: bladwijzers teruggezet (oude bladwijzers van deze laptop: Bookmarks.voor-migratie)"
}

function Restore-Browsers {
    $bron = Open-Backup
    Write-Kop 'Bladwijzers en Firefox-profiel terugzetten'
    $basis = Join-Path $bron 'Browsers'
    if (-not (Test-Path -LiteralPath $basis)) { Write-Let 'Geen browsers in deze back-up.'; return }

    # Chromium: alleen 'Bookmarks'. Preferences terugzetten is riskant: de browser controleert dat
    # bestand tegen 'Secure Preferences' (die aan de oude pc vastzit) en zet instellingen dan terug.
    $kandidaten = @(foreach ($b in $script:Chromium) {
        $bMap = Join-Path $basis $b.Naam
        if (-not (Test-Path -LiteralPath $bMap)) { continue }
        foreach ($p in Get-ChildItem -LiteralPath $bMap -Directory) {
            $bw = Join-Path $p.FullName 'Bookmarks'
            if (Test-Path -LiteralPath $bw) {
                [pscustomobject]@{ Browser = $b.Naam; Profiel = $p.Name; Bestand = $bw; Def = $b }
            }
        }
    })
    if ($kandidaten) {
        Write-Host '  Log je in de browser in met sync (Google- of Microsoft-account)? Dan komen de bladwijzers vanzelf terug.'
        $keuze = @(Select-Items -Items $kandidaten -Titel 'Welke bladwijzers terugzetten?' -Kolommen @('Browser', 'Profiel') -Voorstel { $true })
        foreach ($k in $keuze) { Restore-ChromiumBladwijzers $k }
    }

    # Firefox: het hele profiel over het (nieuwe, lege) standaardprofiel van deze laptop heen
    $ffBackup = Join-Path $basis 'Firefox'
    $ffProfielen = @(if (Test-Path -LiteralPath $ffBackup) { Get-ChildItem -LiteralPath $ffBackup -Directory })
    if (-not $ffProfielen) { return }

    $ffRoot = Join-Path $env:APPDATA 'Mozilla\Firefox'
    $nieuw = Get-FirefoxStandaard (Join-Path $ffRoot 'profiles.ini')
    if (-not $nieuw) {
        Write-Let 'Firefox: start Firefox eerst één keer en sluit hem weer (dan bestaat er een profiel), en draai deze stap opnieuw.'
        return
    }
    $nieuw = $nieuw -replace '/', '\'
    if (-not [IO.Path]::IsPathRooted($nieuw)) { $nieuw = Join-Path $ffRoot $nieuw }

    if ($ffProfielen.Count -eq 1) {
        $gekozen = $ffProfielen[0].FullName
    } else {
        $oud = Get-FirefoxStandaard (Join-Path $ffBackup 'profiles.ini')
        $oudNaam = if ($oud) { Split-Path ($oud -replace '/', '\') -Leaf } else { '' }
        $opties = @($ffProfielen | ForEach-Object {
            [pscustomobject]@{
                Profiel   = $_.Name
                Standaard = if ($_.Name -eq $oudNaam) { 'ja' } else { '' }
                Grootte   = Format-Grootte (Measure-Map $_.FullName).Bytes
                Pad       = $_.FullName
            }
        })
        $gekozen = @(Select-Items -Items $opties -Titel 'Welk Firefox-profiel terugzetten? (kies er één)' `
            -Kolommen @('Profiel', 'Standaard', 'Grootte') -Voorstel { $_.Standaard -eq 'ja' }) |
            Select-Object -First 1 | ForEach-Object { $_.Pad }
        if (-not $gekozen) { return }
    }

    if (-not (Read-JaNee "Firefox-profiel $(Split-Path $gekozen -Leaf) terugzetten? Wat nu in het Firefox-profiel van deze laptop staat wordt overschreven.")) { return }
    if (-not (Confirm-Gesloten 'Firefox' 'firefox')) { Write-Let 'Firefox: overgeslagen, Firefox draait nog.'; return }
    # compatibility.ini hoort bij de Firefox-versie van deze laptop; parent.lock is een vergrendeling
    $code = Invoke-Robocopy $gekozen $nieuw @('/E', '/XF', 'compatibility.ini', 'parent.lock') $script:HerstelLog
    if ($code -lt 8) { Write-Ok "Firefox: profiel teruggezet naar $nieuw" }
    else { Write-Fout "Firefox: fouten bij kopiëren (robocopy $code), zie $script:HerstelLog" }
}

# Waar een map uit Home\ op deze laptop hoort; bekende mappen kunnen omgeleid zijn (OneDrive)
function Get-HerstelDoel([string]$Map) {
    if ($Map -eq '_losse_bestanden') { return $env:USERPROFILE }
    if ($script:BekendeMappen[$Map]) {
        $p = Get-BekendePad $script:BekendeMappen[$Map]
        if ($p) { return $p }
    }
    return Join-Path $env:USERPROFILE $Map
}

function Restore-Bestanden {
    $bron = Open-Backup
    Write-Kop 'Bestanden terugzetten'
    $thuis = Join-Path $bron 'Home'
    if (-not (Test-Path -LiteralPath $thuis)) { Write-Let 'Geen bestanden uit de home-map in deze back-up.'; return }

    # Groottes uit het rapport van de oude pc, zodat de back-up niet opnieuw geteld hoeft te worden
    $grootte = @{}
    try {
        foreach ($h in (Get-Content -LiteralPath (Join-Path $bron 'rapport.json') -Raw -Encoding UTF8 | ConvertFrom-Json).Home) {
            $naam = if ($h.Type -eq 'Losse bestanden') { '_losse_bestanden' } else { $h.Map }
            $grootte[$naam] = [long]$h.Bytes
        }
    } catch { }

    $lijst = @(foreach ($d in Get-ChildItem -LiteralPath $thuis -Directory) {
        $b = [long]$grootte[$d.Name]
        [pscustomobject]@{
            Map     = $d.Name
            Grootte = if ($b) { Format-Grootte $b } else { '?' }
            Naar    = (Get-HerstelDoel $d.Name)
            Bytes   = $b
            Pad     = $d.FullName
        }
    })
    if (-not $lijst) { Write-Let 'Geen bestanden uit de home-map in deze back-up.'; return }

    Write-Host '  Er wordt niets verwijderd. Bestanden die hier al staan en nieuwer zijn, blijven staan.'
    $keuze = @(Select-Items -Items $lijst -Titel 'Welke mappen terugzetten?' -Kolommen @('Map', 'Grootte', 'Naar') -Voorstel { $true })
    if (-not $keuze) { Write-Let 'Niets gekozen.'; return }

    foreach ($m in $keuze | Where-Object { $_.Naar -like '*OneDrive*' }) {
        Write-Let "$($m.Map) gaat naar $($m.Naar); OneDrive uploadt dat daarna allemaal."
    }

    $nodig = ($keuze | Measure-Object Bytes -Sum).Sum
    try {
        $schijf = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($env:USERPROFILE))
        if ($nodig -and $schijf.AvailableFreeSpace -lt $nodig * 1.05) {
            Write-Fout "Te weinig ruimte op $($schijf.Name): $(Format-Grootte $schijf.AvailableFreeSpace) vrij, $(Format-Grootte $nodig) nodig."
            if (-not (Read-JaNee 'Toch beginnen?')) { return }
        }
    } catch { }

    $i = 0
    foreach ($m in $keuze) {
        Write-Progress -Activity 'Bestanden terugzetten' -Status "$($m.Map) ($($m.Grootte))" -PercentComplete ($i / $keuze.Count * 100)
        $i++
        # /XO: een bestand dat op deze laptop nieuwer is niet overschrijven
        $extra = @(if ($m.Map -eq '_losse_bestanden') { '/XO' } else { '/E', '/XO' })
        $extra += @('/COPY:DAT', '/DCOPY:T', '/MT:8')
        $code = Invoke-Robocopy $m.Pad $m.Naar $extra $script:HerstelLog
        if ($code -lt 8) { Write-Ok "$($m.Map) -> $($m.Naar)" }
        else { Write-Fout "$($m.Map): fouten bij kopiëren (robocopy $code), zie $script:HerstelLog" }
    }
    Write-Progress -Activity 'Bestanden terugzetten' -Completed
}

function Show-WachtwoordHerstel {
    $bron = Open-Backup
    Write-Kop 'Wachtwoorden'
    $map = Join-Path $bron 'Wachtwoorden'
    $archieven = @(Get-ChildItem -LiteralPath $map -Filter *.7z -File -ErrorAction SilentlyContinue)
    $csvs = @(Get-ChildItem -LiteralPath $map -Filter *.csv -File -ErrorAction SilentlyContinue)
    if (-not $archieven -and -not $csvs) {
        Write-Let 'Geen geëxporteerde wachtwoorden in deze back-up (Firefox-wachtwoorden zitten in het profiel).'
        return
    }
    if ($csvs) { Write-Let 'Er staan ONVERSLEUTELDE CSV-bestanden in de back-up. Verwijder ze na het importeren.' }
    Write-Host @'
  Wachtwoorden importeren gaat met de hand, zodat ze niet onversleuteld blijven rondslingeren:
  1. Pak het .7z-archief uit met 7-Zip (met het archiefwachtwoord), bijvoorbeeld naar je Bureaublad.
  2. Importeer de CSV:
     - Chrome:  chrome://password-manager/settings > Wachtwoorden importeren
     - Edge:    edge://settings/passwords > Importeren
     - Firefox: about:logins > menu (...) > Importeren uit een bestand
     - of in je wachtwoordmanager (Bitwarden, KeePassXC).
  3. Verwijder de uitgepakte CSV meteen en leeg de prullenbak.
'@
    if (Read-JaNee 'De map met wachtwoorden nu openen?') { Invoke-Item -LiteralPath $map }
}

function Invoke-HerstelAlles {
    Invoke-Stap { Open-Backup | Out-Null }
    if (-not $script:Bron) { return }
    Invoke-Stap { Restore-Programmas }
    Invoke-Stap { Restore-Instellingen }
    Invoke-Stap { Restore-Browsers }
    Invoke-Stap { Restore-Bestanden }
    Invoke-Stap { Show-WachtwoordHerstel }
}

function Show-HerstelMenu {
    Invoke-Stap { Open-Backup | Out-Null }
    if (-not $script:Bron) { return }
    while ($true) {
        Write-Host ''
        Write-Host '  ==== Herstellen op deze laptop ====' -ForegroundColor Cyan
        Write-Host "  Back-up: $script:Bron" -ForegroundColor DarkGray
        Write-Host '  1. Alles achter elkaar'
        Write-Host "  2. Programma's installeren (winget)"
        Write-Host '  3. Programma-instellingen terugzetten'
        Write-Host '  4. Bladwijzers en Firefox-profiel terugzetten'
        Write-Host '  5. Bestanden terugzetten'
        Write-Host '  6. Wachtwoorden: uitleg en map openen'
        Write-Host '  Q. Stoppen'
        switch (Read-Host '  Keuze') {
            '1' { Invoke-HerstelAlles }
            '2' { Invoke-Stap { Restore-Programmas } }
            '3' { Invoke-Stap { Restore-Instellingen } }
            '4' { Invoke-Stap { Restore-Browsers } }
            '5' { Invoke-Stap { Restore-Bestanden } }
            '6' { Invoke-Stap { Show-WachtwoordHerstel } }
            'q' { return }
        }
    }
}

# ---------------------------------------------------------------------------
# Menu
# ---------------------------------------------------------------------------

function Invoke-Stap([scriptblock]$Stap) {
    try { & $Stap } catch { Write-Fout $_.Exception.Message }
    try { Save-Rapport } catch { Write-Let "Rapport niet bijgewerkt: $($_.Exception.Message)" }
}

function Invoke-Alles {
    Invoke-Stap { Get-RunMap | Out-Null }
    if (-not $script:Run) { return }
    Invoke-Stap { Export-Programmas }
    Invoke-Stap { Backup-Browsers }
    Invoke-Stap { Export-Wachtwoorden }
    Invoke-Stap { Backup-Instellingen }
    Invoke-Stap { Invoke-HomeScan }
    Invoke-Stap { Copy-HomeMappen }
    if ($NaarLinux) { Invoke-Stap { Export-LinuxPlan } }
}

function Show-Menu {
    while ($true) {
        Write-Host ''
        Write-Host '  ==== Windows Migratie-Assistent ====' -ForegroundColor Cyan
        if ($script:Run) { Write-Host "  Back-upmap: $script:Run" -ForegroundColor DarkGray }
        Write-Host '  1. Alles achter elkaar'
        Write-Host "  2. Programma's inventariseren"
        Write-Host '  3. Browsers: bladwijzers, instellingen, extensies'
        Write-Host '  4. Wachtwoorden exporteren (begeleid, versleuteld)'
        Write-Host '  5. Programma-instellingen (Outlook, Office, VS Code, ...)'
        Write-Host '  6. Home-map scannen'
        Write-Host '  7. Home-map kopiëren'
        Write-Host '  8. Rapport openen'
        Write-Host '  L. Voorbereiden op Linux (nieuwe computer wordt een Linux-pc)'
        Write-Host '  H. Herstellen op deze laptop (vanuit een back-up)'
        Write-Host '  Q. Stoppen'
        switch (Read-Host '  Keuze') {
            '1' { Invoke-Alles }
            '2' { Invoke-Stap { Export-Programmas } }
            '3' { Invoke-Stap { Backup-Browsers } }
            '4' { Invoke-Stap { Export-Wachtwoorden } }
            '5' { Invoke-Stap { Backup-Instellingen } }
            '6' { Invoke-Stap { Invoke-HomeScan } }
            '7' { Invoke-Stap { Copy-HomeMappen } }
            '8' {
                if ($script:Run) { Invoke-Item (Join-Path $script:Run 'rapport.html') }
                else { Write-Let 'Er is nog geen rapport; voer eerst een stap uit.' }
            }
            'l' { Invoke-Stap { Export-LinuxPlan } }
            'h' { Show-HerstelMenu; return }
            'q' { return }
        }
    }
}

# ---------------------------------------------------------------------------
# Start
# ---------------------------------------------------------------------------

if ($env:OS -ne 'Windows_NT') { Write-Error 'Dit script draait alleen op Windows.'; exit 1 }

# Gestart vanuit een back-upmap (het script kopieert zichzelf daarheen)? Dan ligt herstellen voor de hand.
if (-not $Herstel -and -not $Doel -and $PSScriptRoot -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'HERSTELLEN.txt'))) {
    if (Read-JaNee "Dit script staat in een back-upmap ($PSScriptRoot). Daaruit herstellen op deze laptop?") { $Herstel = $PSScriptRoot }
}

try {
    if ($Herstel) {
        if ($Alles) { Invoke-HerstelAlles } else { Show-HerstelMenu }
    } elseif ($Alles) { Invoke-Alles } else { Show-Menu }
} finally {
    Save-Rapport
    if ($script:Run -or $script:Bron) { try { Stop-Transcript | Out-Null } catch { } }
    if ($script:Run) { Write-Ok "Klaar. Alles staat in $script:Run (zie rapport.html en HERSTELLEN.txt)" }
    if ($script:Bron) { Write-Ok "Klaar met herstellen. Log: $script:HerstelLog" }
}
