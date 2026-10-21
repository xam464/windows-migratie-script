# Windows Migratie-Assistent

Een PowerShell-script met een eigen venster om een Windows-pc over te zetten naar een nieuwe laptop,
en om een (nieuwe) pc klaar te maken. Het heeft drie onderdelen:

1. **Back-up maken** op de oude pc: programma's, browsers, wachtwoorden, programma-instellingen en je
   bestanden. Er wordt alleen gekopieerd; op de oude pc wordt niets verwijderd.
2. **Herstellen**: de back-up terugzetten op een nieuwe Windows-laptop, of de back-up klaarmaken voor
   een Linux-pc.
3. **Configureren**: software installeren (ESET, browsers, basisprogramma's, AnyDesk) en Windows
   instellen met Chris Titus WinUtil.

## Snel starten

Open PowerShell (gewoon, niet als administrator) en plak:

```powershell
irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1 | iex
```

Het venster opent vanzelf. Kies links wat je wilt doen en klik op een stap. Onderin zie je wat het
script doet; vragen en keuzelijsten verschijnen als venster.

Krijg je op een oudere Windows 10-pc een melding als *"Could not create SSL/TLS secure channel"*, gebruik
dan:

```powershell
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1 | iex
```

Met opties, bijvoorbeeld de tekstversie of een vaste doelmap:

```powershell
& ([scriptblock]::Create((irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1))) -Cli
```

`start.ps1` is een kleine starter: die downloadt `Migratie-Assistent.ps1` uit de nieuwste release en
start het met Windows PowerShell. Je kunt het script ook zelf downloaden en starten:

```powershell
powershell -ExecutionPolicy Bypass -File .\Migratie-Assistent.ps1
```

Voor de back-up zijn geen administratorrechten nodig. Het script werkt met de bestanden en instellingen
van de gebruiker die het start, dus start het als de gebruiker wiens gegevens je wilt meenemen. Bij het
installeren van software vraagt Windows zelf om toestemming.

## Opties

| Optie | Wat het doet |
|---|---|
| `-Cli` | De tekstversie met menu's in de console in plaats van het venster |
| `-Doel E:\` | Doelmap voor de back-up meteen opgeven |
| `-Herstel <map>` | Het venster opent bij Herstellen, met deze back-upmap |
| `-Alles` | Alle back-upstappen achter elkaar, zonder menu (tekstversie) |
| `-NaarLinux` | Met `-Alles`: de back-up ook klaarmaken voor een Linux-pc |

## Back-up maken

- **Programma's**: een lijst uit het register en de Store, plus een `winget`-export om alles in één keer
  opnieuw te installeren.
- **Browsers** (Chrome, Edge, Brave, Vivaldi, Opera, Firefox): bladwijzers (ook als HTML), instellingen en
  een lijst van extensies. Firefox gaat als heel profiel mee.
- **Wachtwoorden**: automatische export uit Chrome, Edge en andere Chromium-browsers naar CSV. Je
  bevestigt alleen met je Windows-pincode of -wachtwoord. De CSV-bestanden zijn leesbaar: bewaar de
  back-up goed en verwijder ze na het importeren.
- **Programma-instellingen**: Outlook (handtekeningen en PST-bestanden), Office-sjablonen, VS Code,
  Notepad++, Thunderbird, FileZilla, Windows Terminal, Sticky Notes, eigen lettertypen, PuTTY en WinSCP.
- **eM Client**: de hele database, met accounts, serverinstellingen, mail en de opgeslagen
  (versleutelde) wachtwoorden. Ook een eigen databasemap wordt gevonden. Op de nieuwe pc werkt alles
  weer na het terugzetten; alleen de licentie moet je opnieuw activeren.
- **Home-map**: een scan (grootte, aantal bestanden, grootste bestanden) en een kopie met robocopy,
  met een controle achteraf. Staan je bureaublad, Documenten of Afbeeldingen in OneDrive (OneDrive-back-up),
  dan gaan die ook mee, en bij het herstellen komen ze weer op het bureaublad of in Documenten terecht.

Alles komt in één map, `Migratie_<pc>_<datum>`, met een `rapport.html`, een log (`transcript.log`) en
uitleg in `HERSTELLEN.txt`.

## Herstellen

### Windows

Het script kopieert zichzelf naar de back-upmap. Start het daar op de nieuwe laptop; het biedt dan zelf
aan om te herstellen. Je kunt ook gewoon het `irm`-commando gebruiken en bij **Herstellen** de
back-upmap kiezen.

- Programma's installeren met `winget` (je kiest welke), en ESET Security aanbieden als virusscanner.
- Instellingen, bladwijzers, het Firefox-profiel en je bestanden terugzetten. Bestanden die op de nieuwe
  laptop nieuwer zijn, blijven staan.
- De wachtwoorden automatisch importeren in Chrome, Edge of een andere Chromium-browser. Staat de
  browser waar een bestand uit komt er niet op, dan biedt het script aan die te installeren, in de taal
  van Windows.

### Linux

Kies op de oude pc **Herstellen → Back-up klaarmaken voor Linux**. Het script:

- zoekt per Windows-programma de Linux-versie op Flathub (of geeft een advies als die er niet is);
- zet PuTTY- en WinSCP-sessies om naar `~/.ssh/config`;
- schrijft `Linux/herstel-linux.sh` in de back-up en toont het commando voor de Linux-pc.

Op de Linux-pc, in de back-upmap (niet met `sudo`):

```bash
bash Linux/herstel-linux.sh
```

Of altijd de nieuwste versie van het herstelscript, ook bij een oudere back-up:

```bash
bash <(curl -fsSL https://github.com/xam464/windows-migratie-script/releases/latest/download/herstel-linux.sh) /pad/naar/Migratie_<pc>_<datum>
```

Dit installeert de programma's via Flatpak, zet instellingen, bladwijzers, het Firefox- en
Thunderbird-profiel en je bestanden terug (`Documents` gaat naar `~/Documenten`, enzovoort), en helpt
bij het importeren van de wachtwoorden. Het werkt op Ubuntu, Mint, Debian, Fedora, openSUSE en Arch.

## Configureren

**Software installeren** toont een lijst met wat er al op staat. Alles gaat via `winget`, in de taal van
Windows als het pakket die heeft (Firefox in het Nederlands is bijvoorbeeld een apart pakket).

| Groep | Software |
|---|---|
| Beveiliging | ESET Security |
| Browsers | Google Chrome, Mozilla Firefox, Brave |
| Basis | 7-Zip, VLC, Adobe Acrobat Reader, Notepad++ |
| Hulp op afstand | AnyDesk |

Is de `winget` op de pc te oud voor een pakket (foutcode `-1978335225`, komt voor op oudere Windows
10-installaties), dan biedt het script aan winget bij te werken met de nieuwste versie van
[github.com/microsoft/winget-cli](https://github.com/microsoft/winget-cli) en probeert het opnieuw.

**Chris Titus WinUtil** start [WinUtil](https://github.com/ChrisTitusTech/winutil) als beheerder in een
eigen venster: Windows-instellingen, tweaks, programma's en updates. Het is een los hulpprogramma; het
script downloadt het rechtstreeks uit de nieuwste release op GitHub.

## Vereisten

- Windows 10 of 11 met Windows PowerShell 5.1 (standaard aanwezig)
- `winget` (App Installer) voor het installeren van programma's en software

## Waarom er geen handtekening op zit

De execution policy van Windows controleert alleen scriptbestanden die je start. Met `irm | iex`
voer je tekst uit, en de starter start het script met `-ExecutionPolicy Bypass`. Een handtekening
(Authenticode) is daarom niet nodig. Dit is dezelfde manier waarop bijvoorbeeld
[winutil](https://github.com/ChrisTitusTech/winutil) wordt gestart.

Let wel: `irm | iex` voert uit wat er op dat adres staat. Bekijk het script gerust eerst op GitHub.

## Een nieuwe versie uitbrengen

Maak op GitHub een release aan met een tag als `v0.7.0` (kleine `v`) en **vink "Set as the latest
release" uit**. De workflow in `.github/workflows/release.yml` controleert de scripts, hangt
`Migratie-Assistent.ps1`, `start.ps1` en `herstel-linux.sh` aan de release en maakt hem daarna pas de
nieuwste. Zo wijst `releases/latest/download/start.ps1` nooit naar een release zonder bestanden.

Vergeet je het uitvinken, dan werkt het ook, maar wijst die link de eerste paar seconden naar een release
waar de bestanden nog niet aan hangen. De starter probeert het downloaden van het script daarom een paar
keer.

## Status

Getest in een Windows 10 22H2-VM (Nederlands) met Edge en Chrome: het venster, de back-upstappen, de
wachtwoord-export en -import, herstellen, software installeren en WinUtil. Het Linux-herstelscript is
getest in een Ubuntu 24.04-container. Eerdere versies zijn gebruikt op een echte Windows 10-pc.
