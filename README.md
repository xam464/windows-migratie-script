# Windows Migratie-Assistent

Een PowerShell-script dat een Windows 10-pc inventariseert en een back-up maakt, zodat je alles kunt
overzetten naar een nieuwe laptop. Die nieuwe laptop mag Windows draaien, maar ook Linux.

Er wordt alleen gekopieerd: op de oude pc wordt niets verwijderd.

## Wat het meeneemt

- **Programma's**: een lijst uit het register en de Store, plus een `winget`-export om alles in één keer
  opnieuw te installeren.
- **Browsers** (Chrome, Edge, Brave, Vivaldi, Opera, Firefox): bladwijzers (ook als HTML), instellingen en
  een lijst van extensies. Firefox gaat als heel profiel mee.
- **Wachtwoorden**: automatische export uit Chrome, Edge en andere Chromium-browsers naar CSV; je bevestigt alleen met je Windows-pincode of -wachtwoord.
- **Programma-instellingen**: Outlook (handtekeningen en PST-bestanden), Office-sjablonen, VS Code,
  Notepad++, Thunderbird, FileZilla, Windows Terminal, Sticky Notes, eigen lettertypen, PuTTY en WinSCP.
- **eM Client**: de hele database, met accounts, serverinstellingen, mail en de opgeslagen
  (versleutelde) wachtwoorden. Ook een eigen databasemap wordt gevonden. Op de nieuwe pc werkt alles
  weer na het terugzetten; alleen de licentie moet je opnieuw activeren.
- **Home-map**: een scan (grootte, aantal bestanden, grootste bestanden) en een kopie met robocopy,
  met een controle achteraf. Staan je bureaublad, Documenten of Afbeeldingen in OneDrive (OneDrive-back-up),
  dan gaan die ook mee, en bij het herstellen komen ze weer op het bureaublad of in Documenten terecht.

Alles komt in één map, `Migratie_<pc>_<datum>`, met een `rapport.html` en uitleg in `HERSTELLEN.txt`.

## Snel starten

Open PowerShell op de oude pc (gewoon, niet als administrator) en plak:

```powershell
irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1 | iex
```

Met opties, bijvoorbeeld de grafische dialogen:

```powershell
& ([scriptblock]::Create((irm https://github.com/xam464/windows-migratie-script/releases/latest/download/start.ps1))) -Gui
```

`start.ps1` is een kleine starter: die downloadt `Migratie-Assistent.ps1` uit de nieuwste release en
start het met Windows PowerShell. Je kunt het script ook zelf downloaden en starten:

```powershell
powershell -ExecutionPolicy Bypass -File .\Migratie-Assistent.ps1
```

Administratorrechten zijn niet nodig. Het script werkt met de bestanden en instellingen van de
gebruiker die het start, dus start het als de gebruiker wiens gegevens je wilt meenemen.

## Opties

| Optie | Wat het doet |
|---|---|
| `-Gui` | Kies de doelmap en maak selecties in vensters in plaats van in de terminal |
| `-Doel E:\` | Doelmap meteen opgeven |
| `-Alles` | Alle stappen achter elkaar, zonder menu |
| `-NaarLinux` | Met `-Alles`: de nieuwe computer wordt een Linux-pc (zie hieronder) |
| `-Herstel <map>` | Een back-up terugzetten op deze laptop |

## Herstellen op een nieuwe Windows-laptop

Het script kopieert zichzelf naar de back-upmap. Start het daar op de nieuwe laptop; het biedt dan zelf
aan om te herstellen. Het installeert de programma's met `winget` (je kiest welke), biedt ESET Security
aan als virusscanner en zet instellingen, bladwijzers, het Firefox-profiel en je bestanden terug.
Bestanden die op de nieuwe laptop nieuwer zijn, blijven staan.

## Overstappen naar Linux

Kies in het menu **L. Voorbereiden op Linux**. Het script:

- zoekt per Windows-programma de Linux-versie op Flathub (of geeft een advies als die er niet is);
- zet PuTTY- en WinSCP-sessies om naar `~/.ssh/config`;
- schrijft `Linux/herstel-linux.sh` in de back-up.

Op de Linux-pc, in de back-upmap (niet met `sudo`):

```bash
bash Linux/herstel-linux.sh
```

Of altijd de nieuwste versie van het herstelscript, ook bij een oudere back-up:

```bash
bash <(curl -fsSL https://github.com/xam464/windows-migratie-script/releases/latest/download/herstel-linux.sh) /pad/naar/Migratie_<pc>_<datum>
```

Dit installeert de programma's via Flatpak, zet instellingen, bladwijzers, het Firefox- en
Thunderbird-profiel en je bestanden terug (`Documents` gaat naar `~/Documenten`, enzovoort), en pakt de
wachtwoorden alleen in het geheugen uit om ze te importeren. Het werkt op Ubuntu, Mint, Debian, Fedora,
openSUSE en Arch.

## Vereisten

- Windows 10 of 11 met Windows PowerShell 5.1 (standaard aanwezig)
- Optioneel: `winget` (App Installer)

## Waarom er geen handtekening op zit

De execution policy van Windows controleert alleen scriptbestanden die je start. Met `irm | iex`
voer je tekst uit, en de starter start het script met `-ExecutionPolicy Bypass`. Een handtekening
(Authenticode) is daarom niet nodig. Dit is dezelfde manier waarop bijvoorbeeld
[winutil](https://github.com/ChrisTitusTech/winutil) wordt gestart.

Let wel: `irm | iex` voert uit wat er op dat adres staat. Bekijk het script gerust eerst op GitHub.

## Een nieuwe versie uitbrengen

Maak op GitHub een release aan (bijvoorbeeld `v0.2.0`). De workflow in `.github/workflows/release.yml`
controleert de scripts en hangt `Migratie-Assistent.ps1` en `start.ps1` aan de release. De link
`releases/latest/download/start.ps1` wijst daarna vanzelf naar die versie.

## Status

Het script leest foutloos in en de onderdelen zijn getest met nagebootste Windows-onderdelen. Het
Linux-herstelscript is getest in een Ubuntu 24.04-container. Een volledige testrun op een echte
Windows 10-pc moet nog gebeuren.
