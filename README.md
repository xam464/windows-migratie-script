# Windows Migratie-Assistent

Een PowerShell-script dat een Windows 10-pc inventariseert en een back-up maakt, zodat je alles kunt
overzetten naar een nieuwe laptop. Die nieuwe laptop mag Windows draaien, maar ook Linux.

Er wordt alleen gekopieerd: op de oude pc wordt niets verwijderd.

## Wat het meeneemt

- **Programma's**: een lijst uit het register en de Store, plus een `winget`-export om alles in één keer
  opnieuw te installeren.
- **Browsers** (Chrome, Edge, Brave, Vivaldi, Opera, Firefox): bladwijzers (ook als HTML), instellingen en
  een lijst van extensies. Firefox gaat als heel profiel mee.
- **Wachtwoorden**: een begeleide export uit de browser, daarna versleuteld met 7-Zip (AES-256).
- **Programma-instellingen**: Outlook (handtekeningen en PST-bestanden), Office-sjablonen, VS Code,
  Notepad++, Thunderbird, FileZilla, Windows Terminal, Sticky Notes, eigen lettertypen, PuTTY en WinSCP.
- **Home-map**: een scan (grootte, aantal bestanden, grootste bestanden) en een kopie met robocopy,
  met een controle achteraf.

Alles komt in één map, `Migratie_<pc>_<datum>`, met een `rapport.html` en uitleg in `HERSTELLEN.txt`.

## Gebruik

Op de oude pc, met een externe schijf of NAS-share als doel:

```powershell
powershell -ExecutionPolicy Bypass -File .\Migratie-Assistent.ps1
```

| Optie | Wat het doet |
|---|---|
| `-Gui` | Kies de doelmap en maak selecties in vensters in plaats van in de terminal |
| `-Doel E:\` | Doelmap meteen opgeven |
| `-Alles` | Alle stappen achter elkaar, zonder menu |
| `-NaarLinux` | Met `-Alles`: de nieuwe computer wordt een Linux-pc (zie hieronder) |
| `-Herstel <map>` | Een back-up terugzetten op deze laptop |

## Herstellen op een nieuwe Windows-laptop

Het script kopieert zichzelf naar de back-upmap. Start het daar op de nieuwe laptop; het biedt dan zelf
aan om te herstellen. Het installeert de programma's met `winget` (je kiest welke) en zet instellingen,
bladwijzers, het Firefox-profiel en je bestanden terug. Bestanden die op de nieuwe laptop nieuwer zijn,
blijven staan.

## Overstappen naar Linux

Kies in het menu **L. Voorbereiden op Linux**. Het script:

- zoekt per Windows-programma de Linux-versie op Flathub (of geeft een advies als die er niet is);
- zet PuTTY- en WinSCP-sessies om naar `~/.ssh/config`;
- schrijft `Linux/herstel-linux.sh` in de back-up.

Op de Linux-pc, in de back-upmap (niet met `sudo`):

```bash
bash Linux/herstel-linux.sh
```

Dit installeert de programma's via Flatpak, zet instellingen, bladwijzers, het Firefox- en
Thunderbird-profiel en je bestanden terug (`Documents` gaat naar `~/Documenten`, enzovoort), en pakt de
wachtwoorden alleen in het geheugen uit om ze te importeren. Het werkt op Ubuntu, Mint, Debian, Fedora,
openSUSE en Arch.

## Vereisten

- Windows 10 of 11 met Windows PowerShell 5.1 (standaard aanwezig)
- Optioneel: `winget` (App Installer) en 7-Zip; het script biedt aan 7-Zip te installeren

## Status

Het script leest foutloos in en de onderdelen zijn getest met nagebootste Windows-onderdelen. Het
Linux-herstelscript is getest in een Ubuntu 24.04-container. Een volledige testrun op een echte
Windows 10-pc moet nog gebeuren.
