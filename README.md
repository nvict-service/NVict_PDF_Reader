# NVict Reader

<div align="center">

![NVict Reader Logo](https://www.nvict.nl/images/reader_logo.png)

**Professionele PDF Reader voor Windows**

[![Version](https://img.shields.io/badge/version-1.7.1-blue.svg)](https://github.com/nvict-service/nvict-reader/releases)
[![License](https://img.shields.io/badge/license-Proprietary-red.svg)](LICENSE)
[![Windows](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)

[Website](https://www.nvict.nl) • [Documentatie](https://github.com/nvict-service/nvict-reader/wiki) • [Updates](https://www.nvict.nl/software/updates/)

</div>

## 📖 Over NVict Reader

NVict Reader is een moderne, gebruiksvriendelijke PDF viewer ontwikkeld door NVict Service. De applicatie biedt een snelle en efficiënte manier om PDF-documenten te bekijken, afdrukken en beheren met een moderne gebruikersinterface.

### ✨ Belangrijkste Features

- **Modern Design** - Donker en licht thema met moderne UI
- **Multi-tab Interface** - Open meerdere PDF's tegelijk in tabs
- **Zoekfunctionaliteit** - Geavanceerd zoeken in PDF's
- **Bladwijzers** - Sla je favoriete pagina's op
- **Aantekeningen** - Voeg notities toe aan PDF pagina's
- **Direct Printen** - Print direct naar elke printer via Windows GDI (USB, netwerk, virtueel)
- **Standaard PDF Viewer** - Stel in als standaard PDF programma
- **Automatische Updates** - Blijf altijd up-to-date met scrollbare release notes
- **Single Instance** - Nieuw geopende bestanden worden automatisch in bestaand venster geladen
- **Smart File Management** - Voorkomt dat hetzelfde bestand dubbel wordt geopend
- **Thumbnail Preview** - Bekijk miniaturen van alle pagina's
- **Zoom & Rotatie** - Volledige controle over weergave

## 🚀 Installatie

### Optie 1: Gecompileerde Installer (Aanbevolen)

Download de laatste versie van de [Releases](https://github.com/nvict-service/nvict-reader/releases) pagina en voer de installer uit.

### Optie 2: Vanaf Broncode

**Vereisten:**
- Python 3.8 of hoger
- Windows 10/11

**Installatie:**

```bash
# Clone de repository
git clone https://github.com/nvict-service/nvict-reader.git
cd nvict-reader

# Installeer dependencies
pip install -r requirements.txt

# Voor Windows print functionaliteit (vereist)
pip install pywin32

# Start de applicatie
python NVict_Reader.py
```

## 🛠️ Bouwen vanaf Broncode

Om een standalone executable te maken:

```bash
# Installeer PyInstaller
pip install pyinstaller

# Maak executable
pyinstaller --onefile --windowed --icon=favicon.ico --add-data "favicon.ico;." --add-data "PDF_File_icon.ico;." NVict_Reader.py
```

De executable wordt aangemaakt in de `dist` folder.

## 📋 Systeemvereisten

- **Besturingssysteem:** Windows 10 of hoger (64-bit aanbevolen)
- **RAM:** Minimaal 2GB (4GB aanbevolen)
- **Schijfruimte:** 100MB vrije ruimte
- **Display:** Minimaal 1280x720 resolutie
- **Dependencies:** pywin32 voor print functionaliteit (v1.7+)

## 🎯 Gebruik

### Basisgebruik

1. **Open een PDF:** Klik op "Open bestand" of sleep een PDF naar het venster
2. **Navigeren:** Gebruik de pijltoetsen of klik op de paginaknoppen
3. **Zoeken:** Druk op Ctrl+F of klik op het zoek-icoon
4. **Afdrukken:** Druk op Ctrl+P of gebruik het menu

### Sneltoetsen

| Sneltoets | Actie |
|-----------|-------|
| `Ctrl+O` | Open bestand |
| `Ctrl+W` | Sluit tab |
| `Ctrl+F` | Zoeken |
| `Ctrl+P` | Afdrukken |
| `Ctrl+Plus` | Inzoomen |
| `Ctrl+Min` | Uitzoomen |
| `Ctrl+0` | Standaard zoom |
| `Ctrl+Tab` | Volgende tab |
| `Ctrl+Shift+Tab` | Vorige tab |
| `Pijl links/rechts` | Vorige/volgende pagina |

### Instellen als Standaard PDF Viewer

1. Open NVict Reader
2. Ga naar Instellingen → "Instellen als standaard"
3. Volg de instructies in Windows Instellingen

## 🔧 Configuratie

De applicatie slaat instellingen op in:
```
%LOCALAPPDATA%\NVict Service\NVict Reader\
```

Configuratiebestanden:
- `bookmarks.json` - Opgeslagen bladwijzers
- `settings.json` - Applicatie-instellingen
- `recent_files.json` - Recent geopende bestanden

## 🤝 Bijdragen

Dit is een proprietary project van NVict Service. Voor vragen of suggesties, neem contact op via [www.nvict.nl](https://www.nvict.nl).

## 📜 Licentie

Copyright © 2026 NVict Service. Alle rechten voorbehouden.

Deze software is eigendom van NVict Service en mag niet worden gedistribueerd, gewijzigd of gebruikt zonder expliciete toestemming.

## 🔐 Code Signing

Releases worden automatisch ondertekend via GitHub Actions met een geldig code signing certificaat voor Windows Authenticode.

## 📞 Contact & Support

- **Website:** [www.nvict.nl](https://www.nvict.nl)
- **Support:** Via website contact formulier
- **Updates:** Automatisch via de applicatie of [handmatig downloaden](https://www.nvict.nl/software/updates/)

## 📄 Changelog

### Versie 1.7.1 (Huidig - December 2025)
**UI Verbeteringen:**
- ✨ Scrollbar toegevoegd aan update dialoog voor lange release notes
- 🐛 Betere zichtbaarheid van alle update informatie

### Versie 1.7 (December 2025)
**Grote Print Update:**
- ✨ **Direct Printen via Windows GDI** - Volledig herschreven print functionaliteit
  - Print direct naar elke printer (USB, netwerk, virtueel)
  - Geen tijdelijke bestanden meer nodig
  - Universeel compatibel met alle Windows printers
  - Intelligente error handling met concrete oplossingen
  - Automatische standaard printer selectie
  
**UI Verbeteringen:**
- 🎨 Toolbar knoppen hebben nu tekstlabels (Opslaan, Zoeken, Kopiëren, Printen, Info)
- 🖨️ Print knop verplaatst naar logische positie na Opslaan
- 🐛 Printer dropdown weergave gecorrigeerd

**Gebruikerservaring:**
- 🔒 **Duplicate File Detection** - Hetzelfde bestand wordt niet meer dubbel geopend
  - Automatisch switchen naar bestaande tab
  - Efficiënter geheugengebruik
- 🪟 **Smart Window Management** - Programma komt automatisch naar voren bij openen bestand
  - Ook wanneer geminimaliseerd
  - Betere workflow

**Technisch:**
- ⚙️ Vereist `pywin32` library voor print functionaliteit
- 🔧 Stabielere codebase met betere error handling

### Versie 1.6 (November 2025)
- Modern UI design met dark/light theme
- Multi-tab ondersteuning
- Single instance functionaliteit
- Automatische update checker
- Verbeterde zoekfunctie
- Bladwijzers en aantekeningen
- Thumbnail preview
- Uitgebreide afdrukopties
- Windows registry integratie voor standaard PDF viewer

## 🔄 Update Proces

NVict Reader controleert automatisch op updates bij het opstarten. Wanneer een update beschikbaar is:

1. Een dialoog verschijnt met volledige release notes (scrollbaar vanaf v1.7.1)
2. Klik "Download & Installeer" voor automatische installatie
3. Of klik "Alleen Download" om handmatig te installeren
4. Sluit NVict Reader af voordat je installeert

## ⚙️ Technische Details

### Print Functionaliteit (v1.7+)

NVict Reader gebruikt Windows GDI voor directe printer toegang:
- Rendert PDF pagina's naar high-resolution images op printer DPI
- Gebruikt native Windows print API (dezelfde als Word, Excel, Chrome)
- Werkt met alle printer types zonder extra software
- **Vereist:** `pywin32` library (inbegrepen in installer)

### Dependencies

Belangrijkste libraries:
- `PyMuPDF` (fitz) - PDF rendering en manipulatie
- `tkinter` - GUI framework  
- `Pillow` (PIL) - Image processing
- `pywin32` - Windows API toegang voor print functionaliteit

Zie `requirements.txt` voor volledige lijst.

## 🐛 Bekende Issues

Geen bekende kritieke issues in v1.7.1.

Voor bug reports of feature requests, neem contact op via de website.

---

<div align="center">
Gemaakt met ❤️ door <a href="https://www.nvict.nl">NVict Service</a>
</div>
