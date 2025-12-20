# NVict Reader

<div align="center">

![NVict Reader Logo](https://www.nvict.nl/images/reader_logo.png)

**Professionele PDF Reader voor Windows**

[![Version](https://img.shields.io/badge/version-1.6-blue.svg)](https://github.com/nvict-service/nvict-reader/releases)
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
- **Afdrukken** - Uitgebreide afdrukopties
- **Standaard PDF Viewer** - Stel in als standaard PDF programma
- **Automatische Updates** - Blijf altijd up-to-date
- **Single Instance** - Nieuw geopende bestanden worden automatisch in bestaand venster geladen
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

## 📝 Licentie

Copyright © 2024 NVict Service. Alle rechten voorbehouden.

Deze software is eigendom van NVict Service en mag niet worden gedistribueerd, gewijzigd of gebruikt zonder expliciete toestemming.

## 🔒 Code Signing

Releases worden automatisch ondertekend via GitHub Actions met een geldig code signing certificaat voor Windows Authenticode.

## 📞 Contact & Support

- **Website:** [www.nvict.nl](https://www.nvict.nl)
- **Support:** Via website contact formulier
- **Updates:** Automatisch via de applicatie of [handmatig downloaden](https://www.nvict.nl/software/updates/)

## 🔄 Changelog

### Versie 1.6 (Huidig)
- Modern UI design met dark/light theme
- Multi-tab ondersteuning
- Single instance functionaliteit
- Automatische update checker
- Verbeterde zoekfunctie
- Bladwijzers en aantekeningen
- Thumbnail preview
- Uitgebreide afdrukopties
- Windows registry integratie voor standaard PDF viewer

---

<div align="center">
Gemaakt met ❤️ door <a href="https://www.nvict.nl">NVict Service</a>
</div>
