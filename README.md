<p align="center">
  <img
    src="https://raw.githubusercontent.com/uzairshahidgithub/Libre-Merger/f87fa006bee5699c183b24862326045741389518/Libre%20Merger.png"
    alt="Libre Merger Logo"
    width="260"
    height='420'
  />
</p>

<h2 align="center">Libre Merger - Open-Source Automated Document Merger</h2>

<p align="center" style="max-width: 800px; margin: 0 auto;">
  A free, community-driven document tool that harnesses the full power of
  <strong>LibreOffice</strong> to automatically merge, split, convert, and watermark
  <strong>any document type</strong> you attach entirely offline, entirely free,
  with no subscriptions, no cloud uploads, and no proprietary lock-in.
  Built for developers, legal professionals, students, and anyone who works with documents.
</p>

<br />

<p align="center">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License" />
  <img src="https://img.shields.io/badge/LibreOffice-Powered-1C87D0.svg" alt="LibreOffice Powered" />
  <img src="https://img.shields.io/badge/Platform-Linux%20Ubuntu-E95420.svg" alt="Ubuntu Linux" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D4.svg" alt="Windows" />
  <img src="https://img.shields.io/badge/macOS-Coming%20Soon-lightgrey.svg" alt="macOS Coming Soon" />
  <img src="https://img.shields.io/badge/Status-Active-brightgreen.svg" alt="Active" />
  <img src="https://img.shields.io/badge/Python-3.x-blue.svg" alt="Python 3.x" />
</p>

<br />

## Key Highlights
* **Multi-format support** merge DOCX, ODT, PDF, XLSX, PPTX, TXT, RTF, and CSV in any combination
* **Watermark text & PDF password** stamp a custom watermark and lock your output PDF with a password before sharing
* **Flexible and manual sorting** auto-sort by name or date, or fully drag-and-customize document order with per-document page limits
* **Smart compression** reduce output file size with quality optimized for AI attachments and fast sharing
* **DOCX and PDF output** choose your output format and save results into organized folders automatically
* **Original PDF names preserved in outline** source filenames appear as bookmarks inside the merged document for easy navigation
* **Cross-platform** Native Linux Ubuntu `.deb` package & Windows installer supported
* **100% free** no subscriptions, no account, no hidden costs, forever

## Supported Formats
| Format | Extension | Support |
| --- | --- | --- |
| Word Document | `.docx` `.doc` | Full |
| OpenDocument Text | `.odt` | Full |
| PDF | `.pdf` | Full |
| Spreadsheet | `.xlsx` `.ods` `.csv` | Full |
| Presentation | `.pptx` `.odp` | Full |
| Plain Text | `.txt` | Full |
| Rich Text Format | `.rtf` | Full |
| HTML Document | `.html` `.htm` | Full |

## Quick Setup for New Users

### Option A: Linux (Ubuntu / Debian `.deb`)

#### 1. Install LibreOffice
LibreOffice powers document conversion in Libre Merger. Install it via apt:
```bash
sudo apt update
sudo apt install -y libreoffice
```

#### 2. Install the Libre Merger `.deb` Package
Locate `libre-merger_1.0.0_amd64.deb` in the `dist/` directory:
```bash
sudo apt install ./dist/libre-merger_1.0.0_amd64.deb
```
*(Or via dpkg: `sudo dpkg -i dist/libre-merger_1.0.0_amd64.deb`)*

#### 3. Launch Libre Merger
Launch from your application launcher under **Office** or via terminal:
```bash
libre-merger
```
You can also queue documents directly:
```bash
libre-merger doc1.pdf doc2.docx notes.txt
```

---

### Option B: Windows

#### 1. Install LibreOffice (Required)
Download the latest version for Windows from [libreoffice.org](https://www.libreoffice.org/download/download/) and install with default settings.

#### 2. Run the Libre Merger Installer
1. Navigate to the `dist/` directory
2. Double-click `Setup.exe` to launch the installation wizard
3. The wizard will create a Desktop shortcut and Start Menu entry

> **Windows SmartScreen Notice:** If Windows shows a security warning, click **"More info" → "Run anyway"**.

---

### Step 3: Start Merging
1. Open **Libre Merger** from your desktop or application menu
2. Drag and drop or click **＋ Add Documents** to import documents
3. Arrange the file order or set custom page ranges (e.g. `1-3, 5`)
4. Adjust optional watermark, password encryption, or compression settings
5. Click **🔀 Merge** to generate your combined document

## For Developers Build from Source

### Prerequisites
| Tool | Version | Download |
| --- | --- | --- |
| Python | 3.x | [python.org](https://www.python.org/downloads/) |
| LibreOffice | 7.x or higher | [libreoffice.org](https://www.libreoffice.org/download/) |
| pip | Latest | Included with Python |
| Git | Any | [git-scm.com](https://git-scm.com/) |

### Step 1: Clone the Repository
```bash
git clone https://github.com/uzairshahidgithub/Libre-Merger.git
cd "Libre Merger"
```

### Step 2: Set Up Virtual Environment & Dependencies
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt pytest
```

### Step 3: Run in Development Mode
```bash
python merger.py
```

### Step 4: Run Automated Tests
```bash
pytest tests/ -v
```

### Step 5: Build Package
Run the unified build script:
```bash
python build.py
```
* **On Linux (Ubuntu / Debian):** Automatically compiles the standalone binary and packages `dist/libre-merger_1.0.0_amd64.deb`.
* **On Windows:** Packages `dist/LibreMerger.exe` and `dist/Setup.exe`.

## Troubleshooting

**LibreOffice not detected:**
Open `config.json` and set the path manually:
```json
{
  "libreoffice_path": "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
}
```
On Linux / macOS:
```bash
which libreoffice
```

**`Setup.exe` blocked by Windows Defender:**
Click **More info → Run anyway**. The full source code is auditable in this repository.

**Merged PDF has missing fonts:**
Ensure all fonts used in source documents are installed on your system. LibreOffice uses locally installed fonts for rendering.

**Slow performance on large batch merges:**
Split documents into smaller groups, merge each group separately, then merge the outputs together.

## Compatibility
* Linux Ubuntu 20.04 / 22.04 / 24.04 / 26.04+ (Debian `.deb` package)
* Windows 10 / 11
* macOS — Coming Soon
* Python 3.8 — 3.14
* LibreOffice 7.x — 26.x
* No internet connection required

## Security
* 100% offline processing
* No cloud uploads of any kind
* No telemetry or analytics
* No account or login required
* Full open-source — entirely auditable
* Safe for confidential and enterprise documents

## ![License](https://img.shields.io/badge/License-MIT-green.svg) License
Free to use, modify, and distribute under the MIT License.
Ideal for **developers, students, legal professionals, SOC teams, and enterprise document workflows**.

## References
- [LibreOffice Official Documentation](https://documentation.libreoffice.org/)
- [LibreOffice CLI / soffice Parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)
- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [PyPDF2 Documentation](https://pypdf2.readthedocs.io/)
- [PyInstaller Documentation](https://pyinstaller.org/)

## ❤️ Credits
Developed by **[Muhammad Uzair](https://github.com/uzairshahidgithub)**
Dedicated to providing free, open-source productivity tools for everyone.

## Support
For support, email uzairrshahid@gmail.com or join our [Discord Community Codemo Teams](https://linktr.ee/codemoteams).
