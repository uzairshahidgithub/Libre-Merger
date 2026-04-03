<p align="center">
  <img
    src="https://raw.githubusercontent.com/uzairshahidgithub/Libre-Merger/f87fa006bee5699c183b24862326045741389518/Libre%20Merger.png"
    alt="Libre Merger Logo"
    width="460"
    height='360'
  />
</p>

<h2 align="center">Libre Merger - Open-Source Automated Document Merger</h2>

<p align="center" style="max-width: 800px; margin: 0 auto;">
  A free, community-driven document tool that harnesses the full power of
  <strong>LibreOffice</strong> to automatically merge, split, convert, and watermark
  <strong>any document type</strong> you attach — entirely offline, entirely free,
  with no subscriptions, no cloud uploads, and no proprietary lock-in.
  Built for developers, legal professionals, students, and anyone who works with documents.
</p>

<br />

<p align="center">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License" />
  <img src="https://img.shields.io/badge/LibreOffice-Powered-1C87D0.svg" alt="LibreOffice Powered" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D4.svg" alt="Windows" />
  <img src="https://img.shields.io/badge/macOS-Coming%20Soon-lightgrey.svg" alt="macOS Coming Soon" />
  <img src="https://img.shields.io/badge/Linux-Coming%20Soon-lightgrey.svg" alt="Linux Coming Soon" />
  <img src="https://img.shields.io/badge/Status-Active-brightgreen.svg" alt="Active" />
  <img src="https://img.shields.io/badge/Python-3.x-blue.svg" alt="Python 3.x" />
</p>

<br />

## Key Highlights
* **Multi-format support** — merge DOCX, ODT, PDF, XLSX, PPTX, TXT, RTF, and CSV in any combination
* **Watermark text & PDF password** — stamp a custom watermark and lock your output PDF with a password before sharing
* **Flexible and manual sorting** — auto-sort by name or date, or fully drag-and-customize document order with per-document page limits
* **Smart compression** — reduce output file size with quality optimized for AI attachments and fast sharing
* **DOCX and PDF output** — choose your output format and save results into organized folders automatically
* **Original PDF names preserved in outline** — source filenames appear as bookmarks inside the merged document for easy navigation
* **Currently Windows only** — macOS and Linux support coming soon
* **100% free** — no subscriptions, no account, no hidden costs, forever

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

### Step 1: Install LibreOffice (Required)
Libre Merger uses LibreOffice to process Word, Excel, and PowerPoint files. It must be installed before running the app.

Download the latest stable version for your OS:
```
https://www.libreoffice.org/download/download/
```
Run the installer and follow the default on-screen instructions. No custom configuration needed.

Verify the installation by opening **LibreOffice Writer** from your applications. If it launches, you are ready.

### Step 2: Run the Libre Merger Installer
Once LibreOffice is confirmed working:

1. Navigate to the **`dist/`** folder in this project directory
2. Locate the file named **`Setup.exe`**
3. **Double-click `Setup.exe`** to launch the installation wizard
4. Follow the prompts — the installer will create:
   - A Desktop shortcut
   - A Start Menu entry

> **Windows SmartScreen Notice:** If Windows shows a security warning, click **"More info"** then **"Run anyway"**. This is standard behavior for open-source unsigned installers.

### Step 3: Start Merging
1. Open **Libre Merger** from your Desktop or Start Menu
2. Drag and drop your documents into the application window
3. Arrange the file order if needed
4. Click **Merge** to generate your combined document
5. Choose your output location and format

## For Developers — Build from Source

### Prerequisites
| Tool | Version | Download |
| --- | --- | --- |
| Python | 3.x | [python.org](https://www.python.org/downloads/) |
| LibreOffice | 7.x or higher | [libreoffice.org](https://www.libreoffice.org/download/) |
| pip | Latest | Included with Python |
| Git | Any | [git-scm.com](https://git-scm.com/) |

### Step 1: Clone the Repository
```bash
git clone https://github.com/uzairshahidgithub/libre-merger.git
cd libre-merger
```

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Run in Development Mode
```bash
python main.py
```

### Step 4: Build the Installer
Run the automated build script to generate `Setup.exe` inside `dist/`:
```bash
python build.py
```

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
* Windows 10 / 11
* macOS — Coming Soon
* Linux — Coming Soon
* Python 3.8 — 3.12
* LibreOffice 7.x — 24.x
* No WSL required
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
