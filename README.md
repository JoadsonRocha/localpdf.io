<div align="right">

🌐 **Language** | **English** · [Português 🇧🇷](README.pt-br.md)

</div>

<div align="center">

# 🌟 LocalPDF.io

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-3.1.3-green.svg)](https://flask.palletsprojects.com/)
[![pywebview](https://img.shields.io/badge/pywebview-6.0+-blueviolet.svg)](https://pywebview.flowrl.com/)
[![Docker](https://img.shields.io/badge/Docker-Ready-blue.svg)](https://www.docker.com/)
[![Live Demo](https://img.shields.io/badge/Demo%20Online-localpdf.up.railway.app-blueviolet.svg)](https://localpdf.up.railway.app/)
[![Windows Release](https://img.shields.io/badge/Download-LocalPDF.msi-0078D4.svg?logo=windows&logoColor=white)](https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![All Contributors](https://img.shields.io/github/all-contributors/virgiliojr94/localpdf.io?color=ee8449&style=flat-square)](#-contributors)

<br>

<a href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi">
  <img src="https://img.shields.io/badge/Download_for_Windows-LocalPDF.msi-0078D4?style=for-the-badge&logo=windows&logoColor=white" height="48" alt="Download LocalPDF for Windows">
</a>

<br><br>

> Every PDF tool you need — 100% local, 100% private, zero cloud tracking.

**💻 [Download Windows Installer (MSI)](https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi)** • **🚀 [Try Live Web Demo on Railway](https://localpdf.up.railway.app/)** • **🌐 [Official Project Website](https://virgiliojr94.github.io/localpdf.io/)**

[Download](#-windows-desktop--official-installer) •
[Features](#-features--ux) •
[Architecture](#-architecture) •
[Usage](#-usage) •
[Roadmap](ROADMAP.md) •
[Contributing](CONTRIBUTING.md) •
[License](#-license)

</div>

---

## 📋 What is it?

LocalPDF.io is a modern application for PDF and document manipulation. Every file is processed directly on your computer — nothing is sent to the cloud.

No accounts. No subscriptions. No artificial page limits. Complete document privacy.

---

## ✨ Features & UX

All operations run locally with full hardware performance and a 100 MB configurable file limit.

### 📥 Convert to PDF
- **🖼️ Images → PDF** - Combine multiple JPG and PNG images into a single PDF.
- **📝 Word → PDF** - Convert one or more DOCX documents into high-fidelity PDF with preserved typography, headings (Title, H1, H2), inline styling (**bold**, *italic*, <u>underline</u>), and structured tables with auto-wrapping (powered by ReportLab `SimpleDocTemplate`).
- **📊 Excel → PDF** - Convert XLSX spreadsheets into neatly formatted PDF tables with auto-scaled columns, proper date formatting, and formula evaluation.
- **📄 Text → PDF** - Convert plain TXT files into clean, readable PDF pages.

### 📤 Convert from PDF
- **🖼️ PDF → JPG** - Convert PDF pages into compressed JPG images.
- **🖼️ PDF → PNG** - Convert PDF pages into high-definition PNG images.
- **🖼️ PDF → Images** - Extract every PDF page into separate image files.
- **📊 PDF → Excel** - Extract tables from PDF directly into editable Excel (.xlsx) spreadsheets (powered by `pdfplumber` and `openpyxl`).
- **📝 PDF → Word** - Convert PDF files into fully editable Word (.docx) documents (powered by `pdf2docx`).
- **📄 PDF → Text** - Extract selectable text into plain TXT files.
- **🔒 PDF → PDF/A** - Standardize documents to the archival PDF/A-1b standard (via Ghostscript with dynamic DLL discovery).
- **🔍 OCR PDF** - Optical Character Recognition for scanned PDFs and images using local Tesseract OCR, with smart fallback to embedded native text when OCR is not installed.

### 🔄 Organize, Edit and Secure PDF
- **🔗 Merge PDFs** - Combine multiple PDF files in any order.
- **✂️ Split PDF** - Extract all pages or specific page intervals (e.g. `1-3, 5, 8-10`).
- **📦 Compress PDF** - Optimize images and reduce file size while preserving readability.
- **🔐 Protect PDF** - Encrypt PDF documents with AES-256 local passwords.
- **🔓 Unlock PDF** - Permanently remove password protection from protected PDF documents.
- **💧 Professional Watermark** - Apply subtle, 45° diagonal semi-transparent watermarks (or header/footer) with custom color and opacity controls.
- **🔢 Page numbers** - Add custom page numbering to headers or footers.
- **🖥️ Visual PDF Page Editor** - Interactive canvas to reorder by drag-and-drop, insert pages from external PDFs/images, rotate, duplicate, add blank pages, and delete pages.

### 💻 Native Windows Desktop Experience
- **🪟 Native PyWebView Window**: Runs in a lightweight, frameless native desktop window powered by Microsoft Edge Chromium (WebView2) with no browser chrome distractions.
- **💾 Native Save File Dialog**: Integrated Windows `SaveFileDialog` allowing direct selection of output destination folder and filename with extension filters.
- **📂 Explorer & System Integration**: "Open Folder" (`explorer.exe /select`) and "Open File" (`os.startfile`) buttons for instantaneous access to generated documents.
- **⚙️ Auto-Detection of Dependencies**: Automatically locates Tesseract OCR and Ghostscript installations in `vendor/` or standard Windows directories (`%LOCALAPPDATA%`, `Program Files`) on startup and sets required environment paths.
- **🛡️ Single-Instance Coordination**: Windows Mutex (`CreateMutexW`) and active port file coordinate to prevent duplicate background processes, automatically restoring and focusing the running native window when launched again.
- **🌐 Seamless Browser Fallback**: Gracefully falls back to the system's default web browser if PyWebView runtime is not detected.

### 💎 Modern User Interface & Flow
- **🎨 Custom SVG Icons**: Dedicated vector icon for every tool instead of generic emojis.
- **🏷️ Real-time Category Tabs**: Instant filtering by *All*, *Organize PDF*, *Convert PDF*, *Optimize PDF*, and *OCR*.
- **📑 Direct Tool Routing**: Dedicated clean URLs (`/tool/<tool_name>`, `/editor`, and `/about`) with synchronized browser titles and rich SEO metadata.
- **⚡ Dynamic Loading & Stage Feedback**: Animated progress bar (`progressShimmer`), real-time contextual stage messages (e.g. OCR text recognition, image compression), and elapsed timer (`⏱️ 00:03`).
- **📁 File Management**: File extension badges, human-readable file sizes (`KB` / `MB`), and individual removal buttons (`✕ Remove`).
- **📥 One-Click Re-Download**: Result card retains downloaded files in memory, allowing instant re-download without reprocessing.
- **🏷️ Standardized File Renaming**: Generated files strictly retain your original document's base name (e.g. `report_compressed.pdf`, `contract_merged.pdf`, `financial.xlsx`, `document_ocr.txt`).
- **ℹ️ Dedicated About Page**: Built-in `/about` page detailing zero-cloud architecture, comparative benchmarks against online services, and open source documentation.
- **🌐 Full Bilingual Support**: Toggle seamlessly between Portuguese (pt-BR) and English (EN).

---

## 🏗️ Architecture

LocalPDF.io is built around a dual-target architecture: a native Windows Desktop application (Portable / MSI with PyWebView) and a web service (Docker / Cloud).

```mermaid
graph TD
    User([User]) --> InterfaceChoice{Client Mode}
    
    subgraph Desktop Experience
        InterfaceChoice -->|Native Desktop App| PyWebView[Native PyWebView Window - EdgeChromium]
        PyWebView --> DesktopAPI[DesktopAPI: Native File Dialogs & Explorer Actions]
        DesktopAPI --> WinLauncher[windows_launcher.py]
        WinLauncher --> Mutex[Single-Instance Mutex & Window Restore]
        WinLauncher --> NativeDep[Auto-Detect Tesseract OCR & Ghostscript]
        WinLauncher --> WaitressWSGI[Waitress WSGI on 127.0.0.1]
    end

    subgraph Web & Cloud Experience
        InterfaceChoice -->|Web Browser| WebUI[Modern Responsive Web UI]
        WebUI --> FlaskApp[Flask Web Backend]
        FlaskApp --> Railway[Railway Cloud: Ephemeral Volatile Processing]
        FlaskApp --> DockerEnv[Docker Container / Local Server]
    end

    WaitressWSGI --> ProcessingEngines
    FlaskApp --> ProcessingEngines

    subgraph Local Processing Core
        ProcessingEngines[Engines: PyMuPDF / pdf2docx / pdfplumber / ReportLab / Tesseract OCR / Ghostscript]
    end

    subgraph Packaging & Distribution
        WinLauncher --> PyInstaller[PyInstaller Bundle: LocalPDF.exe]
        PyInstaller --> WiX[WiX Toolset v4: LocalPDF.msi Installer]
        WiX --> SignedRelease[Authenticode SHA-256 Code Signed]
    end
```

### 1. Windows Desktop Architecture
- **Native GUI Window**: Powered by `pywebview` with EdgeChromium backend.
- **Portable Executable**: Standalone bundle compiled with PyInstaller (`dist/LocalPDF/LocalPDF.exe`).
- **MSI Installer**: Enterprise-grade Windows Installer package generated with WiX Toolset v4 (`dist/LocalPDF.msi`) for per-machine installation in `Program Files`.
- **Local Isolation & Security**:
  - Automatically binds strictly to `127.0.0.1` on a dynamically allocated free port.
  - Protected by a Windows Mutex (`CreateMutexW`) to prevent duplicate or zombie processes.
  - Temporary processing files reside in `%TEMP%` — never writing to restricted directories like `C:\Program Files`.
  - Desktop-only endpoints (`/api/desktop/*`) allow native file saving, "Show in Folder", and "Open File".
- **Authenticode Code Signing**:
  - Scripts provided in `packaging/windows/` for signing binaries and installers with SHA256 and RFC 3161 timestamps (DigiCert).
  - `install-cert.ps1` allows one-click local trust configuration to prevent Windows SmartScreen unknown root warnings.

### 2. Web & Cloud Architecture
- **Backend Framework**: Python 3.11+ with Flask 3.1.3 and Waitress WSGI.
- **Processing Core**: PyMuPDF (MuPDF C++ bindings) for high-speed PDF rendering, pdf2docx, Pillow, ReportLab, and openpyxl.
- **Cloud Demo**: Deployed on Railway at [https://localpdf.up.railway.app/](https://localpdf.up.railway.app/).
- **Ephemeral Processing**: Web mode runs strictly in volatile memory with instant post-download file purging and hardened security headers (`X-Content-Type-Options: nosniff`, `Cache-Control: no-store`).
- **Containerization**: Official Docker image available on GitHub Container Registry.

---

## 💻 Windows Desktop & Official Installer

Get the official, digitally signed Windows installer:

<div align="center">
  <br>
  <a href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi">
    <img src="https://img.shields.io/badge/Download_for_Windows-LocalPDF.msi-0078D4?style=for-the-badge&logo=windows&logoColor=white" height="48" alt="Download Windows Installer">
  </a>
  <p><em>Compatible with Windows 10 & 11 (64-bit) • ~99 MB • Authenticode Signed • Start Menu & Desktop Shortcuts</em></p>
  <br>
</div>

- **Native System Installation:** Installs into `C:\Program Files\LocalPDF.io` with clean uninstall support.
- **Offline & Private:** Processed strictly on `127.0.0.1` — your files never leave your machine.
- **Desktop Features:** Native save dialog, "Open Folder" in Windows Explorer, and "Open File" in default app.
- **Single Instance:** Automatically restores the existing window without creating redundant background processes.

> 💡 **Developers**: See [packaging/windows/README.md](packaging/windows/README.md) for full instructions on building the package, compiling the WiX MSI installer, and signing releases with Authenticode.

---

## 🚀 Usage

### 1. Online Demo (Instant Access)
Try all features directly in your browser:
👉 **[https://localpdf.up.railway.app/](https://localpdf.up.railway.app/)**

### 2. Docker (Fastest Self-Hosted)
Run directly from the GitHub Container Registry:

```bash
docker run -p 5000:5000 ghcr.io/virgiliojr94/localpdf.io:latest
```

Open: **http://localhost:5000**

### 3. Run from Source

```bash
# Clone the repository
git clone https://github.com/JoadsonRocha/localpdf.io.git
cd localpdf.io

# Setup Python virtual environment
python -m venv .venv
.\.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run server
python app.py
```

Open: **http://localhost:5000**

---

## 🛠️ Tech Stack

- **Flask 3.1.3** - Core Python web application framework
- **pywebview 6.0+** - Native desktop window shell (Edge Chromium)
- **PyMuPDF 1.23.8 (Fitz)** - High performance C++ PDF rendering and manipulation engine
- **Pillow 12.1.1 (PIL)** - Image manipulation and rasterization
- **pdf2docx 0.5.8** - PDF to Word conversion engine
- **python-docx 0.8.11** - Word document reading and structuring
- **pdfplumber 0.11.9** - High-precision PDF table and layout extraction
- **ReportLab 4.0.5** - Dynamic PDF document generation (`SimpleDocTemplate`, `Table`, `Paragraph`)
- **OpenPyXL 3.1.2** - Excel spreadsheet parser and builder
- **Tesseract OCR (pytesseract 0.3.10)** - Optical Character Recognition engine with native text fallback
- **Ghostscript 0.7** - PDF/A archival compliance and color rendering
- **Waitress** - Production-grade WSGI server for local desktop loopback
- **WiX Toolset v4** - Windows Installer XML (MSI) build system
- **PyInstaller** - Standalone Windows executable packager with embedded vendor support

---

## 🔒 Privacy & Security

All document processing occurs strictly **on your local machine** when using the desktop application or self-hosted Docker. No document content, metadata, or telemetry is ever transmitted to external servers.

---

## 🤝 Contributing

Contributions are welcome! Please check out the [contribution guide](CONTRIBUTING.md) to get started.

---

## 👥 Contributors

Thank you to everyone who has contributed to LocalPDF.io! ✨

<a href="https://github.com/virgiliojr94/localpdf.io/graphs/contributors">
  <img alt="Repository contributors" src="https://contrib.rocks/image?repo=virgiliojr94/localpdf.io" />
</a>

---

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.
