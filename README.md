<div align="right">

🌐 **Language** | **English** · [Português 🇧🇷](README.pt-br.md)

</div>

<div align="center">

# 🌟 LocalPDF.io

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-2.3.3-green.svg)](https://flask.palletsprojects.com/)
[![Docker](https://img.shields.io/badge/Docker-Ready-blue.svg)](https://www.docker.com/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![All Contributors](https://img.shields.io/github/all-contributors/virgiliojr94/localpdf.io?color=ee8449&style=flat-square)](#-contributors)
[![Website](https://img.shields.io/badge/Website-Online-success.svg)](https://virgiliojr94.github.io/localpdf.io/)

> Every PDF tool you need — 100% local, 100% private.

**🌐 [Visit the Official Website](https://virgiliojr94.github.io/localpdf.io/)**

[Features](#-features) •
[Usage](#-usage) •
[Roadmap](ROADMAP.md) •
[Contributing](CONTRIBUTING.md) •
[License](#-license)

</div>

---

## 📋 What is it?

LocalPDF.io is a self-hosted web app for PDF and document manipulation. Every file is processed on your own machine — nothing is uploaded to any server.

No accounts. No cloud. No data leaving your computer.

## ✨ Features

All operations are available without accounts or artificial premium limits when running locally. Limits are determined by the computer's available resources and the configured 100 MB upload limit.

### 📥 Convert to PDF
- **🖼️ Images → PDF** - Combine JPG and PNG images into one PDF
- **📝 Word → PDF** - Convert one or more DOCX files into PDF
- **📊 Excel → PDF** - Convert XLSX spreadsheets into PDF
- **📄 Text → PDF** - Convert TXT files into formatted PDF

### 📤 Convert from PDF
- **🖼️ PDF → Images** - Export every page as a PNG image
- **📝 PDF → Word** - Convert PDF into an editable DOCX document
- **📄 PDF → Text** - Extract selectable text into a TXT file
- **🔒 PDF → PDF/A** - Convert to PDF/A-1b when Ghostscript is installed
- **🔍 OCR PDF** - Extract text from scanned PDFs and images with local Tesseract OCR

### 🔄 Organize and secure PDF
- **🔗 Merge PDFs** - Combine multiple PDF files
- **✂️ Split PDF** - Export individual pages
- **📦 Compress PDF** - Reduce file size locally
- **🔐 Protect PDF** - Encrypt a PDF with an AES-256 password
- **💧 Watermark PDF** - Add configurable text watermarks
- **🔢 Page numbers** - Add header or footer numbering

### 🖥️ Visual PDF page editor
- Reorder pages by drag and drop.
- Insert selected pages from another PDF or images.
- Add blank pages, duplicate, rotate and delete pages.
- Undo and redo changes before exporting.
- Preserve the original file and download a new PDF.

### 🌐 Interface
- Portuguese by default with an English interface option.
- Responsive layout for desktop and mobile.
- No account, login, analytics or AI features.

`PDF → Excel`, free-text editing of existing PDF content and password removal are not implemented yet. Their status is tracked in [ROADMAP.md](ROADMAP.md).

## 🚀 Usage

### With Docker (Recommended)

#### Pull and run (fastest)

```bash
docker run -p 5000:5000 ghcr.io/virgiliojr94/localpdf.io:latest
```

#### Build locally

```bash
git clone https://github.com/virgiliojr94/localpdf.io.git
cd localpdf.io
docker build -t localpdf .
docker run -p 5000:5000 localpdf
```

Open: **http://localhost:5000**

### Local MSI (planned)

The future Windows MSI will package the same local processing base without Railway, accounts or external uploads. The MSI is not published yet; see [ROADMAP.md](ROADMAP.md) for the packaging plan.

### Deploy on Railway

The application is compatible with Railway and reads the platform-provided `PORT` variable automatically.

1. Create a new Railway project from this repository.
2. Configure the service to build with the repository `dockerfile`.
3. Deploy and generate a public domain in Railway.

Unlike local or private Docker usage, files sent to a public Railway deployment are processed inside Railway infrastructure. Do not use it for sensitive documents unless the deployment, access control and data-retention policy have been reviewed.

### Without Docker

```bash
git clone https://github.com/virgiliojr94/localpdf.io.git
cd localpdf.io

pip install -r requirements.txt
# Instale o Ghostscript e Tesseract no sistema
# Debian/Ubuntu:
apt-get install ghostscript tesseract-ocr tesseract-ocr-por tesseract-ocr-eng

python app.py
```

Open: **http://localhost:5000**

## 🛠️ Tech stack

- **Flask** - Framework web Python
- **PyMuPDF** - Manipulação de PDFs
- **Pillow** - Processamento de imagens
- **python-docx** - Manipulação de arquivos Word
- **ReportLab** - Geração de PDFs
- **OpenPyXL** - Manipulação de planilhas Excel
- **PDF2Docx** - Conversor de PDF para Docx
- **Tesseract OCR** - Reconhecimento óptico de caracteres

## 🔒 Privacy

All files are processed **locally** on your machine. No data is sent to external servers — ever.

## 🤝 Contributing

Contributions are welcome! See the [contribution guide](CONTRIBUTING.md) to get started.

## 👥 Contributors

Thanks to everyone who has contributed to this project! ✨

<a href="https://github.com/virgiliojr94/localpdf.io/graphs/contributors">
  <img alt="Repository contributors" src="https://contrib.rocks/image?repo=virgiliojr94/localpdf.io" />
</a>

## 📝 License

MIT License — free to use and modify.

---

⭐ If this project was useful to you, consider giving it a star on GitHub!

## :star2: Star History

[![Star History Chart](https://api.star-history.com/svg?repos=virgiliojr94/localpdf.io&type=timeline&legend=top-left)](https://www.star-history.com/#virgiliojr94/localpdf.io&type=timeline&legend=top-left)
