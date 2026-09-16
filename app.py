import base64
import io
import json
import os
import shutil
import tempfile
import zipfile

import fitz  # PyMuPDF
import openpyxl
from flask import (
    Flask,
    jsonify,
    render_template_string,
    request,
    send_file,
)
from pdf2docx import Converter
from pdf2docx.converter import ConversionException
import pytesseract
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from werkzeug.utils import secure_filename

try:
    import ghostscript
except (ImportError, RuntimeError):
    ghostscript = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["OUTPUT_FOLDER"] = "outputs"

# Criar diretórios se não existirem
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "docx", "txt", "xlsx", "jpg", "jpeg", "png"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# Template HTML
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="icon" type="image/svg+xml" href="/favicon.svg">
    <title>LocalPDF.io</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: linear-gradient(135deg, #dbeafe 0%, #eff6ff 100%); min-height: 100vh; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; color: white; margin-bottom: 40px; }
        .header h1 { font-size: 3em; margin-bottom: 10px; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .tools-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 40px; }
        .tool-card { background: white; border-radius: 15px; padding: 30px; text-align: center; box-shadow: 0 10px 30px rgba(0,0,0,0.1); transition: transform 0.3s ease; cursor: pointer; }
        .tool-card:hover { transform: translateY(-5px); }
        .tool-card h3 { color: #333; margin-bottom: 15px; font-size: 1.5em; }
        .tool-card p { color: #666; margin-bottom: 20px; }
        .upload-area { border: 2px dashed #ddd; border-radius: 10px; padding: 40px; text-align: center; background: #f9f9f9; margin: 20px 0; transition: all 0.3s ease; }
        .upload-area:hover { border-color: #2563eb; background: #eff6ff; }
        .upload-area.dragover { border-color: #2563eb; background: #dbeafe; }
        .file-input { display: none; }
        .upload-btn { background: #2563eb; color: white; padding: 12px 30px; border: none; border-radius: 25px; cursor: pointer; font-size: 1.1em; transition: background 0.3s ease; }
        .upload-btn:hover { background: #1d4ed8; }
        .convert-btn { background: #2563eb; color: white; padding: 15px 40px; border: none; border-radius: 25px; cursor: pointer; font-size: 1.2em; margin-top: 20px; transition: background 0.3s ease; }
        .convert-btn:hover { background: #1d4ed8; }
        .convert-btn:disabled { background: #ccc; cursor: not-allowed; }
        .file-list { margin-top: 20px; }
        .file-item { background: #f8f9fa; padding: 10px 15px; margin: 5px 0; border-radius: 5px; display: flex; justify-content: space-between; align-items: center; }
        .progress { width: 100%; background: #f0f0f0; border-radius: 10px; margin: 20px 0; }
        .progress-bar { height: 20px; background: #2563eb; border-radius: 10px; width: 0%; transition: width 0.3s ease; }
        .result { margin-top: 20px; padding: 20px; background: #d4edda; border-radius: 10px; color: #155724; }
        .error { margin-top: 20px; padding: 20px; background: #f8d7da; border-radius: 10px; color: #721c24; }
        .hidden { display: none; }
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); }
        .modal-content { background: white; margin: 5% auto; padding: 30px; width: 80%; max-width: 600px; border-radius: 15px; position: relative; }
        .close { position: absolute; right: 20px; top: 15px; font-size: 30px; cursor: pointer; color: #aaa; }
        .close:hover { color: #000; }
        .back-btn { background: #6c757d; color: white; padding: 10px 20px; border: none; border-radius: 25px; cursor: pointer; margin-bottom: 20px; }
        .back-btn:hover { background: #545b62; }
        .footer { text-align: center; color: white; margin-top: 40px; padding: 20px 0; border-top: 1px solid #ddd; }
        .footer p { margin-bottom: 10px; }
        .footer a { color: #2563eb; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }
        .social-icons { margin-top: 10px; }
        .social-icons a { margin: 0 10px; color: #2563eb; font-size: 1.2em; }
        .editor-shell { background: #f8f9fa; border-radius: 15px; padding: 24px; margin-top: 20px; }
        .editor-toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-bottom: 20px; }
        .editor-toolbar button, .editor-actions button { border: 0; border-radius: 8px; padding: 10px 14px; cursor: pointer; font-weight: 600; }
        .editor-toolbar button { background: #e9ecef; color: #343a40; }
        .editor-toolbar button:hover { background: #dee2e6; }
        .editor-toolbar button:disabled { opacity: 0.5; cursor: not-allowed; }
        .editor-pages { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 18px; min-height: 180px; }
        .editor-page { position: relative; background: white; border: 2px solid transparent; border-radius: 10px; padding: 10px; box-shadow: 0 3px 12px rgba(0,0,0,0.08); cursor: grab; }
        .editor-page.selected { border-color: #2563eb; }
        .editor-page.dragging { opacity: 0.45; }
        .editor-page img { display: block; width: 100%; aspect-ratio: 0.72; object-fit: contain; background: #e9ecef; border-radius: 5px; }
        .editor-page-number { font-weight: 700; color: #495057; margin: 8px 0; }
        .editor-page-actions { display: flex; gap: 5px; flex-wrap: wrap; }
        .editor-page-actions button { flex: 1; min-width: 42px; padding: 7px 5px; border: 0; border-radius: 6px; background: #edf0f2; cursor: pointer; }
        .editor-page-actions button:hover { background: #dfe4e8; }
        .editor-empty { color: #6c757d; text-align: center; padding: 50px 20px; border: 2px dashed #ced4da; border-radius: 10px; grid-column: 1 / -1; }
        .editor-actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 20px; align-items: center; }
        .editor-actions .primary { background: #2563eb; color: white; }
        .editor-actions .secondary { background: #1d4ed8; color: white; }
        .editor-status { color: #6c757d; font-size: 0.95em; }
        .editor-file-input { display: none; }
        @media (max-width: 600px) { .editor-pages { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; } .editor-shell { padding: 14px; } }
        body { background: #f6f7f9; color: #24272b; }
        .container { max-width: 1180px; padding: 0 28px; }
        .site-nav { display: flex; align-items: center; justify-content: space-between; padding: 22px 0; border-bottom: 1px solid #e4e6e9; }
        .site-brand { color: #24272b; font-size: 1.35rem; font-weight: 800; letter-spacing: -0.03em; text-decoration: none; }
        .site-brand span { color: #2563eb; }
        .site-brand small { font-size: 0.8em; }
        .site-nav-links { display: flex; align-items: center; gap: 22px; }
        .site-nav-links a { color: #5f6368; text-decoration: none; font-size: 0.92rem; font-weight: 600; }
        .site-nav-links a:hover { color: #1d4ed8; }
        .language-toggle { border: 1px solid #bfdbfe; border-radius: 999px; background: #fff; color: #1d4ed8; padding: 7px 11px; font: inherit; font-size: 0.82rem; font-weight: 800; cursor: pointer; }
        .language-toggle:hover { background: #eff6ff; }
        .privacy-pill { color: #1d4ed8 !important; background: #dbeafe; border-radius: 999px; padding: 8px 13px; }
        .header { color: #24272b; margin: 0 auto; padding: 38px 0 24px; max-width: 700px; }
        .header h1 { font-size: clamp(2rem, 4vw, 3.2rem); line-height: 1.05; letter-spacing: -0.045em; margin-bottom: 10px; }
        .header p { color: #6c7178; font-size: 1rem; line-height: 1.5; }
        .home-eyebrow { display: inline-block; color: #1d4ed8; background: #dbeafe; border-radius: 999px; padding: 6px 11px; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 11px; }
        .category-tabs { display: flex; flex-wrap: wrap; justify-content: center; gap: 8px; margin-bottom: 28px; }
        .category-tabs span { background: #fff; border: 1px solid #e3e5e8; border-radius: 999px; color: #656a70; padding: 8px 14px; font-size: 0.86rem; font-weight: 700; }
        .category-tabs span:first-child { background: #24272b; color: #fff; border-color: #24272b; }
        .tools-grid { grid-template-columns: repeat(auto-fit, minmax(245px, 1fr)); gap: 14px; margin-bottom: 58px; }
        .tool-card { border: 1px solid #e3e5e8; border-radius: 10px; padding: 22px; text-align: left; box-shadow: 0 5px 18px rgba(36,39,43,0.04); transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease; }
        .tool-card { position: relative; overflow: hidden; min-height: 142px; background: rgba(255,255,255,0.96); }
        .tool-card::before { content: ''; position: absolute; inset: 0 auto 0 0; width: 3px; background: #dbeafe; transition: background 0.2s ease; }
        .tool-card:hover { transform: translateY(-3px); border-color: #93c5fd; box-shadow: 0 12px 28px rgba(37,99,235,0.12); }
        .tool-card:hover::before { background: #2563eb; }
        .tool-card h3 { color: #24272b; font-size: 1.05rem; margin-bottom: 8px; }
        .tool-card p { color: #747980; font-size: 0.9rem; line-height: 1.5; margin-bottom: 0; }
        .tools-grid .tool-card:last-child { border-color: #2563eb; box-shadow: 0 8px 24px rgba(37,99,235,0.14); }
        .footer { color: #747980; border-top: 1px solid #e3e5e8; }
        .footer a { color: #2563eb; }
        .footer .social-icons a { color: #747980; }
        .footer-grid { display: grid; grid-template-columns: 1.4fr 1fr 1fr; gap: 28px; max-width: 900px; margin: 0 auto 24px; text-align: left; }
        .footer-block h4 { color: #24272b; margin-bottom: 8px; font-family: Georgia, "Times New Roman", serif; }
        .footer-block p, .footer-block a { font-size: 0.88rem; line-height: 1.7; }
        .footer-block a { display: block; }
        .footer-credit { border-top: 1px solid #e3e5e8; padding-top: 18px; }
        body { font-family: "Avenir Next", "Segoe UI", sans-serif; }
        .header h1, .tool-card h3, .editor-page-number { font-family: Georgia, "Times New Roman", serif; }
        .editor-shell { background: #f5f9ff; border: 1px solid #bfdbfe; }
        .editor-toolbar button, .editor-actions button { border: 1px solid #eadbd8; background: #fff; color: #3b3534; }
        .editor-toolbar button:hover, .editor-actions button:hover { border-color: #2563eb; color: #1d4ed8; background: #fff; }
        .editor-page.selected { border-color: #2563eb; box-shadow: 0 0 0 3px #dbeafe; }
        .editor-insert-panel { background: #fff; border: 1px solid #bfdbfe; border-radius: 10px; padding: 16px; margin-bottom: 18px; }
        .editor-insert-panel h4 { color: #24272b; margin-bottom: 5px; }
        .editor-insert-panel p { color: #747980; font-size: 0.9rem; margin-bottom: 12px; }
        .editor-pending-pages { display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 10px; max-height: 300px; overflow: auto; }
        .editor-pending-page { position: relative; border: 2px solid #e8e9eb; border-radius: 8px; padding: 5px; background: #fafafa; cursor: pointer; }
        .editor-pending-page.selected { border-color: #2563eb; background: #eff6ff; }
        .editor-pending-page img { display: block; width: 100%; aspect-ratio: 0.72; object-fit: contain; background: #eee; }
        .editor-pending-page label { display: flex; gap: 6px; align-items: center; padding: 6px 2px 2px; font-size: 0.78rem; color: #4c5157; }
        .editor-insert-actions { display: flex; gap: 8px; margin-top: 14px; }
        .editor-insert-actions button { border: 0; border-radius: 7px; padding: 9px 12px; cursor: pointer; font-weight: 700; }
        .editor-insert-actions .primary { background: #2563eb; color: #fff; }
        .editor-insert-actions .secondary { background: #f1eded; color: #4c4544; }
        #options { display: grid; gap: 7px; margin-top: 18px; text-align: left; }
        #options label { color: #3b3534; font-weight: 700; }
        #options input, #options select { width: 100%; border: 1px solid #d8dadd; border-radius: 7px; padding: 11px 12px; font: inherit; background: #fff; }
        #options small { color: #747980; }
        .upload-area { border-color: #bfdbfe; background: #fbfdff; box-shadow: inset 0 0 0 1px rgba(37,99,235,0.03); }
        .upload-area p:first-child { color: #3f4b5a; font-weight: 600; }
        .upload-btn, .convert-btn { border-radius: 9px; font-weight: 700; box-shadow: 0 5px 12px rgba(37,99,235,0.18); }
        .upload-btn:focus-visible, .convert-btn:focus-visible, .back-btn:focus-visible, button:focus-visible, a:focus-visible { outline: 3px solid #93c5fd; outline-offset: 3px; }
        .back-btn { background: #334155; border-radius: 9px; font-weight: 700; }
        .result, .error { border: 1px solid rgba(37,99,235,0.12); box-shadow: 0 8px 20px rgba(36,39,43,0.05); }
        .file-item { border: 1px solid #e5e7eb; background: #fff; }
        .editor-page, .editor-insert-panel { box-shadow: 0 8px 20px rgba(36,39,43,0.06); }
        @media (max-width: 700px) {
            .container { padding: 0 16px; }
            .site-nav { gap: 14px; padding: 15px 0; align-items: flex-start; }
            .site-brand { flex: 0 0 auto; font-size: 1.15rem; }
            .site-nav-links { min-width: 0; max-width: 72vw; gap: 14px; overflow-x: auto; padding-bottom: 4px; scrollbar-width: none; }
            .site-nav-links::-webkit-scrollbar { display: none; }
            .site-nav-links a { flex: 0 0 auto; font-size: 0.78rem; white-space: nowrap; }
            .privacy-pill { padding: 7px 10px; }
            .header { padding: 28px 0 20px; }
            .header h1 { font-size: 2rem; letter-spacing: -0.035em; }
            .header p { font-size: 0.94rem; line-height: 1.45; }
            .home-eyebrow { font-size: 0.7rem; }
            .category-tabs { justify-content: flex-start; overflow-x: auto; flex-wrap: nowrap; margin: 0 -16px 22px; padding: 0 16px 5px; scrollbar-width: none; }
            .category-tabs::-webkit-scrollbar { display: none; }
            .category-tabs span { flex: 0 0 auto; font-size: 0.78rem; }
            .tools-grid { grid-template-columns: 1fr; gap: 10px; margin-bottom: 38px; }
            .tool-card { padding: 18px; }
            .editor-shell { padding: 12px; }
            .editor-toolbar { gap: 8px; }
            .editor-toolbar button, .editor-actions button { flex: 1 1 100%; min-height: 42px; }
            .editor-pages { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 9px; }
            .editor-page { padding: 7px; }
            .editor-page-actions button { min-width: 0; padding: 8px 3px; }
            .editor-pending-pages { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .upload-area { padding: 26px 14px; }
            .footer { margin-top: 28px; }
            .footer a { display: inline-block; margin: 4px 0; }
            .footer-grid { grid-template-columns: 1fr; gap: 18px; text-align: center; }
        }
    </style>
</head>
<body>
    <div class="container">
        <nav class="site-nav">
            <a class="site-brand" href="#" onclick="showHome(); return false;">local<span>pdf</span><small>.io</small></a>
            <div class="site-nav-links">
                <a href="#tools">Juntar PDF</a>
                <a href="#tools">Dividir PDF</a>
                <a href="#tools">Comprimir PDF</a>
                <a href="#tools">Converter PDF</a>
                <a href="#tools">Todas as ferramentas</a>
                <a class="privacy-pill" href="#privacy-note">100% local</a>
                <button id="language-toggle" class="language-toggle" type="button" onclick="toggleLanguage()">EN</button>
            </div>
        </nav>
        <div class="header">
            <span class="home-eyebrow">PDF simples, privado e local</span>
            <h1>Trabalhe com seus PDFs sem complicação</h1>
            <p>Converta, organize e edite documentos diretamente no seu computador. Sem contas, sem nuvem e sem enviar seus arquivos para fora.</p>
        </div>

        <div id="home-view">
            <div id="tools" class="category-tabs" aria-label="Categorias de ferramentas">
                <span>Todas</span><span>Organizar PDF</span><span>Converter PDF</span><span>Otimizar PDF</span><span>OCR</span>
            </div>
            <div class="tools-grid">
                <div class="tool-card" onclick="showTool('pdf-to-images')">
                    <h3>🖼️ PDF para Imagens</h3>
                    <p>Converta páginas PDF em imagens JPG ou PNG</p>
                </div>
                <div class="tool-card" onclick="showTool('images-to-pdf')">
                    <h3>📄 Imagens para PDF</h3>
                    <p>Combine várias imagens em um único PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('merge-pdf')">
                    <h3>🔗 Mesclar PDFs</h3>
                    <p>Combine vários PDFs em um documento único</p>
                </div>
                <div class="tool-card" onclick="showTool('split-pdf')">
                    <h3>✂️ Dividir PDF</h3>
                    <p>Extraia páginas específicas do seu PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('compress-pdf')">
                    <h3>📦 Comprimir PDF</h3>
                    <p>Reduza o tamanho do seu arquivo PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('protect-pdf')">
                    <h3>🔐 Proteger PDF</h3>
                    <p>Adicione uma senha local ao seu documento PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('watermark-pdf')">
                    <h3>💧 Marca d'água</h3>
                    <p>Adicione uma marca d'água de texto ao PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('page-numbers-pdf')">
                    <h3>🔢 Números de página</h3>
                    <p>Numere as páginas do documento localmente</p>
                </div>
                <div class="tool-card" onclick="showTool('pdf-to-pdfa')">
                    <h3>🔒 PDF para PDF/A</h3>
                    <p>Padronize seu PDF para arquivamento (PDF/A)</p>
                </div>
                <div class="tool-card" onclick="showTool('word-to-pdf')">
                    <h3>📝 Word para PDF</h3>
                    <p>Converta um ou mais documentos DOCX para PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('excel-to-pdf')">
                    <h3>📊 Excel para PDF</h3>
                    <p>Converta planilhas XLSX para PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('txt-to-pdf')">
                    <h3>📄 TXT para PDF</h3>
                    <p>Converta arquivos de texto simples para PDF</p>
                </div>
                <div class="tool-card" onclick="showTool('pdf-to-word')">
                    <h3>🔄 PDF para Word</h3>
                    <p>Converta documentos PDF para Word (.docx) editável</p>
                </div>
                <div class="tool-card" onclick="showTool('pdf-to-text')">
                    <h3>📄 PDF para Texto</h3>
                    <p>Extraia o texto do PDF para um arquivo TXT editável</p>
                </div>
                <div class="tool-card" onclick="showTool('ocr-pdf')">
                    <h3>🔍 OCR em PDF</h3>
                    <p>Extraia texto de PDFs e imagens escaneadas com OCR</p>
                </div>
                <div class="tool-card" onclick="showEditor()">
                    <h3>🖥️ Editar PDF</h3>
                    <p>Reordene, insira, gire, duplique e exclua páginas diretamente no PDF</p>
                </div>
            </div>
        </div>

        <div id="editor-view" class="hidden">
            <button class="back-btn" onclick="showHomeFromEditor()">← Voltar</button>
            <div class="tool-card">
                <h3>🖥️ Editor de páginas PDF</h3>
                <p>Organize a estrutura do seu PDF sem alterar o arquivo original.</p>
                <div class="upload-area" id="editor-upload-area" onclick="document.getElementById('editor-file-input').click()">
                    <input type="file" id="editor-file-input" class="editor-file-input" accept=".pdf">
                    <p>📁 Escolha um PDF para começar</p>
                    <button class="upload-btn" type="button">Abrir PDF</button>
                </div>
                <div id="editor-shell" class="editor-shell hidden">
                    <div class="editor-toolbar">
                        <button id="editor-undo" type="button" onclick="editorUndo()" disabled>↶ Desfazer</button>
                        <button id="editor-redo" type="button" onclick="editorRedo()" disabled>↷ Refazer</button>
                        <button type="button" onclick="addBlankEditorPage()">＋ Página em branco</button>
                        <button type="button" onclick="document.getElementById('editor-add-input').click()">＋ Inserir PDF ou imagem</button>
                        <input type="file" id="editor-add-input" class="editor-file-input" accept=".pdf,.jpg,.jpeg,.png" multiple>
                    </div>
                    <div id="editor-insert-panel" class="editor-insert-panel hidden">
                        <h4>Escolha as páginas para inserir</h4>
                        <p>Selecione uma ou mais páginas do arquivo adicional antes de adicioná-las ao documento.</p>
                        <div id="editor-pending-pages" class="editor-pending-pages"></div>
                        <div class="editor-insert-actions">
                            <button class="primary" type="button" onclick="confirmEditorInsert()">Inserir selecionadas</button>
                            <button class="secondary" type="button" onclick="cancelEditorInsert()">Cancelar</button>
                        </div>
                    </div>
                    <div id="editor-pages" class="editor-pages">
                        <div class="editor-empty">As páginas do PDF aparecerão aqui.</div>
                    </div>
                    <div class="editor-actions">
                        <button class="primary" type="button" onclick="exportEditedPdf()">Salvar PDF editado</button>
                        <button class="secondary" type="button" onclick="selectAllEditorPages()">Selecionar todas</button>
                        <button type="button" onclick="deleteSelectedEditorPages()">Excluir selecionadas</button>
                        <span id="editor-status" class="editor-status">Nenhum PDF aberto.</span>
                    </div>
                </div>
                <div id="editor-result" class="hidden"></div>
            </div>
        </div>

        <!-- Tool Views -->
        <div id="tool-views" class="hidden">
            <button class="back-btn" onclick="showHome()">← Voltar</button>
            <div class="tool-card">
                <h3 id="tool-title"></h3>
                <p id="tool-description"></p>

                <div class="upload-area" id="upload-area" onclick="document.getElementById('file-input').click()">
                    <input type="file" id="file-input" class="file-input" multiple accept=".pdf,.docx,.jpg,.jpeg,.png,.txt,.xlsx">
                    <p>📁 Clique aqui ou arraste arquivos para fazer upload</p>
                    <button class="upload-btn">Escolher Arquivos</button>
                </div>

                <div id="file-list" class="file-list"></div>

                <div id="options" class="hidden">
                    <!-- Opções específicas para cada ferramenta -->
                </div>

                <button id="convert-btn" class="convert-btn hidden" onclick="convertFiles()">Converter</button>

                <div id="progress" class="progress hidden">
                    <div id="progress-bar" class="progress-bar"></div>
                </div>

                <div id="result" class="hidden"></div>
            </div>
        </div>

        <div id="privacy-note" class="footer">
            <div class="footer-grid">
                <div class="footer-block">
                    <h4>LocalPDF.io</h4>
                    <p>Ferramentas PDF gratuitas, locais e privadas. Sem cadastro e sem upload externo no modo local.</p>
                </div>
                <div class="footer-block">
                    <h4>Repositório</h4>
                    <a href="https://github.com/virgiliojr94/localpdf.io" target="_blank">Código no GitHub</a>
                    <a href="https://github.com/virgiliojr94/localpdf.io/blob/main/README.pt-br.md" target="_blank">Documentação</a>
                    <a href="https://github.com/virgiliojr94/localpdf.io/blob/main/ROADMAP.md" target="_blank">Roadmap</a>
                </div>
                <div class="footer-block">
                    <h4>Desenvolvimento</h4>
                    <p>Desenvolvido por Virgilio Borges</p>
                    <a href="mailto:virgilio.junior94@gmail.com">virgilio.junior94@gmail.com</a>
                    <a href="https://www.linkedin.com/in/virgiliojunior94/" target="_blank">LinkedIn do Virgilio</a>
                </div>
                <div class="footer-block">
                    <h4>Contribuição</h4>
                    <p>Joadson Rocha<br>Desenvolvedor Full Stack &amp; Desktop</p>
                    <a href="https://joadsonrocha.github.io/" target="_blank">Portfólio e repositório</a>
                </div>
            </div>
            <p class="footer-credit">Licença MIT · Processamento local · <a href="https://github.com/virgiliojr94/localpdf.io" target="_blank">Contribua com o projeto</a></p>
        </div>
    </div>

    <script>
        let currentTool = '';
        let currentLanguage = localStorage.getItem('localpdf-language') || 'pt';

        const languageTexts = {
            'PDF simples, privado e local': 'Simple, private and local PDF',
            'Trabalhe com seus PDFs sem complicação': 'Work with your PDFs without the hassle',
            'Converta, organize e edite documentos diretamente no seu computador. Sem contas, sem nuvem e sem enviar seus arquivos para fora.': 'Convert, organize and edit documents directly on your computer. No accounts, no cloud and no files sent elsewhere.',
            'Juntar PDF': 'Merge PDF',
            'Dividir PDF': 'Split PDF',
            'Comprimir PDF': 'Compress PDF',
            'Converter PDF': 'Convert PDF',
            'Todas as ferramentas': 'All tools',
            '100% local': '100% local',
            'Todas': 'All',
            'Organizar PDF': 'Organize PDF',
            'Otimizar PDF': 'Optimize PDF',
            'OCR': 'OCR',
            'Desenvolvimento': 'Development',
            'Contribuição': 'Contribution',
            'Repositório': 'Repository',
            'Código no GitHub': 'Code on GitHub',
            'Documentação': 'Documentation',
            'Portfólio e repositório': 'Portfolio and repository',
            'Desenvolvido por Virgilio Borges': 'Developed by Virgilio Borges',
            'LinkedIn do Virgilio': "Virgilio's LinkedIn",
            'Licença MIT · Processamento local · ': 'MIT License · Local processing · ',
            'Contribua com o projeto': 'Contribute to the project',
            'Desenvolvedor Full Stack & Desktop': 'Full Stack & Desktop Developer',
            '← Voltar': '← Back',
            '🖥️ Editor de páginas PDF': '🖥️ PDF page editor',
            'Organize a estrutura do seu PDF sem alterar o arquivo original.': 'Organize your PDF without changing the original file.',
            '📁 Escolha um PDF para começar': '📁 Choose a PDF to start',
            'Abrir PDF': 'Open PDF',
            '↶ Desfazer': '↶ Undo',
            '↷ Refazer': '↷ Redo',
            '＋ Página em branco': '＋ Blank page',
            '＋ Inserir PDF ou imagem': '＋ Insert PDF or image',
            'As páginas do PDF aparecerão aqui.': 'PDF pages will appear here.',
            'Salvar PDF editado': 'Save edited PDF',
            'Selecionar todas': 'Select all',
            'Excluir selecionadas': 'Delete selected',
            'Nenhum PDF aberto.': 'No PDF open.',
            '📁 Clique aqui ou arraste arquivos para fazer upload': '📁 Click here or drag files to upload',
            'Escolher Arquivos': 'Choose files',
            'Converter': 'Convert',
            'Escolha as páginas para inserir': 'Choose pages to insert',
            'Selecione uma ou mais páginas do arquivo adicional antes de adicioná-las ao documento.': 'Select one or more pages from the additional file before adding them to the document.',
            'Inserir selecionadas': 'Insert selected',
            'Cancelar': 'Cancel',
            'A senha é usada somente durante o processamento local.': 'The password is used only during local processing.'
            ,'🖼️ PDF para Imagens': '🖼️ PDF to Images'
            ,'📄 Imagens para PDF': '📄 Images to PDF'
            ,'🔗 Mesclar PDFs': '🔗 Merge PDFs'
            ,'✂️ Dividir PDF': '✂️ Split PDF'
            ,'📦 Comprimir PDF': '📦 Compress PDF'
            ,'🔐 Proteger PDF': '🔐 Protect PDF'
            ,"💧 Marca d'água": '💧 Watermark'
            ,'🔢 Números de página': '🔢 Page numbers'
            ,'🔒 PDF para PDF/A': '🔒 PDF to PDF/A'
            ,'📝 Word para PDF': '📝 Word to PDF'
            ,'📊 Excel para PDF': '📊 Excel to PDF'
            ,'📄 TXT para PDF': '📄 TXT to PDF'
            ,'🔄 PDF para Word': '🔄 PDF to Word'
            ,'📄 PDF para Texto': '📄 PDF to Text'
            ,'🔍 OCR em PDF': '🔍 OCR PDF'
            ,'Converta páginas PDF em imagens JPG ou PNG': 'Convert PDF pages into JPG or PNG images'
            ,'Combine várias imagens em um único PDF': 'Combine multiple images into one PDF'
            ,'Combine vários arquivos PDF em um documento único': 'Combine multiple PDF files into one document'
            ,'Extraia páginas específicas do seu PDF': 'Extract specific pages from your PDF'
            ,'Reduza o tamanho do seu arquivo PDF': 'Reduce your PDF file size'
            ,'Adicione uma senha local ao seu documento PDF': 'Add a local password to your PDF'
            ,"Adicione uma marca d'água de texto ao PDF": 'Add a text watermark to your PDF'
            ,'Numere as páginas do documento localmente': 'Number your document pages locally'
            ,'Padronize seu PDF para arquivamento (PDF/A)': 'Convert your PDF to the archival PDF/A standard'
            ,'Converta um ou mais documentos DOCX para PDF': 'Convert one or more DOCX documents to PDF'
            ,'Converta planilhas XLSX para PDF': 'Convert XLSX spreadsheets to PDF'
            ,'Converta arquivos de texto simples para PDF': 'Convert plain text files to PDF'
            ,'Converta documentos PDF para Word (.docx) editável': 'Convert PDF documents to editable Word (.docx) files'
            ,'Extraia o texto do PDF para um arquivo TXT editável': 'Extract PDF text into an editable TXT file'
            ,'Extraia texto de PDFs e imagens escaneadas com OCR': 'Extract text from scanned PDFs and images with OCR'
        };

        const toolTranslations = {
            'pdf-to-images': ['🖼️ PDF to Images', 'Convert each PDF page into separate JPG or PNG images'],
            'images-to-pdf': ['📄 Images to PDF', 'Combine multiple images into one PDF'],
            'merge-pdf': ['🔗 Merge PDFs', 'Combine multiple PDF files into one document'],
            'split-pdf': ['✂️ Split PDF', 'Extract specific pages from your PDF'],
            'compress-pdf': ['📦 Compress PDF', 'Reduce PDF file size while preserving quality'],
            'protect-pdf': ['🔐 Protect PDF', 'Create a password-protected copy of your PDF'],
            'watermark-pdf': ["💧 Watermark", 'Add a text watermark to every page'],
            'page-numbers-pdf': ['🔢 Page numbers', 'Add numbering to your PDF document'],
            'pdf-to-pdfa': ['🔒 PDF to PDF/A', 'Convert PDFs to the PDF/A-1b archival standard'],
            'word-to-pdf': ['📝 Word to PDF', 'Convert one or more DOCX files to PDF'],
            'excel-to-pdf': ['📊 Excel to PDF', 'Convert Excel spreadsheets to PDF'],
            'txt-to-pdf': ['📄 TXT to PDF', 'Convert plain text files to formatted PDF'],
            'pdf-to-word': ['🔄 PDF to Word', 'Convert PDF documents into editable DOCX files'],
            'pdf-to-text': ['📄 PDF to Text', 'Extract selectable text from every PDF page'],
            'ocr-pdf': ['🔍 OCR PDF', 'Extract text from scanned PDFs and images using OCR']
        };

        function translatePage() {
            document.documentElement.lang = currentLanguage === 'en' ? 'en' : 'pt-BR';
            document.getElementById('language-toggle').textContent = currentLanguage === 'en' ? 'PT' : 'EN';
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            const translations = currentLanguage === 'en' ? languageTexts : Object.fromEntries(Object.entries(languageTexts).map(([pt, en]) => [en, pt]));
            let node;
            while ((node = walker.nextNode())) {
                const value = node.nodeValue.trim();
                if (translations[value]) node.nodeValue = node.nodeValue.replace(value, translations[value]);
            }
        }

        function toggleLanguage() {
            currentLanguage = currentLanguage === 'pt' ? 'en' : 'pt';
            localStorage.setItem('localpdf-language', currentLanguage);
            translatePage();
            if (currentTool) showTool(currentTool);
        }
        let uploadedFiles = [];

        const tools = {
            'pdf-to-images': {
                title: '🖼️ PDF para Imagens',
                description: 'Converta cada página do seu PDF em imagens separadas',
                accept: '.pdf',
                multiple: false
            },
            'images-to-pdf': {
                title: '📄 Imagens para PDF',
                description: 'Combine múltiplas imagens em um único arquivo PDF',
                accept: '.jpg,.jpeg,.png',
                multiple: true
            },
            'merge-pdf': {
                title: '🔗 Mesclar PDFs',
                description: 'Combine vários arquivos PDF em um documento único',
                accept: '.pdf',
                multiple: true
            },
            'split-pdf': {
                title: '✂️ Dividir PDF',
                description: 'Extraia páginas específicas do seu PDF',
                accept: '.pdf',
                multiple: false
            },
            'compress-pdf': {
                title: '📦 Comprimir PDF',
                description: 'Reduza o tamanho do arquivo PDF mantendo a qualidade',
                accept: '.pdf',
                multiple: false
            },
            'protect-pdf': {
                title: '🔐 Proteger PDF',
                description: 'Crie uma cópia protegida do seu PDF com senha',
                accept: '.pdf',
                multiple: false,
                options: 'password'
            },
            'watermark-pdf': {
                title: '💧 Marca d\'água',
                description: 'Adicione uma marca d\'água de texto em todas as páginas',
                accept: '.pdf',
                multiple: false,
                options: 'watermark'
            },
            'page-numbers-pdf': {
                title: '🔢 Números de página',
                description: 'Adicione numeração ao seu documento PDF',
                accept: '.pdf',
                multiple: false,
                options: 'page-numbers'
            },
            'pdf-to-pdfa': {
                title: '🔒 PDF para PDF/A',
                description: 'Converta PDFs para o padrão de arquivamento PDF/A-1b',
                accept: '.pdf',
                multiple: true
            },
            'word-to-pdf': {
                title: '📝 Word para PDF',
                description: 'Converta documentos Word (.docx) para PDF - aceita múltiplos arquivos',
                accept: '.docx',
                multiple: true
            },
            'excel-to-pdf': {
                title: '📊 Excel para PDF',
                description: 'Converta planilhas Excel (.xlsx) para PDF',
                accept: '.xlsx',
                multiple: false
            },
            'txt-to-pdf': {
                title: '📄 TXT para PDF',
                description: 'Converta arquivos de texto simples (.txt) para PDF',
                accept: '.txt',
                multiple: false
            },
            'pdf-to-word': {
                title: '🔄 PDF para Word',
                description: 'Converta seus documentos PDF para Word (.docx) editável',
                accept: '.pdf',
                multiple: false
            },
            'pdf-to-text': {
                title: '📄 PDF para Texto',
                description: 'Extraia o texto selecionável de todas as páginas do PDF',
                accept: '.pdf',
                multiple: false
            },
            'ocr-pdf': {
                title: '🔍 OCR em PDF',
                description: 'Extraia texto de PDFs e imagens escaneadas usando reconhecimento óptico de caracteres (Tesseract)',
                accept: '.pdf,.jpg,.jpeg,.png',
                multiple: false
            }
        };

        function showTool(toolName) {
            currentTool = toolName;
            const tool = tools[toolName];

            document.getElementById('home-view').classList.add('hidden');
            document.getElementById('tool-views').classList.remove('hidden');
            const translatedTool = toolTranslations[toolName];
            document.getElementById('tool-title').innerText = currentLanguage === 'en' ? translatedTool[0] : tool.title;
            document.getElementById('tool-description').innerText = currentLanguage === 'en' ? translatedTool[1] : tool.description;
            document.getElementById('file-input').accept = tool.accept;
            document.getElementById('file-input').multiple = tool.multiple;
            renderToolOptions(tool.options);

            uploadedFiles = [];
            updateFileList();
            hideResult();
        }

        function renderToolOptions(optionType) {
            const options = document.getElementById('options');
            if (optionType === 'password') {
                options.innerHTML = `
                    <label for="pdf-password">Senha do PDF</label>
                    <input id="pdf-password" type="password" minlength="4" autocomplete="new-password" placeholder="Digite uma senha com pelo menos 4 caracteres">
                    <small>A senha é usada somente durante o processamento local.</small>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'watermark') {
                options.innerHTML = `
                    <label for="watermark-text">Texto da marca d'água</label>
                    <input id="watermark-text" type="text" maxlength="80" placeholder="Ex.: CONFIDENCIAL">
                    <label for="watermark-position">Posição</label>
                    <select id="watermark-position">
                        <option value="center">Centro</option>
                        <option value="top">Parte superior</option>
                        <option value="bottom">Parte inferior</option>
                    </select>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'page-numbers') {
                options.innerHTML = `
                    <label for="page-number-position">Posição da numeração</label>
                    <select id="page-number-position">
                        <option value="bottom-center">Rodapé central</option>
                        <option value="bottom-right">Rodapé direito</option>
                        <option value="top-center">Cabeçalho central</option>
                        <option value="top-right">Cabeçalho direito</option>
                    </select>
                `;
                options.classList.remove('hidden');
                return;
            }
            options.innerHTML = '';
            options.classList.add('hidden');
        }

        function showHome() {
            document.getElementById('home-view').classList.remove('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            uploadedFiles = [];
        }

        function updateFileList() {
            const fileList = document.getElementById('file-list');
            const convertBtn = document.getElementById('convert-btn');

            if (uploadedFiles.length === 0) {
                fileList.innerHTML = '';
                convertBtn.classList.add('hidden');
                return;
            }

            fileList.innerHTML = uploadedFiles.map((file, index) => `
                <div class="file-item">
                    <span>📄 ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)</span>
                    <button onclick="removeFile(${index})" style="background: #dc3545; color: white; border: none; padding: 5px 10px; border-radius: 5px; cursor: pointer;">Remover</button>
                </div>
            `).join('');

            convertBtn.classList.remove('hidden');
        }

        function removeFile(index) {
            uploadedFiles.splice(index, 1);
            updateFileList();
        }

        function hideResult() {
            document.getElementById('result').classList.add('hidden');
            document.getElementById('progress').classList.add('hidden');
        }

        // Upload de arquivos
        document.getElementById('file-input').addEventListener('change', function(e) {
            const files = Array.from(e.target.files);
            if (tools[currentTool].multiple) {
                uploadedFiles = uploadedFiles.concat(files);
            } else {
                uploadedFiles = files.slice(0, 1);
            }
            updateFileList();
        });

        // Drag and drop
        const uploadArea = document.getElementById('upload-area');
        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');

            const files = Array.from(e.dataTransfer.files);
            if (tools[currentTool].multiple) {
                uploadedFiles = uploadedFiles.concat(files);
            } else {
                uploadedFiles = files.slice(0, 1);
            }
            updateFileList();
        });

        async function convertFiles() {
            if (uploadedFiles.length === 0) return;

            const formData = new FormData();
            uploadedFiles.forEach(file => {
                formData.append('files', file);
            });
            formData.append('tool', currentTool);
            const passwordInput = document.getElementById('pdf-password');
            if (passwordInput) {
                if (!passwordInput.value || passwordInput.value.length < 4) {
                    document.getElementById('result').innerHTML = '<h4>⚠️ Senha inválida</h4><p>Informe uma senha com pelo menos 4 caracteres.</p>';
                    document.getElementById('result').classList.remove('hidden');
                    return;
                }
                formData.append('password', passwordInput.value);
            }
            const watermarkText = document.getElementById('watermark-text');
            if (watermarkText) {
                if (!watermarkText.value.trim()) {
                    document.getElementById('result').innerHTML = '<h4>⚠️ Texto obrigatório</h4><p>Informe o texto da marca d\'água.</p>';
                    document.getElementById('result').classList.remove('hidden');
                    return;
                }
                formData.append('watermark_text', watermarkText.value.trim());
                formData.append('watermark_position', document.getElementById('watermark-position').value);
            }
            const pageNumberPosition = document.getElementById('page-number-position');
            if (pageNumberPosition) {
                formData.append('page_number_position', pageNumberPosition.value);
            }

            document.getElementById('progress').classList.remove('hidden');
            document.getElementById('convert-btn').disabled = true;

            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = response.headers.get('Content-Disposition')?.split('filename=')[1] || 'converted_file.zip';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);

                    document.getElementById('result').innerHTML = '<h4>✅ Sucesso!</h4><p>Arquivo convertido e baixado com sucesso!</p>';
                    document.getElementById('result').classList.remove('hidden');
                } else {
                    throw new Error('Erro na conversão');
                }
            } catch (error) {
                document.getElementById('result').innerHTML = '<h4>❌ Erro!</h4><p>Ocorreu um erro durante a conversão. Tente novamente.</p>';
                document.getElementById('result').classList.remove('hidden');
            } finally {
                document.getElementById('progress').classList.add('hidden');
                document.getElementById('convert-btn').disabled = false;
            }
        }

        let editorFiles = [];
        let editorPages = [];
        let editorHistory = [];
        let editorFuture = [];
        let editorDraggedIndex = null;
        let editorPendingPages = [];
        let editorPendingFiles = [];

        function showEditor() {
            document.getElementById('home-view').classList.add('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            document.getElementById('editor-view').classList.remove('hidden');
            resetEditor();
        }

        function showHomeFromEditor() {
            document.getElementById('editor-view').classList.add('hidden');
            document.getElementById('home-view').classList.remove('hidden');
            resetEditor();
        }

        function resetEditor() {
            editorFiles = [];
            editorPages = [];
            editorHistory = [];
            editorFuture = [];
            editorPendingPages = [];
            editorPendingFiles = [];
            document.getElementById('editor-file-input').value = '';
            document.getElementById('editor-add-input').value = '';
            document.getElementById('editor-shell').classList.add('hidden');
            document.getElementById('editor-insert-panel').classList.add('hidden');
            document.getElementById('editor-result').classList.add('hidden');
            renderEditorPages();
            updateEditorHistoryButtons();
        }

        function editorSnapshot() {
            return editorPages.map(page => ({...page, thumbnail: page.thumbnail}));
        }

        function editorPushHistory() {
            editorHistory.push(editorSnapshot());
            editorFuture = [];
            updateEditorHistoryButtons();
        }

        function editorUndo() {
            if (!editorHistory.length) return;
            editorFuture.push(editorSnapshot());
            editorPages = editorHistory.pop();
            renderEditorPages();
            updateEditorHistoryButtons();
        }

        function editorRedo() {
            if (!editorFuture.length) return;
            editorHistory.push(editorSnapshot());
            editorPages = editorFuture.pop();
            renderEditorPages();
            updateEditorHistoryButtons();
        }

        function updateEditorHistoryButtons() {
            document.getElementById('editor-undo').disabled = !editorHistory.length;
            document.getElementById('editor-redo').disabled = !editorFuture.length;
        }

        function editorFormData(files) {
            const formData = new FormData();
            files.forEach(file => formData.append('files', file));
            return formData;
        }

        async function openEditorFiles(files, append = false) {
            if (!files.length) return;
            const response = await fetch('/editor/preview', {
                method: 'POST',
                body: editorFormData(files)
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'Não foi possível abrir o arquivo.');

            const sourceOffset = editorFiles.length;
            const incomingPages = data.pages.map(page => ({
                ...page,
                source_file: page.source_file + sourceOffset,
                selected: false
            }));
            if (!append) {
                editorPages = incomingPages;
                editorFiles = [...files];
            } else {
                editorPendingPages = incomingPages;
                editorPendingFiles = [...files];
                renderPendingEditorPages();
                document.getElementById('editor-insert-panel').classList.remove('hidden');
                document.getElementById('editor-status').textContent = 'Escolha as páginas que deseja inserir.';
                return;
            }
            document.getElementById('editor-shell').classList.remove('hidden');
            document.getElementById('editor-status').textContent = `${editorPages.length} página(s) no documento.`;
            renderEditorPages();
        }

        function renderPendingEditorPages() {
            const container = document.getElementById('editor-pending-pages');
            container.innerHTML = editorPendingPages.map((page, index) => `
                <div class="editor-pending-page${page.selected ? ' selected' : ''}" data-pending-index="${index}">
                    <img src="${page.thumbnail}" alt="Página disponível ${index + 1}">
                    <label><input type="checkbox" ${page.selected ? 'checked' : ''}> Página ${index + 1}</label>
                </div>
            `).join('');
            container.querySelectorAll('.editor-pending-page').forEach(card => {
                const index = Number(card.dataset.pendingIndex);
                card.addEventListener('click', event => {
                    if (event.target.tagName !== 'INPUT') event.preventDefault();
                    editorPendingPages[index].selected = !editorPendingPages[index].selected;
                    renderPendingEditorPages();
                });
            });
        }

        function confirmEditorInsert() {
            const selectedPages = editorPendingPages.filter(page => page.selected);
            if (!selectedPages.length) {
                document.getElementById('editor-status').textContent = 'Selecione pelo menos uma página para inserir.';
                return;
            }
            editorPushHistory();
            editorPages = editorPages.concat(selectedPages);
            editorFiles = editorFiles.concat(editorPendingFiles);
            editorPendingPages = [];
            editorPendingFiles = [];
            document.getElementById('editor-insert-panel').classList.add('hidden');
            document.getElementById('editor-status').textContent = `${editorPages.length} página(s) no documento.`;
            renderEditorPages();
        }

        function cancelEditorInsert() {
            editorPendingPages = [];
            editorPendingFiles = [];
            document.getElementById('editor-insert-panel').classList.add('hidden');
            document.getElementById('editor-add-input').value = '';
            document.getElementById('editor-status').textContent = `${editorPages.length} página(s) no documento.`;
        }

        function renderEditorPages() {
            const container = document.getElementById('editor-pages');
            if (!editorPages.length) {
                container.innerHTML = '<div class="editor-empty">As páginas do PDF aparecerão aqui.</div>';
                return;
            }
            container.innerHTML = editorPages.map((page, index) => `
                <div class="editor-page${page.selected ? ' selected' : ''}" draggable="true" data-editor-index="${index}">
                    <img src="${page.thumbnail}" alt="Página ${index + 1}">
                    <div class="editor-page-number">Página ${index + 1}</div>
                    <div class="editor-page-actions">
                        <button type="button" data-action="rotate" title="Girar página">↻</button>
                        <button type="button" data-action="duplicate" title="Duplicar página">⧉</button>
                        <button type="button" data-action="delete" title="Excluir página">✕</button>
                    </div>
                </div>
            `).join('');

            container.querySelectorAll('.editor-page').forEach(card => {
                const index = Number(card.dataset.editorIndex);
                card.addEventListener('click', event => {
                    if (event.target.closest('button')) return;
                    editorPages[index].selected = !editorPages[index].selected;
                    renderEditorPages();
                });
                card.addEventListener('dragstart', () => {
                    editorDraggedIndex = index;
                    card.classList.add('dragging');
                });
                card.addEventListener('dragend', () => card.classList.remove('dragging'));
                card.addEventListener('dragover', event => event.preventDefault());
                card.addEventListener('drop', event => {
                    event.preventDefault();
                    if (editorDraggedIndex === null || editorDraggedIndex === index) return;
                    editorPushHistory();
                    const [movedPage] = editorPages.splice(editorDraggedIndex, 1);
                    editorPages.splice(index, 0, movedPage);
                    editorDraggedIndex = null;
                    renderEditorPages();
                });
                card.querySelector('[data-action="rotate"]').addEventListener('click', () => {
                    editorPushHistory();
                    editorPages[index].rotation = (editorPages[index].rotation + 90) % 360;
                    editorPages[index].selected = false;
                    renderEditorPages();
                });
                card.querySelector('[data-action="duplicate"]').addEventListener('click', () => {
                    editorPushHistory();
                    editorPages.splice(index + 1, 0, {...editorPages[index], id: `${editorPages[index].id}-copy-${Date.now()}`, selected: false});
                    renderEditorPages();
                });
                card.querySelector('[data-action="delete"]').addEventListener('click', () => deleteEditorPage(index));
            });
        }

        function deleteEditorPage(index) {
            if (editorPages.length <= 1) {
                alert('O PDF precisa manter pelo menos uma página.');
                return;
            }
            editorPushHistory();
            editorPages.splice(index, 1);
            renderEditorPages();
        }

        function selectAllEditorPages() {
            editorPages.forEach(page => page.selected = true);
            renderEditorPages();
        }

        function deleteSelectedEditorPages() {
            const selectedCount = editorPages.filter(page => page.selected).length;
            if (!selectedCount) return;
            if (selectedCount >= editorPages.length) {
                alert('O PDF precisa manter pelo menos uma página.');
                return;
            }
            editorPushHistory();
            editorPages = editorPages.filter(page => !page.selected);
            renderEditorPages();
        }

        function addBlankEditorPage() {
            editorPushHistory();
            editorPages.push({
                id: `blank-${Date.now()}`,
                kind: 'blank',
                source_file: -1,
                page_index: 0,
                rotation: 0,
                width: 595,
                height: 842,
                selected: false,
                thumbnail: 'data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" width="300" height="420"%3E%3Crect width="100%25" height="100%25" fill="white"/%3E%3C/svg%3E'
            });
            renderEditorPages();
        }

        async function exportEditedPdf() {
            if (!editorPages.length || !editorFiles.length) return;
            const status = document.getElementById('editor-status');
            status.textContent = 'Gerando PDF editado...';
            const formData = editorFormData(editorFiles);
            formData.append('pages', JSON.stringify(editorPages.map(({thumbnail, selected, ...page}) => page)));
            try {
                const response = await fetch('/editor/export', {method: 'POST', body: formData});
                if (!response.ok) {
                    const data = await response.json();
                    throw new Error(data.error || 'Não foi possível exportar o PDF.');
                }
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'localpdf-editado.pdf';
                link.click();
                URL.revokeObjectURL(url);
                status.textContent = 'PDF editado baixado com sucesso.';
            } catch (error) {
                status.textContent = error.message;
            }
        }

        document.getElementById('editor-file-input').addEventListener('change', async event => {
            try {
                await openEditorFiles(Array.from(event.target.files));
            } catch (error) {
                document.getElementById('editor-status').textContent = error.message;
            }
        });

        document.getElementById('editor-add-input').addEventListener('change', async event => {
            try {
                await openEditorFiles(Array.from(event.target.files), true);
            } catch (error) {
                document.getElementById('editor-status').textContent = error.message;
            }
        });
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/favicon.svg")
def favicon():
    return send_file(
        os.path.join(os.path.dirname(__file__), "favicon.svg"),
        mimetype="image/svg+xml",
    )


def render_pdf_page(page, scale=0.2):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return "data:image/png;base64," + base64.b64encode(pixmap.tobytes("png")).decode(
        "ascii"
    )


def editor_page_descriptor(page, source_file, page_index, thumbnail):
    return {
        "id": f"file-{source_file}-page-{page_index}",
        "kind": "pdf",
        "source_file": source_file,
        "page_index": page_index,
        "rotation": page.rotation,
        "width": page.rect.width,
        "height": page.rect.height,
        "thumbnail": thumbnail,
    }


def editor_preview_for_file(file, source_file, temp_dir):
    filename = secure_filename(file.filename or "")
    extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    input_path = os.path.join(temp_dir, f"source_{source_file}_{filename}")
    file.save(input_path)

    if extension == "pdf":
        with fitz.open(input_path) as document:
            return [
                editor_page_descriptor(
                    page,
                    source_file,
                    page_index,
                    render_pdf_page(page),
                )
                for page_index, page in enumerate(document)
            ]

    if extension in {"jpg", "jpeg", "png"}:
        with Image.open(input_path) as image:
            width, height = image.size
        image_document = fitz.open()
        page = image_document.new_page(width=width, height=height)
        page.insert_image(page.rect, filename=input_path)
        thumbnail = render_pdf_page(page)
        image_document.close()
        return [
            {
                "id": f"file-{source_file}-image-0",
                "kind": "image",
                "source_file": source_file,
                "page_index": 0,
                "rotation": 0,
                "width": width,
                "height": height,
                "thumbnail": thumbnail,
            }
        ]

    raise ValueError(f"Formato não suportado pelo editor: {filename}")


@app.route("/editor/preview", methods=["POST"])
def editor_preview():
    files = request.files.getlist("files")
    if not files or any(not file.filename for file in files):
        return jsonify({"error": "Envie pelo menos um PDF ou uma imagem."}), 400

    if any(
        not allowed_file(file.filename)
        or file.filename.rsplit(".", 1)[-1].lower() not in {"pdf", "jpg", "jpeg", "png"}
        for file in files
    ):
        return jsonify({"error": "O editor aceita apenas PDF, JPG e PNG."}), 400

    temp_dir = tempfile.mkdtemp()
    try:
        pages = []
        for source_file, file in enumerate(files):
            pages.extend(editor_preview_for_file(file, source_file, temp_dir))
        if not pages:
            return jsonify({"error": "O arquivo não possui páginas editáveis."}), 400
        return jsonify({"pages": pages})
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao gerar preview do editor")
        return jsonify({"error": "Não foi possível abrir o documento."}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def insert_editor_image(document, file_path, rotation=0):
    with Image.open(file_path) as image:
        width, height = image.size
    page = document.new_page(width=width, height=height)
    page.insert_image(page.rect, filename=file_path)
    page.set_rotation(rotation % 360)


def export_editor_pages(files, pages, temp_dir):
    if not isinstance(pages, list) or not pages:
        raise ValueError("O editor precisa receber pelo menos uma página.")

    source_paths = []
    for source_file, file in enumerate(files):
        filename = secure_filename(file.filename or "")
        extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        if extension not in {"pdf", "jpg", "jpeg", "png"}:
            raise ValueError("O editor aceita apenas PDF, JPG e PNG.")
        source_path = os.path.join(temp_dir, f"source_{source_file}_{filename}")
        file.save(source_path)
        source_paths.append((extension, source_path))

    source_documents = {
        index: fitz.open(path)
        for index, (extension, path) in enumerate(source_paths)
        if extension == "pdf"
    }
    output_document = fitz.open()
    try:
        for page_data in pages:
            kind = page_data.get("kind")
            source_file = int(page_data.get("source_file", -1))
            page_index = int(page_data.get("page_index", -1))
            rotation = int(page_data.get("rotation", 0)) % 360

            if kind == "blank":
                output_page = output_document.new_page()
                output_page.set_rotation(rotation)
                continue

            if source_file < 0 or source_file >= len(source_paths):
                raise ValueError("Página vinculada a arquivo inexistente.")

            extension, source_path = source_paths[source_file]
            if extension == "pdf":
                source_document = source_documents[source_file]
                if page_index < 0 or page_index >= len(source_document):
                    raise ValueError("Índice de página inválido.")
                output_document.insert_pdf(
                    source_document,
                    from_page=page_index,
                    to_page=page_index,
                )
                output_document[-1].set_rotation(rotation)
            elif extension in {"jpg", "jpeg", "png"}:
                insert_editor_image(output_document, source_path, rotation)
            else:
                raise ValueError("Formato de página não suportado.")

        output_path = os.path.join(temp_dir, "edited.pdf")
        output_document.save(output_path)
        return output_path
    finally:
        output_document.close()
        for document in source_documents.values():
            document.close()


@app.route("/editor/export", methods=["POST"])
def editor_export():
    files = request.files.getlist("files")
    if not files or not files[0].filename:
        return jsonify({"error": "Envie o PDF que será editado."}), 400

    try:
        pages = json.loads(request.form.get("pages", "[]"))
    except json.JSONDecodeError:
        return jsonify({"error": "A composição de páginas é inválida."}), 400

    temp_dir = tempfile.mkdtemp()
    try:
        output_path = export_editor_pages(files, pages, temp_dir)
        with open(output_path, "rb") as output_file:
            data = output_file.read()
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name="localpdf-editado.pdf",
        )
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao exportar PDF do editor")
        return jsonify({"error": "Não foi possível exportar o PDF."}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def excel_to_pdf(file, temp_dir):
    xlsx_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(xlsx_path)

    pdf_path = os.path.join(temp_dir, "excel_to_pdf.pdf")
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position = height - 50

    try:
        workbook = openpyxl.load_workbook(xlsx_path)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            c.setFont("Helvetica", 10)
            c.drawString(50, y_position, f"--- Planilha: {sheet_name} ---")
            y_position -= 20

            for row_idx, row in enumerate(sheet.iter_rows()):
                row_data = [
                    str(cell.value) if cell.value is not None else "" for cell in row
                ]
                line_text = " | ".join(row_data)

                # Simples quebra de linha para caber na página
                max_line_width = int(
                    (width - 100) / 6
                )  # Estimativa de caracteres por linha
                if len(line_text) > max_line_width:
                    # Implementação mais robusta de quebra de linha seria necessária
                    line_text = line_text[:max_line_width] + "..."

                if y_position < 50:
                    c.showPage()
                    y_position = height - 50
                    c.setFont("Helvetica", 10)  # Reset font after new page

                c.drawString(50, y_position, line_text)
                y_position -= 15  # Espaçamento menor para linhas de planilha

            y_position -= 30  # Espaçamento entre planilhas
            if (
                y_position < 50 and sheet_name != workbook.sheetnames[-1]
            ):  # Only show new page if not last sheet
                c.showPage()
                y_position = height - 50

    except Exception as e:
        # Handle potential errors with Excel files
        c.drawString(50, y_position - 20, f"Erro ao ler planilha: {e}")
        print(f"Erro ao ler planilha Excel: {e}")

    c.save()
    return [pdf_path]


def txt_to_pdf(file, temp_dir):
    txt_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(txt_path)

    pdf_path = os.path.join(temp_dir, "text_to_pdf.pdf")
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position = height - 50

    c.setFont("Helvetica", 12)

    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                # Simples quebra de linha para caber na página
                text_line = line.strip()
                max_width_px = width - 100  # Margens de 50px de cada lado

                # Estimar a largura do texto para quebrar linhas
                # ReportLab não tem quebra automática de texto complexa por default
                # Esta é uma estimativa MUITO simples; para algo robusto, precisaria de TextObject
                approx_char_width_px = 7  # Média para Helvetica 12
                chars_per_line = int(max_width_px / approx_char_width_px)

                if len(text_line) > chars_per_line:
                    # Quebra simples da linha
                    chunks = [
                        text_line[i : i + chars_per_line]
                        for i in range(0, len(text_line), chars_per_line)
                    ]
                else:
                    chunks = [text_line]

                for chunk in chunks:
                    if y_position < 50:  # Margem inferior
                        c.showPage()
                        y_position = height - 50
                        c.setFont("Helvetica", 12)  # Reset font after new page

                    c.drawString(50, y_position, chunk)
                    y_position -= 15  # Espaçamento entre linhas

    except Exception as e:
        c.drawString(50, y_position - 20, f"Erro ao ler arquivo de texto: {e}")
        print(f"Erro ao ler arquivo de texto: {e}")

    c.save()
    return [pdf_path]


@app.route("/convert", methods=["POST"])
def convert():
    if "files" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    files = request.files.getlist("files")
    tool = request.form.get("tool")

    if not files or files[0].filename == "":
        return jsonify({"error": "Nenhum arquivo selecionado"}), 400

    # Validação de extensão dos arquivos enviados
    for f in files:
        if not allowed_file(f.filename):
            return jsonify({"error": f"Extensão não permitida: {f.filename}"}), 400

    # Criar diretório temporário
    temp_dir = tempfile.mkdtemp()
    response = None
    try:
        if tool == "pdf-to-images":
            output_files = pdf_to_images(files[0], temp_dir)
        elif tool == "images-to-pdf":
            output_files = images_to_pdf(files, temp_dir)
        elif tool == "merge-pdf":
            output_files = merge_pdfs(files, temp_dir)
        elif tool == "split-pdf":
            output_files = split_pdf(files[0], temp_dir)
        elif tool == "compress-pdf":
            output_files = compress_pdf(files[0], temp_dir)
        elif tool == "protect-pdf":
            output_files = protect_pdf(files[0], temp_dir, request.form.get("password", ""))
        elif tool == "watermark-pdf":
            output_files = watermark_pdf(
                files[0],
                temp_dir,
                request.form.get("watermark_text", ""),
                request.form.get("watermark_position", "center"),
            )
        elif tool == "page-numbers-pdf":
            output_files = page_numbers_pdf(
                files[0], temp_dir, request.form.get("page_number_position", "bottom-center")
            )
        elif tool == "pdf-to-pdfa":
            output_files = pdf_to_pdfa(files, temp_dir)
        elif tool == "word-to-pdf":
            output_files = word_to_pdf(files, temp_dir)
        elif tool == "excel-to-pdf":
            output_files = excel_to_pdf(files[0], temp_dir)
        elif tool == "txt-to-pdf":
            output_files = txt_to_pdf(files[0], temp_dir)
        elif tool == "pdf-to-word":
            output_files = pdf_to_word(files[0], temp_dir)
        elif tool == "pdf-to-text":
            output_files = pdf_to_text(files[0], temp_dir)
        elif tool == "ocr-pdf":
            output_files = ocr_pdf(files[0], temp_dir)
        else:
            return jsonify({"error": "Ferramenta não suportada"}), 400

        response = build_response(output_files, temp_dir)
        return response
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao processar ferramenta %s", tool)
        return jsonify({"error": "Não foi possível processar os arquivos."}), 500
    finally:
        # Diretório temporário limpo após preparar resposta (BytesIO) evitando remoção antecipada
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def pdf_to_images(file, temp_dir):
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    output_files = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x resolution
        img_path = os.path.join(temp_dir, f"page_{page_num + 1}.png")
        pix.save(img_path)
        output_files.append(img_path)

    doc.close()
    return output_files


def images_to_pdf(files, temp_dir):
    images = []
    for file in files:
        img_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(img_path)
        img = Image.open(img_path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)

    pdf_path = os.path.join(temp_dir, "images_to_pdf.pdf")
    images[0].save(pdf_path, save_all=True, append_images=images[1:])

    return [pdf_path]


def merge_pdfs(files, temp_dir):
    merged_doc = fitz.open()

    for file in files:
        pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(pdf_path)
        doc = fitz.open(pdf_path)
        merged_doc.insert_pdf(doc)
        doc.close()

    output_path = os.path.join(temp_dir, "merged.pdf")
    merged_doc.save(output_path)
    merged_doc.close()

    return [output_path]


def split_pdf(file, temp_dir):
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    output_files = []

    for page_num in range(len(doc)):
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
        output_path = os.path.join(temp_dir, f"page_{page_num + 1}.pdf")
        new_doc.save(output_path)
        new_doc.close()
        output_files.append(output_path)

    doc.close()
    return output_files


def compress_pdf(file, temp_dir):
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    output_path = os.path.join(temp_dir, "compressed.pdf")
    doc.save(output_path, garbage=4, deflate=True, clean=True)
    doc.close()

    return [output_path]


def protect_pdf(file, temp_dir, password):
    """Cria uma cópia criptografada do PDF usando senha local."""
    if len(password) < 4:
        raise ValueError("A senha precisa ter pelo menos 4 caracteres.")

    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, "protected.pdf")

    with fitz.open(pdf_path) as document:
        document.save(
            output_path,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=password,
            user_pw=password,
            permissions=fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT,
        )

    return [output_path]


def watermark_pdf(file, temp_dir, text, position):
    """Aplica uma marca d'água de texto em todas as páginas do PDF."""
    text = text.strip()
    if not text:
        raise ValueError("Informe o texto da marca d'água.")
    if position not in {"center", "top", "bottom"}:
        raise ValueError("Posição de marca d'água inválida.")

    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, "watermarked.pdf")

    with fitz.open(pdf_path) as document:
        for page in document:
            rect = page.rect
            fontsize = min(36, max(18, rect.width / max(len(text), 10)))
            if position == "top":
                point = (rect.width * 0.12, rect.height * 0.16)
            elif position == "bottom":
                point = (rect.width * 0.12, rect.height * 0.9)
            else:
                point = (rect.width * 0.22, rect.height * 0.55)
            page.insert_text(
                point,
                text,
                fontsize=fontsize,
                fontname="helv",
                color=(0.78, 0.18, 0.18),
                overlay=True,
            )
        document.save(output_path)

    return [output_path]


def page_numbers_pdf(file, temp_dir, position):
    """Adiciona numeração simples a todas as páginas do PDF."""
    positions = {"bottom-center", "bottom-right", "top-center", "top-right"}
    if position not in positions:
        raise ValueError("Posição de numeração inválida.")

    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, "numbered.pdf")

    with fitz.open(pdf_path) as document:
        total_pages = len(document)
        for page_index, page in enumerate(document, start=1):
            label = f"Página {page_index} de {total_pages}"
            rect = page.rect
            text_width = fitz.get_text_length(label, fontname="helv", fontsize=10)
            if position.endswith("center"):
                x_position = (rect.width - text_width) / 2
            else:
                x_position = rect.width - text_width - 36
            y_position = 28 if position.startswith("top") else rect.height - 24
            page.insert_text(
                (x_position, y_position),
                label,
                fontsize=10,
                fontname="helv",
                color=(0.25, 0.27, 0.3),
                overlay=True,
            )
        document.save(output_path)

    return [output_path]


def pdf_to_pdfa(files, temp_dir):
    """Converte um ou mais PDFs para PDF/A-1b usando Ghostscript."""
    if ghostscript is None:
        raise RuntimeError(
            "Ghostscript não está instalado no sistema. Instale-o para usar PDF/A."
        )
    if not isinstance(files, list):
        files = [files]

    output_files = []

    for file in files:
        input_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(input_path)

        base_name, _ = os.path.splitext(os.path.basename(input_path))
        output_path = os.path.join(temp_dir, f"{base_name}_pdfa.pdf")

        gs_args = [
            "gs",
            "-dPDFA=1",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOOUTERSAVE",
            "-dUseCIEColor",
            "-sProcessColorModel=DeviceRGB",
            "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=UseDeviceIndependentColor",
            "-dPDFACompatibilityPolicy=1",
            f"-sOutputFile={output_path}",
            input_path,
        ]
        gs_args = [
            arg.encode("utf-8") if isinstance(arg, str) else arg for arg in gs_args
        ]

        try:
            ghostscript.Ghostscript(*gs_args)
        except Exception as e:
            raise RuntimeError(
                f"Erro ao converter {file.filename} para PDF/A: {e}"
            ) from e

        output_files.append(output_path)

    return output_files


def word_to_pdf(files, temp_dir):
    """
    Converte um ou múltiplos arquivos DOCX para PDF
    Se houver múltiplos arquivos, mescla todos em um único PDF
    """
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    # Criar PDF de saída
    pdf_path = os.path.join(temp_dir, "word_to_pdf.pdf")
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position = height - 50

    # Se for apenas um arquivo (compatibilidade)
    if not isinstance(files, list):
        files = [files]

    # Processar cada arquivo DOCX
    for file_idx, file in enumerate(files):
        docx_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(docx_path)

        # Lê o documento Word
        doc = Document(docx_path)

        # Adicionar separador visual (exceto no primeiro documento)
        if file_idx > 0:
            # Quebra de página
            c.showPage()
            y_position = height - 50

            # Adicionar cabeçalho com nome do arquivo
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, f"{'=' * 60}")
            y_position -= 20
            c.drawString(50, y_position, f"Documento: {file.filename}")
            y_position -= 20
            c.drawString(50, y_position, f"{'=' * 60}")
            y_position -= 30
            c.setFont("Helvetica", 11)

        # Processar parágrafos
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                # Quebra texto longo em múltiplas linhas
                text = paragraph.text
                max_width = width - 100

                # Estimativa simples de largura de texto
                approx_char_width = 6
                chars_per_line = int(max_width / approx_char_width)

                words = text.split()
                lines = []
                current_line = []

                for word in words:
                    if len(" ".join(current_line + [word])) <= chars_per_line:
                        current_line.append(word)
                    else:
                        if current_line:
                            lines.append(" ".join(current_line))
                            current_line = [word]
                        else:
                            lines.append(word)

                if current_line:
                    lines.append(" ".join(current_line))

                for line in lines:
                    if y_position < 50:
                        c.showPage()
                        y_position = height - 50

                    c.drawString(50, y_position, line)
                    y_position -= 20

        # Processar tabelas (se houver)
        for table in doc.tables:
            # Adicionar espaçamento antes da tabela
            y_position -= 10

            if y_position < 100:
                c.showPage()
                y_position = height - 50

            # Desenhar linhas da tabela
            c.setFont("Helvetica", 9)
            for row in table.rows:
                row_text = " | ".join([cell.text for cell in row.cells])

                # Quebrar texto da linha se necessário
                if len(row_text) > 100:
                    row_text = row_text[:97] + "..."

                if y_position < 50:
                    c.showPage()
                    y_position = height - 50

                c.drawString(50, y_position, row_text)
                y_position -= 15

            # Espaçamento após tabela
            y_position -= 10
            c.setFont("Helvetica", 11)

    c.save()
    return [pdf_path]


def pdf_to_word(file, temp_dir):
    """
    Convert PDF to Word (.docx) format.
    """
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    docx_filename = os.path.splitext(secure_filename(file.filename))[0] + ".docx"
    docx_path = os.path.join(temp_dir, docx_filename)

    cv = None
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path)
    except ValueError as e:
        raise RuntimeError(f"Erro no arquivo PDF: {e}") from e
    except ConversionException as e:
        raise RuntimeError(f"Erro interno na conversão: {e}") from e
    except Exception as e:
        raise RuntimeError(f"Erro ao converter {file.filename} para Word: {e}") from e
    finally:
        if cv:
            cv.close()

    return [docx_path]


def pdf_to_text(file, temp_dir):
    """Extrai texto selecionável de um PDF e retorna um arquivo TXT."""
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(temp_dir, filename)
    file.save(pdf_path)

    base_name = os.path.splitext(filename)[0]
    txt_path = os.path.join(temp_dir, f"{base_name}.txt")

    with fitz.open(pdf_path) as document, open(txt_path, "w", encoding="utf-8") as output:
        for page_index, page in enumerate(document):
            if page_index:
                output.write("\n\n")
            output.write(f"--- Página {page_index + 1} ---\n")
            output.write(page.get_text("text"))

    return [txt_path]


def ocr_pdf(file, temp_dir):
    """
    Extrai texto de um PDF ou imagem usando Tesseract OCR.
    Retorna um arquivo TXT com o texto extraído.
    """
    filename = secure_filename(file.filename)
    input_path = os.path.join(temp_dir, filename)
    file.save(input_path)

    ext = filename.rsplit(".", 1)[1].lower()
    extracted_text = []

    if ext == "pdf":
        with fitz.open(input_path) as doc:
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x resolução
                img_path = os.path.join(temp_dir, f"ocr_page_{page_num + 1}.png")
                pix.save(img_path)

                with Image.open(img_path) as img:
                    text = pytesseract.image_to_string(img, lang="por+eng")
                extracted_text.append(f"--- Página {page_num + 1} ---\n{text}")
    elif ext in ("jpg", "jpeg", "png"):
        # Aplicar OCR diretamente na imagem
        with Image.open(input_path) as img:
            text = pytesseract.image_to_string(img, lang="por+eng")
        extracted_text.append(text)
    else:
        raise RuntimeError(f"Formato não suportado para OCR: {ext}")

    # Salvar texto extraído em arquivo TXT
    base_name = os.path.splitext(filename)[0]
    txt_path = os.path.join(temp_dir, f"{base_name}_ocr.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(extracted_text))

    return [txt_path]


def build_response(output_files, temp_dir):
    """Monta resposta enviando arquivos como attachment sem risco de remoção prematura do diretório temporário."""
    if len(output_files) == 1:
        file_path = output_files[0]
        filename = os.path.basename(file_path)
        with open(file_path, "rb") as f:
            data = f.read()
        return send_file(io.BytesIO(data), as_attachment=True, download_name=filename)
    else:
        zip_path = os.path.join(temp_dir, "converted_files.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file_path in output_files:
                zipf.write(file_path, os.path.basename(file_path))
        with open(zip_path, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data), as_attachment=True, download_name="converted_files.zip"
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
