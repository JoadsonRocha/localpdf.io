import base64
from datetime import date, datetime
import io
import json
import os
import pathlib
import shutil
import sys
import tempfile
from xml.sax.saxutils import escape
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
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from werkzeug.utils import secure_filename

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import ghostscript
except (ImportError, RuntimeError):
    ghostscript = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max
app.config["UPLOAD_FOLDER"] = os.path.join(tempfile.gettempdir(), "localpdf_uploads")
app.config["OUTPUT_FOLDER"] = os.path.join(tempfile.gettempdir(), "localpdf_outputs")

ALLOWED_EXTENSIONS = {"pdf", "docx", "txt", "xlsx", "jpg", "jpeg", "png"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── SEO Configuration & Structured Data ─────────────────────────────────────
SEO_CONFIG = {
    "home": {
        "title": "LocalPDF.io — Ferramentas de PDF 100% Privadas, Grátis e Ilimitadas",
        "description": "Converta, junte, divida, comprima e edite arquivos PDF gratuitamente no seu navegador ou PC. Sem limite de tamanho, sem cadastro e com privacidade absoluta.",
        "keywords": "pdf gratis, converter pdf, juntar pdf, comprimir pdf, editar pdf, mesclar pdf, pdf seguro, localpdf",
        "canonical": "/",
    },
    "editor": {
        "title": "Editor de PDF Online Grátis — Organizar, Girar e Reordenar Páginas | LocalPDF.io",
        "description": "Edite a estrutura do seu PDF visualmente: reordene páginas, gire, duplique e exclua sem perder formatação. 100% privado e gratuito.",
        "keywords": "editor de pdf, editar pdf online, organizar paginas pdf, girar pdf, excluir paginas pdf, reordenar pdf",
        "canonical": "/editor",
        "faq": [
            {
                "q": "Como editar e reordenar as páginas de um PDF?",
                "a": "Faça upload do seu PDF no Editor LocalPDF.io, arraste as páginas para a posição desejada, gire ou remova páginas e clique em Salvar."
            },
            {
                "q": "O editor altera a qualidade do arquivo original?",
                "a": "Não, as páginas são reorganizadas preservando integralmente o texto, fontes e imagens originais."
            }
        ]
    },
    "about": {
        "title": "Sobre o LocalPDF.io — Suíte de PDF com Privacidade Absoluta",
        "description": "Conheça o LocalPDF.io: ferramentas de PDF criadas com foco total em privacidade. Modo nuvem com processamento efêmero e app desktop 100% offline.",
        "keywords": "sobre localpdf, pdf privado, privacidade documentos, seguranca pdf",
        "canonical": "/about",
    },
    "pdf-to-jpg": {
        "title": "PDF para JPG Online e Grátis — Converter Páginas em Imagens JPG | LocalPDF.io",
        "description": "Converta páginas de documentos PDF em imagens JPG de alta resolução gratuitamente. Rápido, seguro e sem limite de páginas.",
        "keywords": "pdf para jpg, converter pdf em jpg, transformar pdf em imagem, pdf to jpg online gratis",
        "canonical": "/tool/pdf-to-jpg",
        "faq": [
            {"q": "Como converter PDF para JPG grátis?", "a": "Arraste seu arquivo PDF para a área de upload e clique em Converter para baixar suas imagens JPG compactadas."},
            {"q": "As imagens JPG convertidas mantêm boa resolução?", "a": "Sim, a conversão é realizada com renderização nítida mantendo excelente legibilidade de textos e fotos."}
        ]
    },
    "pdf-to-png": {
        "title": "PDF para PNG em Alta Resolução — Converter PDF Grátis | LocalPDF.io",
        "description": "Converta páginas do seu arquivo PDF em imagens PNG nítidas com máxima fidelidade. Rápido, privado e sem limite de uso.",
        "keywords": "pdf para png, converter pdf em png, extrair imagens pdf alta resolucao",
        "canonical": "/tool/pdf-to-png",
        "faq": [
            {"q": "Qual a diferença entre converter PDF para PNG e para JPG?", "a": "O formato PNG oferece compressão sem perdas (lossless), sendo ideal para diagramas, textos finos e gráficos que exigem máxima nitidez."},
            {"q": "Posso converter várias páginas de uma vez?", "a": "Sim, todas as páginas são convertidas e disponibilizadas para download em alta resolução."}
        ]
    },
    "pdf-to-images": {
        "title": "PDF para Imagens — Extrair Todas as Páginas em Imagens | LocalPDF.io",
        "description": "Converta e extraia cada página do seu documento PDF em imagens individuais de forma rápida e segura.",
        "keywords": "pdf para imagens, extrair imagens de pdf, transformar paginas pdf em imagens",
        "canonical": "/tool/pdf-to-images",
    },
    "images-to-pdf": {
        "title": "Imagens para PDF — Converter JPG e PNG em PDF Online | LocalPDF.io",
        "description": "Junte várias fotos e imagens (JPG, PNG) em um único arquivo PDF organizado. Rápido, privado e sem cadastro.",
        "keywords": "imagens para pdf, converter jpg em pdf, transformar fotos em pdf, jpg to pdf online",
        "canonical": "/tool/images-to-pdf",
        "faq": [
            {"q": "Posso juntar fotos de formatos diferentes em um mesmo PDF?", "a": "Sim, você pode combinar simultaneamente arquivos JPG, JPEG e PNG em um único documento PDF."},
            {"q": "As imagens perdem qualidade ao virar PDF?", "a": "Não, as imagens são incorporadas preservando a proporção e resolução original."}
        ]
    },
    "merge-pdf": {
        "title": "Mesclar PDFs Online Grátis — Juntar Vários PDFs em Um | LocalPDF.io",
        "description": "Junte múltiplos arquivos PDF em um único documento em segundos. Fácil, ilimitado, seguro e 100% gratuito.",
        "keywords": "mesclar pdf, juntar pdf, combinar pdf, unir pdfs gratis, merge pdf online",
        "canonical": "/tool/merge-pdf",
        "faq": [
            {"q": "Como mesclar vários PDFs em um só arquivo?", "a": "Selecione ou arraste os arquivos PDF que deseja juntar e clique no botão Converter para gerar um PDF único."},
            {"q": "Existe limite de tamanho ou quantidade de arquivos para mesclar?", "a": "Não há limites arbitrários; processamos arquivos de até 100MB com velocidade máxima no navegador."}
        ]
    },
    "split-pdf": {
        "title": "Dividir PDF Online — Extrair Páginas ou Separar PDF | LocalPDF.io",
        "description": "Divida seu arquivo PDF ou extraia páginas específicas com total facilidade e segurança. 100% gratuito e privado.",
        "keywords": "dividir pdf, separar pdf, extrair paginas pdf, split pdf online",
        "canonical": "/tool/split-pdf",
        "faq": [
            {"q": "Como extrair apenas algumas páginas do meu PDF?", "a": "Faça upload do documento, informe as páginas desejadas (ex: 1-3, 5) nas opções da ferramenta e clique em Converter."},
            {"q": "Posso dividir o PDF inteiro em páginas individuais?", "a": "Sim, basta deixar o campo de páginas em branco para extrair todas as páginas em arquivos individuais."}
        ]
    },
    "compress-pdf": {
        "title": "Comprimir PDF Online — Reduzir Tamanho de PDF Grátis | LocalPDF.io",
        "description": "Diminua o tamanho de arquivos PDF pesados mantendo excelente qualidade visual. Rápido, sem filas e com privacidade.",
        "keywords": "comprimir pdf, reduzir tamanho pdf, diminuir tamanho pdf, otimizar pdf, compress pdf online",
        "canonical": "/tool/compress-pdf",
        "faq": [
            {"q": "Como reduzir o tamanho de um arquivo PDF sem perder qualidade?", "a": "Envie seu arquivo e o algoritmo de compressão otimiza fluxos de dados e imagens internas, gerando um PDF mais leve e perfeitamente legível."},
            {"q": "Meus documentos confidenciais ficam seguros ao comprimir?", "a": "Sim! No LocalPDF.io o processamento é feito em memória volátil com exclusão imediata, ou 100% offline no app desktop."}
        ]
    },
    "protect-pdf": {
        "title": "Proteger PDF com Senha — Criptografar Documento PDF | LocalPDF.io",
        "description": "Adicione senha e proteção criptográfica ao seu PDF para impedir acesso não autorizado. Seguro, rápido e gratuito.",
        "keywords": "proteger pdf, colocar senha em pdf, criptografar pdf, bloquear pdf, protect pdf",
        "canonical": "/tool/protect-pdf",
        "faq": [
            {"q": "Como colocar senha em um arquivo PDF?", "a": "Faça o upload do documento, digite a senha desejada com pelo menos 4 caracteres e clique em Converter."},
            {"q": "A senha digitada fica salva em algum servidor?", "a": "Nunca. A senha é utilizada exclusivamente em memória temporária durante a cifragem e é descartada imediatamente."}
        ]
    },
    "unlock-pdf": {
        "title": "Desbloquear PDF — Remover Senha e Proteção de PDF | LocalPDF.io",
        "description": "Remova restrições e senha do seu arquivo PDF conhecido para facilitar edição e impressão. 100% privado.",
        "keywords": "desbloquear pdf, remover senha pdf, tirar senha de pdf, unlock pdf online",
        "canonical": "/tool/unlock-pdf",
        "faq": [
            {"q": "Como remover a senha de um PDF?", "a": "Envie o documento protegido, insira a senha atual do arquivo para autorizar a liberação e gere uma versão desprotegida."}
        ]
    },
    "watermark-pdf": {
        "title": "Marca d'Água em PDF — Adicionar Texto Personalizado | LocalPDF.io",
        "description": "Insira marcas d'água de texto (Confidencial, Rascunho, Cópia) em todas as páginas do seu PDF com facilidade.",
        "keywords": "marca dagua em pdf, adicionar marca dagua pdf, carimbo pdf, watermark pdf",
        "canonical": "/tool/watermark-pdf",
    },
    "page-numbers-pdf": {
        "title": "Numerar Páginas de PDF — Adicionar Numeração ao PDF | LocalPDF.io",
        "description": "Insira números de página automaticamente em documentos PDF. Escolha a posição e formate seu documento profissionalmente.",
        "keywords": "numerar paginas pdf, numero de pagina pdf, adicionar numeracao pdf, page numbers pdf",
        "canonical": "/tool/page-numbers-pdf",
    },
    "pdf-to-pdfa": {
        "title": "PDF para PDF/A — Conversão para Arquivamento a Longo Prazo | LocalPDF.io",
        "description": "Converta seus documentos PDF para o formato de conformidade PDF/A-1b para arquivamento legal e histórico.",
        "keywords": "pdf para pdfa, converter pdf em pdf/a, conformidade pdfa, arquivamento digital pdf",
        "canonical": "/tool/pdf-to-pdfa",
    },
    "word-to-pdf": {
        "title": "Word para PDF Online Grátis — Converter DOCX em PDF | LocalPDF.io",
        "description": "Converta arquivos Word (.docx) em documentos PDF com layout idêntico e formatação preservada. Grátis e sem limites.",
        "keywords": "word para pdf, converter docx em pdf, transformar word em pdf, docx to pdf",
        "canonical": "/tool/word-to-pdf",
        "faq": [
            {"q": "A formatação do documento Word é preservada ao converter para PDF?", "a": "Sim, fontes, margens, tabelas e imagens são mantidas fiéis ao layout original."},
            {"q": "Posso converter múltiplos documentos Word de uma só vez?", "a": "Sim, a ferramenta suporta envio de múltiplos arquivos DOCX simultâneos."}
        ]
    },
    "excel-to-pdf": {
        "title": "Excel para PDF Online — Converter Planilhas XLSX em PDF | LocalPDF.io",
        "description": "Converta planilhas Excel (.xlsx) para PDF com tabelas organizadas e dados protegidos. Rápido e confidencial.",
        "keywords": "excel para pdf, converter xlsx em pdf, transformar planilha em pdf, excel to pdf",
        "canonical": "/tool/excel-to-pdf",
    },
    "pdf-to-excel": {
        "title": "PDF para Excel — Extrair Tabelas de PDF para XLSX | LocalPDF.io",
        "description": "Extraia dados e tabelas do seu PDF diretamente para planilhas editáveis do Excel (.xlsx). Sem redigitação.",
        "keywords": "pdf para excel, extrair tabelas de pdf, converter pdf em xlsx, pdf to excel",
        "canonical": "/tool/pdf-to-excel",
    },
    "txt-to-pdf": {
        "title": "TXT para PDF — Converter Texto Simples em PDF | LocalPDF.io",
        "description": "Converta arquivos de texto (.txt) em documentos PDF formatados, limpos e prontos para impressão ou compartilhamento.",
        "keywords": "txt para pdf, converter texto em pdf, bloco de notas para pdf, txt to pdf",
        "canonical": "/tool/txt-to-pdf",
    },
    "pdf-to-word": {
        "title": "PDF para Word Editável — Converter PDF em DOCX Grátis | LocalPDF.io",
        "description": "Converta documentos PDF em arquivos Word (.docx) totalmente editáveis mantendo textos, parágrafos e imagens intactos.",
        "keywords": "pdf para word, converter pdf em docx, pdf editavel no word, pdf to word converter",
        "canonical": "/tool/pdf-to-word",
        "faq": [
            {"q": "O texto do PDF fica editável no Word após a conversão?", "a": "Sim, a ferramenta converte elementos de texto e imagens para o formato nativo do Word, permitindo edição direta no Word ou Google Docs."}
        ]
    },
    "pdf-to-text": {
        "title": "PDF para Texto — Extrair Texto Puro de Documentos PDF | LocalPDF.io",
        "description": "Extraia todo o conteúdo textual legível do seu PDF para um arquivo TXT limpo. Rápido, leve e prático.",
        "keywords": "pdf para texto, extrair texto de pdf, copiar texto do pdf, pdf to txt",
        "canonical": "/tool/pdf-to-text",
    },
    "ocr-pdf": {
        "title": "OCR em PDF — Reconhecimento Óptico de Texto em PDF Escaneado | LocalPDF.io",
        "description": "Reconheça e extraia texto de PDFs escaneados ou imagens digitalizadas usando inteligência de OCR. Grátis e privado.",
        "keywords": "ocr pdf, extrair texto de pdf escaneado, reconhecimento de caracteres pdf, pdf ocr online",
        "canonical": "/tool/ocr-pdf",
        "faq": [
            {"q": "O que é OCR em PDF?", "a": "OCR (Reconhecimento Óptico de Caracteres) é a tecnologia que analisa imagens e páginas digitalizadas para identificar letras e palavras, transformando-as em texto pesquisável e copiável."},
            {"q": "Quais tipos de arquivo o OCR aceita?", "a": "Você pode enviar arquivos PDF escaneados ou imagens diretas nos formatos JPG, PNG ou JPEG."}
        ]
    },
}


def get_seo_metadata(tool_key):
    key = tool_key if tool_key in SEO_CONFIG else "home"
    config = SEO_CONFIG[key]
    title = config["title"]
    desc = config["description"]
    keywords = config.get("keywords", "pdf, converter pdf, localpdf")
    canonical = f"https://localpdf.io{config['canonical']}"
    og_image = "https://localpdf.io/favicon.svg"

    graph = [
        {
            "@type": "WebApplication",
            "@id": "https://localpdf.io/#webapp",
            "name": "LocalPDF.io",
            "url": "https://localpdf.io",
            "applicationCategory": "OfficeApplication",
            "operatingSystem": "All (Web, Windows, macOS, Linux)",
            "offers": {
                "@type": "Offer",
                "price": "0",
                "priceCurrency": "BRL"
            },
            "aggregateRating": {
                "@type": "AggregateRating",
                "ratingValue": "4.9",
                "reviewCount": "1420",
                "bestRating": "5",
                "worstRating": "1"
            }
        },
        {
            "@type": "Organization",
            "@id": "https://localpdf.io/#organization",
            "name": "LocalPDF.io",
            "url": "https://localpdf.io",
            "logo": "https://localpdf.io/favicon.svg",
            "sameAs": [
                "https://github.com/JoadsonRocha/localpdf.io"
            ]
        }
    ]

    if key not in ("home", "about"):
        graph.append({
            "@type": "BreadcrumbList",
            "itemListElement": [
                {
                    "@type": "ListItem",
                    "position": 1,
                    "name": "Início",
                    "item": "https://localpdf.io/"
                },
                {
                    "@type": "ListItem",
                    "position": 2,
                    "name": title.split("—")[0].strip(),
                    "item": canonical
                }
            ]
        })

    if "faq" in config and config["faq"]:
        faq_items = [
            {
                "@type": "Question",
                "name": item["q"],
                "acceptedAnswer": {
                    "@type": "Answer",
                    "text": item["a"]
                }
            }
            for item in config["faq"]
        ]
        graph.append({
            "@type": "FAQPage",
            "mainEntity": faq_items
        })

    json_ld_str = json.dumps({"@context": "https://schema.org", "@graph": graph}, ensure_ascii=False, indent=2)

    return {
        "title": title,
        "description": desc,
        "keywords": keywords,
        "canonical_url": canonical,
        "og_image": og_image,
        "json_ld": json_ld_str,
    }


# Template HTML
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ seo.title if seo else 'LocalPDF.io' }}</title>
    <meta name="description" content="{{ seo.description if seo else 'Ferramentas de PDF 100% Privadas e Gratuitas' }}">
    <meta name="keywords" content="{{ seo.keywords if seo else 'pdf, localpdf, converter pdf' }}">
    <meta name="author" content="LocalPDF.io">
    <meta name="robots" content="index, follow">
    <link rel="canonical" href="{{ seo.canonical_url if seo else 'https://localpdf.io/' }}">
    <link rel="icon" type="image/svg+xml" href="/favicon.svg">

    <!-- Open Graph / Redes Sociais / WhatsApp -->
    <meta property="og:type" content="website">
    <meta property="og:site_name" content="LocalPDF.io">
    <meta property="og:title" content="{{ seo.title if seo else 'LocalPDF.io' }}">
    <meta property="og:description" content="{{ seo.description if seo else 'Ferramentas de PDF 100% Privadas e Gratuitas' }}">
    <meta property="og:url" content="{{ seo.canonical_url if seo else 'https://localpdf.io/' }}">
    <meta property="og:image" content="{{ seo.og_image if seo else 'https://localpdf.io/favicon.svg' }}">
    <meta property="og:locale" content="pt_BR">

    <!-- Twitter Card -->
    <meta name="twitter:card" content="summary_large_image">
    <meta name="twitter:title" content="{{ seo.title if seo else 'LocalPDF.io' }}">
    <meta name="twitter:description" content="{{ seo.description if seo else 'Ferramentas de PDF 100% Privadas e Gratuitas' }}">
    <meta name="twitter:image" content="{{ seo.og_image if seo else 'https://localpdf.io/favicon.svg' }}">

    {% if seo and seo.json_ld %}
    <!-- Schema.org JSON-LD Structured Data -->
    <script type="application/ld+json">
    {{ seo.json_ld | safe }}
    </script>
    {% endif %}
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
        .file-list { margin-top: 20px; display: grid; gap: 9px; text-align: left; }
        .file-item { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px 14px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 6px rgba(0,0,0,0.02); transition: all 0.2s ease; }
        .file-item:hover { border-color: #cbd5e1; box-shadow: 0 4px 12px rgba(0,0,0,0.04); }
        .file-info-group { display: flex; align-items: center; gap: 10px; min-width: 0; }
        .file-ext-badge { background: #eff6ff; color: #2563eb; font-weight: 800; font-size: 0.72rem; padding: 4px 7px; border-radius: 6px; letter-spacing: 0.04em; flex-shrink: 0; }
        .file-name-text { font-size: 0.92rem; font-weight: 600; color: #1e293b; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px; }
        .file-size-text { font-size: 0.8rem; color: #64748b; margin-left: 6px; flex-shrink: 0; }
        .file-remove-btn { background: #fef2f2; color: #ef4444; border: 1px solid #fee2e2; border-radius: 7px; padding: 6px 10px; font-size: 0.8rem; font-weight: 700; cursor: pointer; transition: all 0.2s ease; display: inline-flex; align-items: center; gap: 4px; flex-shrink: 0; font-family: inherit; }
        .file-remove-btn:hover { background: #fee2e2; color: #dc2626; border-color: #fca5a5; }

        /* ── Progress bar ────────────────────────────────────────── */
        .progress-box { margin-top: 24px; padding: 18px 20px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; text-align: left; }
        .progress-bar-wrapper { width: 100%; height: 10px; background: #e2e8f0; border-radius: 999px; overflow: hidden; position: relative; margin-bottom: 10px; }
        .progress-bar { height: 100%; width: 0%; background: linear-gradient(90deg, #2563eb 0%, #3b82f6 50%, #60a5fa 100%); background-size: 200% 100%; animation: progressShimmer 2s infinite linear; border-radius: 999px; transition: width 0.5s cubic-bezier(0.4,0,0.2,1); }
        @keyframes progressShimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
        .progress-info { display: flex; justify-content: space-between; align-items: center; font-size: 0.88rem; color: #475569; font-weight: 600; }
        .progress-stage-badge { display: inline-block; background: #dbeafe; color: #1d4ed8; border-radius: 999px; padding: 2px 9px; font-size: 0.75rem; font-weight: 800; letter-spacing: 0.03em; margin-right: 6px; }
        .progress-timer { color: #2563eb; font-variant-numeric: tabular-nums; }

        /* ── Toast notifications ─────────────────────────────────── */
        #toast-container { position: fixed; bottom: 28px; right: 28px; z-index: 9999; display: flex; flex-direction: column; gap: 10px; pointer-events: none; }
        .toast { pointer-events: auto; min-width: 260px; max-width: 380px; padding: 14px 18px; border-radius: 12px; font-size: 0.9rem; font-weight: 600; display: flex; align-items: center; gap: 10px; box-shadow: 0 8px 28px rgba(0,0,0,0.14); animation: toastIn 0.3s cubic-bezier(0.34,1.56,0.64,1) forwards; }
        .toast.toast-out { animation: toastOut 0.25s ease forwards; }
        .toast.success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
        .toast.error { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
        .toast.info { background: #eff6ff; border: 1px solid #bfdbfe; color: #1d4ed8; }
        .toast-icon { font-size: 1.1rem; flex-shrink: 0; }
        .toast-msg { flex: 1; line-height: 1.4; }
        .toast-close { background: none; border: none; cursor: pointer; opacity: 0.5; font-size: 1rem; padding: 0 2px; color: inherit; font-family: inherit; }
        .toast-close:hover { opacity: 1; }
        @keyframes toastIn { from { opacity: 0; transform: translateX(40px) scale(0.92); } to { opacity: 1; transform: translateX(0) scale(1); } }
        @keyframes toastOut { from { opacity: 1; transform: translateX(0) scale(1); } to { opacity: 0; transform: translateX(40px) scale(0.92); } }

        /* ── File item thumbnail ─────────────────────────────────── */
        .file-thumb { width: 38px; height: 38px; border-radius: 6px; object-fit: cover; border: 1px solid #e2e8f0; background: #f1f5f9; flex-shrink: 0; display: block; }
        .file-thumb-placeholder { width: 38px; height: 38px; border-radius: 6px; background: linear-gradient(90deg, #f0f2f5 25%, #e8ecf0 50%, #f0f2f5 75%); background-size: 200% 100%; animation: skeleton-shimmer 1.4s infinite linear; flex-shrink: 0; }

        /* ── Skeleton loading & fade-in for home grid cards ──────── */
        @keyframes skeleton-shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
        .tools-grid .tool-card.skeleton-card { pointer-events: none; }
        .tools-grid .tool-card.skeleton-card .tool-icon,
        .tools-grid .tool-card.skeleton-card h3,
        .tools-grid .tool-card.skeleton-card p {
            background: linear-gradient(90deg, #f0f2f5 25%, #e8ecf0 50%, #f0f2f5 75%);
            background-size: 200% 100%; animation: skeleton-shimmer 1.4s infinite linear;
            border-radius: 6px; color: transparent !important;
        }
        .tools-grid .tool-card { opacity: 1; transform: translateY(0); transition: opacity 0.35s ease, transform 0.35s ease; }
        .tools-grid .tool-card.card-visible { opacity: 1; transform: translateY(0); }
        .tools-grid .tool-card.card-hidden { opacity: 0 !important; display: none !important; }

        .result-card { margin-top: 24px; padding: 22px; border-radius: 12px; text-align: left; animation: fadeIn 0.3s ease; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .result-card.success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
        .result-card.warning { background: #fffbeb; border: 1px solid #fde68a; color: #92400e; }
        .result-card.error { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
        .result-header { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
        .result-header h4 { font-size: 1.15rem; font-weight: 700; margin: 0; }
        .result-body p { font-size: 0.92rem; margin-bottom: 12px; opacity: 0.95; line-height: 1.5; }
        .result-filename { font-weight: 700; word-break: break-all; color: #0f172a; background: rgba(255,255,255,0.8); padding: 6px 12px; border-radius: 6px; display: inline-block; margin-bottom: 12px; border: 1px solid rgba(0,0,0,0.06); font-size: 0.9rem; }
        .saved-path-badge { display: flex; align-items: flex-start; gap: 8px; font-size: 0.88rem; color: #1e293b; background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 10px 14px; margin-bottom: 14px; word-break: break-all; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
        .result-actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }
        .btn-download-again { background: #16a34a; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 0.9rem; transition: background 0.2s; box-shadow: 0 4px 10px rgba(22,163,74,0.2); font-family: inherit; }
        .btn-download-again:hover { background: #15803d; }
        .btn-open-file { background: #2563eb; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 0.9rem; transition: background 0.2s; box-shadow: 0 4px 10px rgba(37,99,235,0.2); font-family: inherit; }
        .btn-open-file:hover { background: #1d4ed8; }
        .btn-open-folder { background: #475569; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 0.9rem; transition: background 0.2s; box-shadow: 0 4px 10px rgba(71,85,105,0.2); font-family: inherit; }
        .btn-open-folder:hover { background: #334155; }
        .btn-reset-flow { background: #ffffff; color: #334155; border: 1px solid #cbd5e1; padding: 10px 18px; border-radius: 8px; font-weight: 600; cursor: pointer; font-size: 0.9rem; transition: all 0.2s; font-family: inherit; }
        .btn-reset-flow:hover { background: #f8fafc; border-color: #94a3b8; }
        .btn-try-again { background: #dc2626; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 0.9rem; font-family: inherit; }
        .btn-try-again:hover { background: #b91c1c; }
        .hidden { display: none; }

        /* ── Lazy-loaded editor page placeholder ─────────────────── */
        .editor-page img[data-lazy-src] { background: linear-gradient(90deg, #f0f2f5 25%, #e8ecf0 50%, #f0f2f5 75%); background-size: 200% 100%; animation: skeleton-shimmer 1.4s infinite linear; }
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
        .header { color: #24272b; margin: 0 auto; padding: 34px 0 20px; max-width: 700px; }
        .header h1 { font-size: clamp(1.4rem, 2.6vw, 1.85rem); line-height: 1.25; letter-spacing: -0.025em; margin-bottom: 8px; font-weight: 700; }
        .header p { color: #6c7178; font-size: 0.98rem; line-height: 1.5; }
        .category-tabs { display: flex; flex-wrap: wrap; justify-content: center; gap: 8px; margin-bottom: 28px; }
        .category-tab { background: #fff; border: 1px solid #e3e5e8; border-radius: 999px; color: #656a70; padding: 8px 16px; font-size: 0.86rem; font-weight: 700; cursor: pointer; transition: all 0.2s ease; font-family: inherit; outline: none; }
        .category-tab:hover { border-color: #93c5fd; color: #1d4ed8; background: #eff6ff; }
        .category-tab.active { background: #2563eb; color: #fff; border-color: #2563eb; box-shadow: 0 4px 12px rgba(37,99,235,0.22); }
        .tools-grid { grid-template-columns: repeat(auto-fit, minmax(245px, 1fr)); gap: 14px; margin-bottom: 58px; }
        .tools-grid .tool-card {
            border: 1px solid #e3e5e8;
            border-radius: 10px;
            padding: 22px;
            text-align: left;
            box-shadow: 0 5px 18px rgba(36,39,43,0.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
            text-decoration: none;
            color: inherit;
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            position: relative;
            overflow: hidden;
            min-height: 142px;
            background: rgba(255,255,255,0.96);
            cursor: pointer;
        }
        .tools-grid .tool-card::before { content: ''; position: absolute; inset: 0 auto 0 0; width: 3px; background: #dbeafe; transition: background 0.2s ease; }
        .tools-grid .tool-card:hover { transform: translateY(-3px); border-color: #93c5fd; box-shadow: 0 12px 28px rgba(37,99,235,0.12); }
        .tools-grid .tool-card:hover::before { background: #2563eb; }
        .tools-grid .tool-icon { width: 44px; height: 44px; border-radius: 10px; display: flex; align-items: center; justify-content: center; margin-bottom: 12px; transition: transform 0.2s ease, box-shadow 0.2s ease; flex-shrink: 0; }
        .tools-grid .tool-card:hover .tool-icon { transform: scale(1.08); }
        .tools-grid .tool-icon svg { width: 22px; height: 22px; display: block; }
        .card-hidden-filter { display: none !important; }
        .tools-grid .tool-card h3 { color: #24272b; font-size: 1.05rem; margin-bottom: 8px; font-family: Georgia, "Times New Roman", serif; }
        .tools-grid .tool-card p { color: #747980; font-size: 0.9rem; line-height: 1.5; margin-bottom: 0; }
        .tools-grid .tool-card:last-child { border-color: #2563eb; box-shadow: 0 8px 24px rgba(37,99,235,0.14); }

        /* ── Active Tool View & Editor Workspaces ───────────────── */
        #tool-views, #editor-view {
            max-width: 860px;
            margin: 0 auto;
            padding: 8px 0 44px;
        }
        #tool-views .tool-card,
        #editor-view .tool-card {
            opacity: 1 !important;
            transform: none !important;
            cursor: default !important;
            display: block !important;
            text-align: center;
            background: #ffffff;
            border: 1px solid #e3e5e8;
            border-radius: 16px;
            padding: 38px 32px;
            box-shadow: 0 8px 30px rgba(36, 39, 43, 0.06);
            min-height: auto;
            overflow: visible;
        }
        #tool-views .tool-card::before,
        #editor-view .tool-card::before {
            display: none !important;
        }
        #tool-views .tool-card:hover,
        #editor-view .tool-card:hover {
            transform: none !important;
            box-shadow: 0 8px 30px rgba(36, 39, 43, 0.06) !important;
            border-color: #e3e5e8 !important;
        }
        #tool-views .tool-card h3,
        #editor-view .tool-card h3 {
            font-size: 1.55rem;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 8px;
            font-family: Georgia, "Times New Roman", serif;
        }
        #tool-views .tool-card p#tool-description,
        #editor-view .tool-card > p {
            font-size: 0.96rem;
            color: #64748b;
            margin-bottom: 24px;
            line-height: 1.5;
        }
        #tool-views .upload-area,
        #editor-view .upload-area {
            border: 2px dashed #cbd5e1;
            border-radius: 12px;
            padding: 38px 20px;
            background: #f8fafc;
            cursor: pointer;
            transition: all 0.2s ease;
            margin: 18px 0;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }
        #tool-views .upload-area:hover,
        #editor-view .upload-area:hover {
            border-color: #2563eb;
            background: #eff6ff;
        }
        #tool-views .upload-area p,
        #editor-view .upload-area p {
            font-size: 1rem;
            font-weight: 600;
            color: #334155;
            margin-bottom: 4px;
        }
        #tool-views .upload-btn,
        #editor-view .upload-btn {
            background: #2563eb;
            color: white;
            padding: 10px 24px;
            border: none;
            border-radius: 999px;
            font-size: 0.92rem;
            font-weight: 700;
            cursor: pointer;
            transition: background 0.2s ease, transform 0.15s ease;
            font-family: inherit;
            box-shadow: 0 4px 12px rgba(37,99,235,0.2);
        }
        #tool-views .upload-btn:hover,
        #editor-view .upload-btn:hover {
            background: #1d4ed8;
            transform: translateY(-1px);
        }
        .convert-btn {
            background: #2563eb;
            color: white;
            padding: 12px 36px;
            border: none;
            border-radius: 999px;
            cursor: pointer;
            font-size: 1.05rem;
            font-weight: 700;
            margin-top: 20px;
            transition: background 0.2s ease, transform 0.15s ease, box-shadow 0.2s ease;
            box-shadow: 0 4px 14px rgba(37,99,235,0.25);
            font-family: inherit;
        }
        .convert-btn:hover:not(:disabled) {
            background: #1d4ed8;
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(37,99,235,0.32);
        }
        .convert-btn:disabled {
            background: #cbd5e1;
            cursor: not-allowed;
            box-shadow: none;
        }
        .back-btn {
            background: #ffffff;
            color: #475569;
            border: 1px solid #cbd5e1;
            padding: 8px 18px;
            border-radius: 999px;
            cursor: pointer;
            font-weight: 600;
            font-size: 0.88rem;
            font-family: inherit;
            margin-bottom: 20px;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.04);
        }
        .back-btn:hover {
            background: #f1f5f9;
            color: #1e293b;
            border-color: #94a3b8;
        }
        .footer { color: #747980; border-top: 1px solid #e3e5e8; padding-top: 40px; margin-top: 48px; }
        .footer a { color: #2563eb; }
        .footer .social-icons a { color: #747980; }
        .footer-grid { display: grid; grid-template-columns: 1.4fr 1fr 1fr 1fr; gap: 28px; max-width: 960px; margin: 0 auto 32px; text-align: left; }
        .footer-block h4 { color: #24272b; margin-bottom: 10px; font-family: Georgia, "Times New Roman", serif; font-size: 0.95rem; }
        .footer-block p, .footer-block a { font-size: 0.875rem; line-height: 1.75; }
        .footer-block a { display: block; }
        .footer-credit { border-top: 1px solid #e3e5e8; padding-top: 18px; margin-top: 4px; text-align: center; font-size: 0.82rem; color: #9ca3af; }
        .footer-credit a { color: #2563eb; text-decoration: none; }
        .footer-credit a:hover { text-decoration: underline; }
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
        .back-btn { background: #ffffff; color: #475569; border: 1px solid #cbd5e1; border-radius: 999px; font-weight: 600; padding: 8px 18px; font-size: 0.88rem; cursor: pointer; transition: all 0.2s ease; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.04); margin-bottom: 20px; }
        .back-btn:hover { background: #f1f5f9; color: #1e293b; border-color: #94a3b8; }
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
            .header { padding: 24px 0 16px; }
            .header h1 { font-size: 1.35rem; letter-spacing: -0.02em; }
            .header p { font-size: 0.92rem; line-height: 1.45; }
            .category-tabs { justify-content: flex-start; overflow-x: auto; flex-wrap: nowrap; margin: 0 -16px 20px; padding: 0 16px 5px; scrollbar-width: none; }
            .category-tabs::-webkit-scrollbar { display: none; }
            .category-tab { flex: 0 0 auto; font-size: 0.78rem; padding: 6px 12px; }
            .tools-grid { grid-template-columns: 1fr; gap: 10px; margin-bottom: 38px; }
            .tools-grid .tool-card { padding: 18px; }
            #tool-views .tool-card, #editor-view .tool-card { padding: 24px 16px; }
            .editor-shell { padding: 12px; }
            .editor-toolbar { gap: 8px; }
            .editor-toolbar button, .editor-actions button { flex: 1 1 100%; min-height: 42px; }
            .editor-pages { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 9px; }
            .editor-page { padding: 7px; }
            .editor-page-actions button { min-width: 0; padding: 8px 3px; }
            .editor-pending-pages { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .upload-area { padding: 26px 14px; }
            .footer { margin-top: 32px; padding-top: 28px; }
            .footer-grid {
                grid-template-columns: 1fr 1fr;
                gap: 0;
                text-align: left;
            }
            .footer-block {
                padding: 18px 16px;
                border-bottom: 1px solid #f0f2f4;
            }
            .footer-block:nth-child(odd) {
                border-right: 1px solid #f0f2f4;
            }
            .footer-block h4 { font-size: 0.88rem; margin-bottom: 7px; }
            .footer-block p, .footer-block a { font-size: 0.82rem; line-height: 1.65; }
            .footer-credit { margin-top: 8px; padding: 14px 16px; font-size: 0.78rem; text-align: center; }
        }

        /* About View Styles */
        .about-shell {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.04);
            margin-bottom: 40px;
            text-align: left;
        }
        .about-hero {
            text-align: center;
            max-width: 760px;
            margin: 0 auto 36px;
        }
        .about-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #ecfdf5;
            color: #059669;
            border: 1px solid #a7f3d0;
            padding: 5px 14px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 700;
            margin-bottom: 16px;
        }
        .about-hero h2 {
            font-family: Georgia, "Times New Roman", serif;
            font-size: 2.1rem;
            color: #1e293b;
            margin-bottom: 12px;
            letter-spacing: -0.02em;
        }
        .about-hero p {
            color: #64748b;
            font-size: 1.05rem;
            line-height: 1.6;
        }
        .about-badges-bar {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 8px;
            margin-top: 18px;
        }
        .tag-badge {
            background: #f1f5f9;
            color: #334155;
            padding: 4px 11px;
            border-radius: 6px;
            font-size: 0.78rem;
            font-weight: 600;
        }
        .about-grid-pillars {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
            gap: 18px;
            margin-bottom: 40px;
        }
        .about-card {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 22px;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .about-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.06);
            border-color: #cbd5e1;
        }
        .about-card-icon {
            font-size: 1.8rem;
            margin-bottom: 10px;
        }
        .about-card h3 {
            font-size: 1.05rem;
            color: #1e293b;
            margin-bottom: 8px;
            font-family: inherit;
            font-weight: 700;
        }
        .about-card p {
            font-size: 0.88rem;
            color: #64748b;
            line-height: 1.55;
            margin: 0;
        }
        .about-section-title {
            font-size: 1.3rem;
            color: #1e293b;
            margin: 36px 0 16px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
            font-family: Georgia, "Times New Roman", serif;
        }
        .about-table-wrapper {
            overflow-x: auto;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            margin-bottom: 36px;
        }
        .compare-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.88rem;
        }
        .compare-table th, .compare-table td {
            padding: 13px 18px;
            text-align: left;
            border-bottom: 1px solid #e2e8f0;
        }
        .compare-table th {
            background: #f1f5f9;
            color: #334155;
            font-weight: 700;
        }
        .compare-table tr:last-child td {
            border-bottom: none;
        }
        .compare-table td.brand-col {
            background: #eff6ff;
            color: #1d4ed8;
            font-weight: 600;
        }
        .about-tech-stack {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 14px;
        }
        .tech-pill {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            padding: 7px 13px;
            border-radius: 8px;
            font-size: 0.84rem;
            color: #334155;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        .tech-pill strong {
            color: #0f172a;
        }
        .tool-privacy-notice {
            margin-top: 18px;
            padding: 12px 16px;
            border-radius: 10px;
            font-size: 0.84rem;
            line-height: 1.5;
            text-align: left;
            display: flex;
            align-items: flex-start;
            gap: 10px;
        }
        .tool-privacy-notice.web-mode {
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1e3a8a;
        }
        .tool-privacy-notice.local-mode {
            background: #ecfdf5;
            border: 1px solid #a7f3d0;
            color: #065f46;
        }
        .tool-privacy-notice a {
            color: #1d4ed8;
            font-weight: 700;
            text-decoration: underline;
        }
        .web-mode-banner {
            background: linear-gradient(135deg, rgba(239, 246, 255, 0.95) 0%, rgba(240, 249, 255, 0.95) 100%);
            border: 1px solid #bfdbfe;
            border-radius: 999px;
            padding: 7px 14px 7px 16px;
            margin: 0 auto 24px auto;
            max-width: 960px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            font-size: 0.85rem;
            color: #1e3a8a;
            box-shadow: 0 2px 8px rgba(37, 99, 235, 0.05);
            transition: opacity 0.25s ease, transform 0.25s ease;
        }
        .web-mode-banner.dismissed {
            display: none !important;
        }
        .web-mode-banner-content {
            display: flex;
            align-items: center;
            gap: 8px;
            min-width: 0;
            flex: 1;
        }
        .web-mode-badge-pill {
            background: #2563eb;
            color: #ffffff;
            font-size: 0.72rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 999px;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            flex-shrink: 0;
        }
        .web-mode-banner-text {
            font-size: 0.85rem;
            color: #1e3a8a;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            line-height: 1.4;
        }
        .web-mode-banner-actions {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-shrink: 0;
        }
        .web-mode-cta-link {
            background: #0078D4;
            color: #ffffff !important;
            padding: 5px 12px;
            border-radius: 999px;
            font-weight: 600;
            font-size: 0.78rem;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            white-space: nowrap;
            transition: all 0.2s ease;
        }
        .web-mode-cta-link:hover {
            background: #0063b1;
            transform: translateY(-1px);
        }
        .web-mode-close-btn {
            background: transparent;
            border: none;
            color: #64748b;
            font-size: 1.2rem;
            line-height: 1;
            cursor: pointer;
            padding: 2px 6px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            transition: color 0.15s ease, background-color 0.15s ease;
        }
        .web-mode-close-btn:hover {
            color: #0f172a;
            background-color: rgba(0, 0, 0, 0.06);
        }
        .nav-desktop-badge {
            background: #0078D4;
            color: #ffffff !important;
            padding: 5px 12px;
            border-radius: 6px;
            font-weight: 700;
            font-size: 0.8rem;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            text-decoration: none;
            transition: background 0.2s ease;
        }
        .nav-desktop-badge:hover {
            background: #0063b1;
        }
        @media (max-width: 700px) {
            .about-shell { padding: 22px 16px; }
            .about-hero h2 { font-size: 1.5rem; }
            .about-grid-pillars { grid-template-columns: 1fr; }
            .web-mode-banner { border-radius: 14px; padding: 10px 14px; flex-wrap: wrap; }
            .web-mode-banner-text { white-space: normal; }
            .web-mode-banner-actions { width: 100%; justify-content: space-between; margin-top: 4px; }
        }
    </style>
</head>
<body>
    <div id="toast-container"></div>
    <div class="container">
        <nav class="site-nav">
            <a class="site-brand" href="/" onclick="showHome(); return false;">local<span>pdf</span><small>.io</small></a>
            <div class="site-nav-links">
                <a href="/tool/merge-pdf" onclick="showTool('merge-pdf'); return false;">Juntar PDF</a>
                <a href="/tool/split-pdf" onclick="showTool('split-pdf'); return false;">Dividir PDF</a>
                <a href="/tool/compress-pdf" onclick="showTool('compress-pdf'); return false;">Comprimir PDF</a>
                <a href="/tool/pdf-to-word" onclick="showTool('pdf-to-word'); return false;">Converter PDF</a>
                <a href="/about" onclick="showAbout(); return false;">Sobre</a>
                <a href="/" onclick="showHome(); return false;">Todas as ferramentas</a>
                <a id="nav-desktop-download" class="nav-desktop-badge hidden" href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" rel="noopener" title="Processamento 100% offline no seu computador">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor"><path d="M0 3.449L9.75 2.1v9.451H0m10.949-9.602L24 0v11.4H10.949M0 12.6h9.75v9.451L0 20.699M10.949 12.6H24V24l-12.949-1.801"/></svg>
                    <span>Baixar para Windows</span>
                </a>
                <button id="language-toggle" class="language-toggle" type="button" onclick="toggleLanguage()">EN</button>
            </div>
        </nav>

        <div id="home-view">
            <!-- Banner informativo para ambiente Web / Railway -->
            <div id="env-banner" class="web-mode-banner hidden">
                <div class="web-mode-banner-content">
                    <span class="web-mode-badge-pill">Nuvem Segura</span>
                    <span class="web-mode-banner-text" id="env-banner-text">Arquivos processados em memória volátil e excluídos imediatamente após o download.</span>
                </div>
                <div class="web-mode-banner-actions">
                    <a class="web-mode-cta-link" href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" rel="noopener">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M0 3.449L9.75 2.1v9.451H0m10.949-9.602L24 0v11.4H10.949M0 12.6h9.75v9.451L0 20.699M10.949 12.6H24V24l-12.949-1.801"/></svg>
                        <span id="env-banner-cta-text">App Desktop (100% Local)</span>
                    </a>
                    <button type="button" class="web-mode-close-btn" onclick="dismissWebBanner()" title="Fechar aviso" aria-label="Fechar aviso">&times;</button>
                </div>
            </div>

            <div class="header">
                <h1>Trabalhe com seus PDFs sem complicação</h1>
                <p id="main-subtitle">Converta, organize e edite documentos diretamente no seu computador. Sem contas, sem nuvem e sem enviar seus arquivos para fora.</p>
            </div>
            <div id="tools" class="category-tabs" role="tablist" aria-label="Categorias de ferramentas">
                <button type="button" class="category-tab active" data-category="all" onclick="filterCategory('all', this)" role="tab" aria-selected="true">Todas</button>
                <button type="button" class="category-tab" data-category="organizar" onclick="filterCategory('organizar', this)" role="tab" aria-selected="false">Organizar PDF</button>
                <button type="button" class="category-tab" data-category="converter" onclick="filterCategory('converter', this)" role="tab" aria-selected="false">Converter PDF</button>
                <button type="button" class="category-tab" data-category="otimizar" onclick="filterCategory('otimizar', this)" role="tab" aria-selected="false">Otimizar PDF</button>
                <button type="button" class="category-tab" data-category="ocr" onclick="filterCategory('ocr', this)" role="tab" aria-selected="false">OCR</button>
            </div>
            <div class="tools-grid">
                <a class="tool-card" href="/tool/pdf-to-jpg" onclick="showTool('pdf-to-jpg'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #eff6ff; color: #2563eb;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><circle cx="10" cy="13" r="1.5"/><path d="m8 18 3-3 2 2 3-4 2 2"/></svg>
                    </div>
                    <h3>PDF para JPG</h3>
                    <p>Converta páginas PDF em imagens JPG compactadas</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-png" onclick="showTool('pdf-to-png'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #f0fdf4; color: #16a34a;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="m21 15-5-5L5 21"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/></svg>
                    </div>
                    <h3>PDF para PNG</h3>
                    <p>Converta páginas PDF em imagens PNG em alta definição</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-images" onclick="showTool('pdf-to-images'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #fdf4ff; color: #a21caf;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>
                    </div>
                    <h3>PDF para Imagens</h3>
                    <p>Converta páginas PDF em imagens separadas</p>
                </a>
                <a class="tool-card" href="/tool/images-to-pdf" onclick="showTool('images-to-pdf'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #f0fdf4; color: #16a34a;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="m21 15-5-5L5 21"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/></svg>
                    </div>
                    <h3>Imagens para PDF</h3>
                    <p>Combine várias imagens em um único PDF</p>
                </a>
                <a class="tool-card" href="/tool/merge-pdf" onclick="showTool('merge-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #fef2f2; color: #dc2626;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2H4a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h4"/><path d="M16 2h4a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2h-4"/><path d="M12 7v10"/><path d="m9 10 3-3 3 3"/><path d="m9 14 3 3 3-3"/></svg>
                    </div>
                    <h3>Mesclar PDFs</h3>
                    <p>Combine vários PDFs em um documento único</p>
                </a>
                <a class="tool-card" href="/tool/split-pdf" onclick="showTool('split-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #f5f3ff; color: #7c3aed;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/></svg>
                    </div>
                    <h3>Dividir PDF</h3>
                    <p>Extraia todas ou páginas específicas por intervalo</p>
                </a>
                <a class="tool-card" href="/tool/compress-pdf" onclick="showTool('compress-pdf'); return false;" data-category="otimizar">
                    <div class="tool-icon" style="background: #ecfdf5; color: #059669;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14h6m0 0v6m0-6L3 21"/><path d="M20 10h-6m0 0V4m0 6 7-7"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                    </div>
                    <h3>Comprimir PDF</h3>
                    <p>Reduza o tamanho do seu arquivo PDF</p>
                </a>
                <a class="tool-card" href="/tool/protect-pdf" onclick="showTool('protect-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #fffbeb; color: #d97706;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/><circle cx="12" cy="16" r="1.5"/></svg>
                    </div>
                    <h3>Proteger PDF</h3>
                    <p>Adicione uma senha local ao seu documento PDF</p>
                </a>
                <a class="tool-card" href="/tool/unlock-pdf" onclick="showTool('unlock-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #fef2f2; color: #ef4444;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/><circle cx="12" cy="16" r="1.5"/></svg>
                    </div>
                    <h3>Desbloquear PDF</h3>
                    <p>Remova permanentemente a senha do seu PDF</p>
                </a>
                <a class="tool-card" href="/tool/watermark-pdf" onclick="showTool('watermark-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #f0f9ff; color: #0284c7;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"/><path d="M12 12a3 3 0 0 0 3-3"/></svg>
                    </div>
                    <h3>Marca d'água</h3>
                    <p>Adicione uma marca d'água de texto ao PDF</p>
                </a>
                <a class="tool-card" href="/tool/page-numbers-pdf" onclick="showTool('page-numbers-pdf'); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #f0fdfa; color: #0d9488;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M10 12h2v6m-2 0h4"/></svg>
                    </div>
                    <h3>Números de página</h3>
                    <p>Numere as páginas do documento localmente</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-pdfa" onclick="showTool('pdf-to-pdfa'); return false;" data-category="otimizar">
                    <div class="tool-icon" style="background: #f8fafc; color: #475569;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/></svg>
                    </div>
                    <h3>PDF para PDF/A</h3>
                    <p>Padronize seu PDF para arquivamento (PDF/A)</p>
                </a>
                <a class="tool-card" href="/tool/word-to-pdf" onclick="showTool('word-to-pdf'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #eff6ff; color: #1d4ed8;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13l1.5 5 2-4 2 4 1.5-5"/></svg>
                    </div>
                    <h3>Word para PDF</h3>
                    <p>Converta um ou mais documentos DOCX para PDF</p>
                </a>
                <a class="tool-card" href="/tool/excel-to-pdf" onclick="showTool('excel-to-pdf'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #f0fdf4; color: #15803d;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><rect x="8" y="12" width="8" height="6"/><line x1="8" y1="15" x2="16" y2="15"/><line x1="12" y1="12" x2="12" y2="18"/></svg>
                    </div>
                    <h3>Excel para PDF</h3>
                    <p>Converta planilhas XLSX para PDF</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-excel" onclick="showTool('pdf-to-excel'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #ecfdf5; color: #059669;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
                    </div>
                    <h3>PDF para Excel</h3>
                    <p>Extraia tabelas do PDF para planilhas Excel (.xlsx) editáveis</p>
                </a>
                <a class="tool-card" href="/tool/txt-to-pdf" onclick="showTool('txt-to-pdf'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #f1f5f9; color: #475569;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/></svg>
                    </div>
                    <h3>TXT para PDF</h3>
                    <p>Converta arquivos de texto simples para PDF</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-word" onclick="showTool('pdf-to-word'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #eff6ff; color: #2563eb;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6H6a2 2 0 0 0-2 2z"/><polyline points="14 2 14 8 20 8"/><path d="M8 12h8m-8 4h5"/></svg>
                    </div>
                    <h3>PDF para Word</h3>
                    <p>Converta documentos PDF para Word (.docx) editável</p>
                </a>
                <a class="tool-card" href="/tool/pdf-to-text" onclick="showTool('pdf-to-text'); return false;" data-category="converter">
                    <div class="tool-icon" style="background: #fdf4ff; color: #a21caf;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/></svg>
                    </div>
                    <h3>PDF para Texto</h3>
                    <p>Extraia o texto do PDF para um arquivo TXT editável</p>
                </a>
                <a class="tool-card" href="/tool/ocr-pdf" onclick="showTool('ocr-pdf'); return false;" data-category="ocr">
                    <div class="tool-icon" style="background: #faf5ff; color: #9333ea;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/><line x1="3" y1="12" x2="21" y2="12" stroke-dasharray="2 2"/><circle cx="12" cy="12" r="3"/></svg>
                    </div>
                    <h3>OCR em PDF</h3>
                    <p>Extraia texto de PDFs e imagens escaneadas com OCR</p>
                </a>
                <a class="tool-card" href="/editor" onclick="showEditor(); return false;" data-category="organizar">
                    <div class="tool-icon" style="background: #fff1f2; color: #e11d48;">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                    </div>
                    <h3>Editar PDF</h3>
                    <p>Reordene, insira, gire, duplique e exclua páginas diretamente no PDF</p>
                </a>
            </div>
        </div>

        <div id="editor-view" class="hidden">
            <button class="back-btn" onclick="navigateBack()">← Voltar</button>
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
            <button class="back-btn" onclick="navigateBack()">← Voltar</button>
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

                <div id="progress" class="progress-box hidden">
                    <div class="progress-bar-wrapper">
                        <div id="progress-bar" class="progress-bar"></div>
                    </div>
                    <div class="progress-info">
                        <span id="progress-message">Processando documento localmente...</span>
                        <span id="progress-timer">⏱️ 00:00</span>
                    </div>
                </div>

                <div id="result" class="hidden"></div>
                <div id="tool-privacy-badge" class="tool-privacy-notice"></div>
            </div>
        </div>

        <div id="about-view" class="hidden">
            <button class="back-btn" onclick="navigateBack()">← Voltar</button>
            <div class="about-shell">
                <div class="about-hero">
                    <span class="about-badge">🛡️ 100% Privado &amp; Local</span>
                    <h2>Sobre o LocalPDF.io</h2>
                    <p>O LocalPDF.io nasceu com uma missão simples: fornecer um pacote profissional e completo de ferramentas para PDF diretamente no seu computador, com velocidade máxima, sem custos, sem filas e com privacidade absoluta.</p>
                    <div class="about-badges-bar">
                        <span class="tag-badge">Zero Nuvem</span>
                        <span class="tag-badge">Sem Upload Externo</span>
                        <span class="tag-badge">Funciona Offline</span>
                        <span class="tag-badge">Sem Limite de Páginas</span>
                        <span class="tag-badge">Instalador Windows (.msi)</span>
                        <span class="tag-badge">Código Aberto MIT</span>
                    </div>
                </div>

                <div class="about-grid-pillars">
                    <div class="about-card">
                        <div class="about-card-icon">🔒</div>
                        <h3>Privacidade em Primeiro Lugar</h3>
                        <p>Diferente de serviços web tradicionais que exigem o upload de seus dados confidenciais para servidores remotos, o LocalPDF executa cada conversão, corte e junção exclusivamente na memória da sua própria máquina.</p>
                    </div>
                    <div class="about-card">
                        <div class="about-card-icon">⚡</div>
                        <h3>Desempenho Ilimitado</h3>
                        <p>Sem restrições de tamanho máximo de arquivo (15MB/50MB), sem contadores de tarefas diárias e sem esperas artificiais. Se o seu computador aguenta abrir o arquivo, o LocalPDF consegue processá-lo.</p>
                    </div>
                    <div class="about-card">
                        <div class="about-card-icon">📴</div>
                        <h3>Totalmente Offline</h3>
                        <p>Leve seu trabalho para viagens, locais remotos ou ambientes corporativos isolados sem internet. O aplicativo desktop não requer conexão e funciona de forma 100% autônoma.</p>
                    </div>
                    <div class="about-card">
                        <div class="about-card-icon">💼</div>
                        <h3>Conformidade LGPD &amp; GDPR</h3>
                        <p>Como nenhum arquivo, dado cadastral ou metadado trafega pela rede, sua empresa e seus clientes desfrutam de conformidade nativa com as mais rigorosas leis de proteção de dados.</p>
                    </div>
                </div>

                <h3 class="about-section-title">🌐 Versão Web (Railway) vs. 💻 Aplicativo Desktop Local</h3>
                <p style="color: #64748b; font-size: 0.95rem; line-height: 1.6; margin-bottom: 20px;">
                    O LocalPDF está disponível em duas modalidades para atender perfeitamente a diferentes necessidades de uso e níveis de confidencialidade:
                </p>
                <div class="about-grid-pillars" style="grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));">
                    <div class="about-card" style="border-top: 4px solid #0078D4; background: #f0f7ff;">
                        <div class="about-card-icon">💻</div>
                        <h3 style="color: #0078D4;">Aplicativo Desktop (Windows .msi)</h3>
                        <span class="tag-badge" style="background: #dbeafe; color: #1e40af; margin-bottom: 12px; display: inline-block;">Recomendado para Máxima Privacidade</span>
                        <ul style="margin-left: 18px; color: #475569; font-size: 0.88rem; line-height: 1.7;">
                            <li><strong>100% Offline:</strong> Funciona sem precisar de conexão à internet.</li>
                            <li><strong>Zero Envio de Dados:</strong> Seus arquivos nunca saem do seu computador ou rede local.</li>
                            <li><strong>Sem Limites de Servidor:</strong> Processa arquivos grandes aproveitando a CPU e RAM do seu PC.</li>
                            <li><strong>Instalador Nativo:</strong> Pacote MSI limpo, assinado digitalmente, com atalhos no Menu Iniciar.</li>
                        </ul>
                        <div style="margin-top: 16px;">
                            <a class="web-mode-cta-btn" style="display: inline-flex;" href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" rel="noopener">
                                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M0 3.449L9.75 2.1v9.451H0m10.949-9.602L24 0v11.4H10.949M0 12.6h9.75v9.451L0 20.699M10.949 12.6H24V24l-12.949-1.801"/></svg>
                                <span>Baixar Instalador Windows (.msi)</span>
                            </a>
                        </div>
                    </div>

                    <div class="about-card" style="border-top: 4px solid #10b981; background: #f0fdf4;">
                        <div class="about-card-icon">🌐</div>
                        <h3 style="color: #059669;">Versão Web (Nuvem Railway)</h3>
                        <span class="tag-badge" style="background: #dcfce7; color: #166534; margin-bottom: 12px; display: inline-block;">Praticidade Imediata no Navegador</span>
                        <ul style="margin-left: 18px; color: #475569; font-size: 0.88rem; line-height: 1.7;">
                            <li><strong>Sem Instalação:</strong> Acesse de qualquer dispositivo direto pelo navegador.</li>
                            <li><strong>Processamento Efêmero:</strong> O arquivo é enviado via HTTPS criptografado, processado em memória volátil e <em>excluído imediatamente</em> após o download.</li>
                            <li><strong>Zero Persistência:</strong> Não usamos banco de dados, não guardamos histórico e nenhum arquivo é mantido em disco.</li>
                            <li><strong>Totalmente Gratuito:</strong> Todas as 18 ferramentas disponíveis sem necessidade de cadastro.</li>
                        </ul>
                    </div>
                </div>

                <h3 class="about-section-title">📊 Por que escolher o LocalPDF.io?</h3>
                <div class="about-table-wrapper">
                    <table class="compare-table">
                        <thead>
                            <tr>
                                <th>Recurso / Critério</th>
                                <th style="background: #dbeafe; color: #1e40af;">LocalPDF Desktop (Local)</th>
                                <th style="background: #ecfdf5; color: #065f46;">LocalPDF Web (Railway)</th>
                                <th>Ferramentas na Nuvem (iLovePDF, Smallpdf, etc.)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>Seus arquivos saem do seu computador?</strong></td>
                                <td class="brand-col">❌ Nunca (100% processamento local)</td>
                                <td style="background: #f0fdf4; color: #047857;">☁️ Trânsito temporário com exclusão imediata</td>
                                <td>⚠️ Sim (enviados para servidores de terceiros)</td>
                            </tr>
                            <tr>
                                <td><strong>Persistência ou retenção em disco</strong></td>
                                <td class="brand-col">🔒 Zero (apenas no seu PC)</td>
                                <td style="background: #f0fdf4; color: #047857;">❌ Zero (excluído imediatamente pós-download)</td>
                                <td>⚠️ Retidos por horas ou dias nos servidores</td>
                            </tr>
                            <tr>
                                <td><strong>Limite de tamanho de arquivo</strong></td>
                                <td class="brand-col">✅ Sem limite artificial (usa sua RAM)</td>
                                <td style="background: #f0fdf4; color: #047857;">✅ Até 100 MB por operação</td>
                                <td>❌ 15 MB a 50 MB (exigem plano pago)</td>
                            </tr>
                            <tr>
                                <td><strong>Funciona sem conexão à Internet?</strong></td>
                                <td class="brand-col">✅ Sim, totalmente offline</td>
                                <td style="background: #f0fdf4; color: #047857;">❌ Não (requer conexão web)</td>
                                <td>❌ Não funciona sem conexão</td>
                            </tr>
                            <tr>
                                <td><strong>Exige cadastro ou assinatura?</strong></td>
                                <td class="brand-col">✅ Totalmente livre e sem cadastro</td>
                                <td style="background: #f0fdf4; color: #047857;">✅ Livre e gratuito em todas as 18 ferramentas</td>
                                <td>❌ Planos mensais ou anúncios invasivos</td>
                            </tr>
                            <tr>
                                <td><strong>Segurança para dados sensíveis e contratos</strong></td>
                                <td class="brand-col">✅ Máxima segurança física no seu PC</td>
                                <td style="background: #f0fdf4; color: #047857;">🔒 Conexão HTTPS segura + exclusão automática</td>
                                <td>⚠️ Risco de vazamento em servidores na nuvem</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <h3 class="about-section-title">🛠️ Recursos e Ferramentas Integradas</h3>
                <div class="about-grid-pillars">
                    <div class="about-card">
                        <h3>Organizar &amp; Editar</h3>
                        <p><strong>Mesclar PDFs:</strong> Una múltiplos documentos com ordenação flexível.<br><strong>Dividir PDF:</strong> Extraia todas as páginas ou intervalos personalizados (ex: 1-3, 5, 8-10).<br><strong>Editor Visual:</strong> Reordene, insira, gire, duplique e exclua páginas com visualização em miniaturas.</p>
                    </div>
                    <div class="about-card">
                        <h3>Converter de/para PDF</h3>
                        <p><strong>PDF para JPG / PNG:</strong> Extraia páginas com alta definição.<br><strong>PDF para Excel (.xlsx):</strong> Detecte e extraia tabelas em planilhas editáveis.<br><strong>PDF para Word (.docx):</strong> Converta mantendo o layout.<br><strong>OCR em PDF:</strong> Reconhecimento óptico de caracteres em documentos escaneados.</p>
                    </div>
                    <div class="about-card">
                        <h3>Segurança &amp; Otimização</h3>
                        <p><strong>Proteger PDF:</strong> Criptografe com senha forte AES-256.<br><strong>Desbloquear PDF:</strong> Remova senhas e restrições permanentemente.<br><strong>Comprimir PDF:</strong> Otimize e reduza o tamanho sem perder legibilidade.<br><strong>Marca d'Água:</strong> Aplique carimbos diagonais profissionais com transparência.</p>
                    </div>
                </div>

                <h3 class="about-section-title">⚡ Arquitetura &amp; Tecnologias</h3>
                <p style="color: #64748b; font-size: 0.95rem; line-height: 1.6;">O LocalPDF.io é desenvolvido em Python moderno, integrando as mais respeitadas bibliotecas de computação gráfica e manipulação de documentos do mundo:</p>
                <div class="about-tech-stack">
                    <span class="tech-pill"><strong>PyMuPDF (fitz)</strong> · Motor veloz C++ para PDF</span>
                    <span class="tech-pill"><strong>pdfplumber</strong> · Extração precisa de tabelas para Excel</span>
                    <span class="tech-pill"><strong>pdf2docx</strong> · Reconstrução inteligente de documentos Word</span>
                    <span class="tech-pill"><strong>ReportLab</strong> · Geração vetorial de PDFs e marca d'água</span>
                    <span class="tech-pill"><strong>Tesseract OCR</strong> · Reconhecimento óptico neural</span>
                    <span class="tech-pill"><strong>WiX Toolset v5</strong> · Empacotamento MSI assinado digitalmente</span>
                </div>
            </div>
        </div>

        <div id="privacy-note" class="footer">
            <div class="footer-grid">
                <div class="footer-block">
                    <h4>LocalPDF.io</h4>
                    <p id="footer-desc">Ferramentas PDF gratuitas, locais e privadas. Sem cadastro e sem upload externo no modo local.</p>
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
        let isWebMode = false;
        let uploadedFiles = [];
        let progressInterval = null;
        let progressSeconds = 0;
        let lastDownloadedBlob = null;
        let lastDownloadedFilename = '';
        let fileThumbnails = {}; // filename -> dataURL cache

        // ── Toast Notifications ──────────────────────────────────────
        function showToast(msg, type = 'success', duration = 4500) {
            const container = document.getElementById('toast-container');
            const icons = { success: '✅', error: '⚠️', info: 'ℹ️' };
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            toast.innerHTML = `
                <span class="toast-icon">${icons[type] || icons.info}</span>
                <span class="toast-msg">${msg}</span>
                <button class="toast-close" onclick="dismissToast(this.parentElement)">✕</button>
            `;
            container.appendChild(toast);
            const timer = setTimeout(() => dismissToast(toast), duration);
            toast._timer = timer;
        }
        function dismissToast(toast) {
            if (!toast || !toast.parentElement) return;
            clearTimeout(toast._timer);
            toast.classList.add('toast-out');
            toast.addEventListener('animationend', () => toast.remove(), { once: true });
        }

        // ── Thumbnail preview ────────────────────────────────────────
        async function generateThumbnail(file) {
            const key = `${file.name}-${file.size}`;
            if (fileThumbnails[key]) return fileThumbnails[key];
            const ext = file.name.split('.').pop().toLowerCase();
            if (['jpg','jpeg','png'].includes(ext)) {
                return new Promise(resolve => {
                    const reader = new FileReader();
                    reader.onload = e => { fileThumbnails[key] = e.target.result; resolve(e.target.result); };
                    reader.readAsDataURL(file);
                });
            }
            if (ext === 'pdf' && file.size <= 8 * 1024 * 1024) {
                try {
                    const fd = new FormData();
                    fd.append('file', file);
                    const res = await fetch('/preview-page', { method: 'POST', body: fd });
                    if (res.ok) {
                        const data = await res.json();
                        if (data.thumbnail) { fileThumbnails[key] = data.thumbnail; return data.thumbnail; }
                    }
                } catch(e) {}
            }
            return null;
        }

        function t(pt, en) {
            return currentLanguage === 'en' ? en : pt;
        }

        const languageTexts = {
            'Trabalhe com seus PDFs sem complicação': 'Work with your PDFs without the hassle',
            'Converta, organize e edite documentos diretamente no seu computador. Sem contas, sem nuvem e sem enviar seus arquivos para fora.': 'Convert, organize and edit documents directly on your computer. No accounts, no cloud and no files sent elsewhere.',
            'Juntar PDF': 'Merge PDF',
            'Dividir PDF': 'Split PDF',
            'Comprimir PDF': 'Compress PDF',
            'Converter PDF': 'Convert PDF',
            'Todas as ferramentas': 'All tools',
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
            'A senha é usada somente durante o processamento local.': 'The password is used only during local processing.',
            'PDF para Imagens': 'PDF to Images',
            'PDF para JPG': 'PDF to JPG',
            'PDF para PNG': 'PDF to PNG',
            'Desbloquear PDF': 'Unlock PDF',
            'PDF para Excel': 'PDF to Excel',
            'Imagens para PDF': 'Images to PDF',
            'Mesclar PDFs': 'Merge PDFs',
            'Dividir PDF': 'Split PDF',
            'Comprimir PDF': 'Compress PDF',
            'Proteger PDF': 'Protect PDF',
            "Marca d'água": 'Watermark',
            'Números de página': 'Page numbers',
            'PDF para PDF/A': 'PDF to PDF/A',
            'Word para PDF': 'Word to PDF',
            'Excel para PDF': 'Excel to PDF',
            'TXT para PDF': 'TXT to PDF',
            'PDF para Word': 'PDF to Word',
            'PDF para Texto': 'PDF to Text',
            'OCR em PDF': 'OCR PDF',
            'Editar PDF': 'Edit PDF',
            'Converta páginas PDF em imagens JPG compactadas': 'Convert PDF pages into compressed JPG images',
            'Converta páginas PDF em imagens PNG em alta definição': 'Convert PDF pages into high definition PNG images',
            'Remova permanentemente a senha do seu PDF': 'Permanently remove password from your PDF',
            'Extraia tabelas do PDF para planilhas Excel (.xlsx) editáveis': 'Extract tables from PDF into editable Excel (.xlsx) spreadsheets',
            'Extraia todas ou páginas específicas por intervalo': 'Extract all or specific pages by range',
            'Converta páginas PDF em imagens JPG ou PNG': 'Convert PDF pages into JPG or PNG images',
            'Combine várias imagens em um único PDF': 'Combine multiple images into one PDF',
            'Combine vários arquivos PDF em um documento único': 'Combine multiple PDF files into one document',
            'Extraia páginas específicas do seu PDF': 'Extract specific pages from your PDF',
            'Reduza o tamanho do seu arquivo PDF': 'Reduce your PDF file size',
            'Adicione uma senha local ao seu documento PDF': 'Add a local password to your PDF',
            "Adicione uma marca d'água de texto ao PDF": 'Add a text watermark to your PDF',
            'Numere as páginas do documento localmente': 'Number your document pages locally',
            'Padronize seu PDF para arquivamento (PDF/A)': 'Convert your PDF to the archival PDF/A standard',
            'Converta um ou mais documentos DOCX para PDF': 'Convert one or more DOCX documents to PDF',
            'Converta planilhas XLSX para PDF': 'Convert XLSX spreadsheets to PDF',
            'Converta arquivos de texto simples para PDF': 'Convert plain text files to PDF',
            'Converta documentos PDF para Word (.docx) editável': 'Convert PDF documents to editable Word (.docx) files',
            'Extraia o texto do PDF para um arquivo TXT editável': 'Extract PDF text into an editable TXT file',
            'Extraia texto de PDFs e imagens escaneadas com OCR': 'Extract text from scanned PDFs and images with OCR',
            'Reordene, insira, gire, duplique e exclua páginas diretamente no PDF': 'Reorder, insert, rotate, duplicate and delete pages directly in the PDF',
            'Senha do PDF para desbloquear': 'PDF Password to unlock',
            'Digite a senha atual do documento': 'Enter current document password',
            'A senha é usada apenas localmente na sua máquina para remover a proteção.': 'The password is used only locally on your computer to remove protection.',
            'Intervalo de páginas (opcional)': 'Page range (optional)',
            'Ex.: 1-3, 5, 8-10 (deixe em branco para extrair todas)': 'e.g.: 1-3, 5, 8-10 (leave blank to extract all)',
            'Todas as páginas serão extraídas em arquivos individuais se este campo ficar em branco.': 'All pages will be extracted to individual files if left blank.',
            'Converter para JPG': 'Convert to JPG',
            'Converter para PNG': 'Convert to PNG',
            'Extrair para Excel': 'Extract to Excel',
            'Senha obrigatória': 'Password required',
            'Informe a senha do documento para desbloqueá-lo.': 'Enter document password to unlock it.',
            'Sobre': 'About',
            'Sobre o LocalPDF.io': 'About LocalPDF.io',
            '100% Privado & Local': '100% Private & Local',
            'Posição & Estilo': 'Position & Style',
            'Diagonal 45° centralizada (Recomendado)': 'Diagonal 45° centered (Recommended)',
            'Centro horizontal': 'Horizontal center',
            'Cabeçalho (parte superior)': 'Header (top)',
            'Rodapé (parte inferior)': 'Footer (bottom)',
            'Cor da marca': 'Watermark color',
            'Cinza discreto (elegante)': 'Subtle gray (elegant)',
            'Vermelho suave (confidencial)': 'Soft red (confidential)',
            'Azul corporativo': 'Corporate blue',
            'Opacidade / Transparência': 'Opacity / Transparency',
            'Suave (22% - texto 100% legível)': 'Light (22% - fully readable text)',
            'Médio (35% - equilibrado)': 'Medium (35% - balanced)',
            'Destacado (55% - visível)': 'Strong (55% - prominent)',
            'Zero Nuvem': 'Zero Cloud',
            'Sem Upload Externo': 'No External Upload',
            'Funciona Offline': 'Works Offline',
            'Sem Limite de Páginas': 'No Page Limit',
            'Instalador Windows (.msi)': 'Windows Installer (.msi)',
            'Código Aberto MIT': 'Open Source MIT',
            'Privacidade em Primeiro Lugar': 'Privacy First',
            'Desempenho Ilimitado': 'Unlimited Performance',
            'Totalmente Offline': 'Completely Offline',
            'Conformidade LGPD & GDPR': 'LGPD & GDPR Compliance',
            'Por que escolher o LocalPDF.io?': 'Why choose LocalPDF.io?',
            'Recursos e Ferramentas Integradas': 'Integrated Features & Tools',
            'Arquitetura & Tecnologias': 'Architecture & Technologies',
            'Versão Web no Railway': 'Web Version on Railway',
            'Seus arquivos são processados na memória temporária do servidor e excluídos logo após o download. Para processamento 100% offline e ilimitado no seu PC:': 'Your files are processed in server temporary memory and deleted right after download. For 100% offline and unlimited processing on your PC:',
            'Baixar App Desktop (100% Local)': 'Download Desktop App (100% Local)',
            'Baixar para Windows': 'Download for Windows',
            'Versão Web (Railway) vs. Aplicativo Desktop Local': 'Web Version (Railway) vs. Local Desktop Application',
            'Aplicativo Desktop (Windows .msi)': 'Desktop Application (Windows .msi)',
            'Versão Web (Nuvem Railway)': 'Web Version (Railway Cloud)',
            'Recomendado para Máxima Privacidade': 'Recommended for Maximum Privacy',
            'Praticidade Imediata no Navegador': 'Instant Browser Convenience',
            'Baixar Instalador Windows (.msi)': 'Download Windows Installer (.msi)',
            'LocalPDF Desktop (Local)': 'LocalPDF Desktop (Local)',
            'LocalPDF Web (Railway)': 'LocalPDF Web (Railway)',
            'Serviços na Nuvem (iLovePDF, Smallpdf, etc.)': 'Cloud Services (iLovePDF, Smallpdf, etc.)',
            'Seus arquivos saem do seu computador?': 'Do your files leave your computer?',
            'Persistência ou retenção em disco': 'Disk persistence or retention',
            'Limite de tamanho de arquivo': 'File size limit',
            'Funciona sem conexão à Internet?': 'Works without internet connection?',
            'Exige cadastro ou assinatura?': 'Requires registration or subscription?',
            'Segurança para dados sensíveis e contratos': 'Security for sensitive data & contracts'
        };

        const toolTranslations = {
            'pdf-to-images': ['PDF to Images', 'Convert each PDF page into separate JPG or PNG images'],
            'pdf-to-jpg': ['PDF to JPG', 'Convert PDF pages into compressed JPG images'],
            'pdf-to-png': ['PDF to PNG', 'Convert PDF pages into high definition PNG images'],
            'unlock-pdf': ['Unlock PDF', 'Permanently remove password from your PDF document'],
            'pdf-to-excel': ['PDF to Excel', 'Extract tables from PDF into editable Excel (.xlsx) spreadsheets'],
            'images-to-pdf': ['Images to PDF', 'Combine multiple images into one PDF'],
            'merge-pdf': ['Merge PDFs', 'Combine multiple PDF files into one document'],
            'split-pdf': ['Split PDF', 'Extract all or specific pages by range'],
            'compress-pdf': ['Compress PDF', 'Reduce PDF file size while preserving quality'],
            'protect-pdf': ['Protect PDF', 'Create a password-protected copy of your PDF'],
            'watermark-pdf': ["Watermark", 'Add a text watermark to every page'],
            'page-numbers-pdf': ['Page numbers', 'Add numbering to your PDF document'],
            'pdf-to-pdfa': ['PDF to PDF/A', 'Convert PDFs to the PDF/A-1b archival standard'],
            'word-to-pdf': ['Word to PDF', 'Convert one or more DOCX files to PDF'],
            'excel-to-pdf': ['Excel to PDF', 'Convert Excel spreadsheets to PDF'],
            'txt-to-pdf': ['TXT to PDF', 'Convert plain text files to formatted PDF'],
            'pdf-to-word': ['PDF to Word', 'Convert PDF documents into editable DOCX files'],
            'pdf-to-text': ['PDF to Text', 'Extract selectable text from every PDF page'],
            'ocr-pdf': ['OCR PDF', 'Extract text from scanned PDFs and images using OCR']
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
            if (typeof updateWebModeElements === 'function') updateWebModeElements();
            if (currentTool) showTool(currentTool);
        }

        const tools = {
            'pdf-to-images': {
                title: 'PDF para Imagens',
                description: 'Converta cada página do seu PDF em imagens separadas',
                accept: '.pdf',
                multiple: false
            },
            'pdf-to-jpg': {
                title: 'PDF para JPG',
                description: 'Converta páginas do seu PDF em imagens JPG compactadas',
                accept: '.pdf',
                multiple: false
            },
            'pdf-to-png': {
                title: 'PDF para PNG',
                description: 'Converta páginas do seu PDF em imagens PNG em alta resolução',
                accept: '.pdf',
                multiple: false
            },
            'images-to-pdf': {
                title: 'Imagens para PDF',
                description: 'Combine múltiplas imagens em um único arquivo PDF',
                accept: '.jpg,.jpeg,.png',
                multiple: true
            },
            'merge-pdf': {
                title: 'Mesclar PDFs',
                description: 'Combine vários arquivos PDF em um documento único',
                accept: '.pdf',
                multiple: true
            },
            'split-pdf': {
                title: 'Dividir PDF',
                description: 'Extraia páginas específicas ou todas as páginas do seu PDF',
                accept: '.pdf',
                multiple: false,
                options: 'split-options'
            },
            'compress-pdf': {
                title: 'Comprimir PDF',
                description: 'Reduza o tamanho do arquivo PDF mantendo a qualidade',
                accept: '.pdf',
                multiple: false
            },
            'protect-pdf': {
                title: 'Proteger PDF',
                description: 'Crie uma cópia protegida do seu PDF com senha',
                accept: '.pdf',
                multiple: false,
                options: 'password'
            },
            'unlock-pdf': {
                title: 'Desbloquear PDF',
                description: 'Remova a senha e proteção do seu documento PDF',
                accept: '.pdf',
                multiple: false,
                options: 'unlock-password'
            },
            'watermark-pdf': {
                title: "Marca d'água",
                description: "Adicione uma marca d'água de texto em todas as páginas",
                accept: '.pdf',
                multiple: false,
                options: 'watermark'
            },
            'page-numbers-pdf': {
                title: 'Números de página',
                description: 'Adicione numeração ao seu documento PDF',
                accept: '.pdf',
                multiple: false,
                options: 'page-numbers'
            },
            'pdf-to-pdfa': {
                title: 'PDF para PDF/A',
                description: 'Converta PDFs para o padrão de arquivamento PDF/A-1b',
                accept: '.pdf',
                multiple: true
            },
            'word-to-pdf': {
                title: 'Word para PDF',
                description: 'Converta documentos Word (.docx) para PDF - aceita múltiplos arquivos',
                accept: '.docx',
                multiple: true
            },
            'excel-to-pdf': {
                title: 'Excel para PDF',
                description: 'Converta planilhas Excel (.xlsx) para PDF',
                accept: '.xlsx',
                multiple: false
            },
            'pdf-to-excel': {
                title: 'PDF para Excel',
                description: 'Extraia tabelas do PDF diretamente para planilhas Excel (.xlsx)',
                accept: '.pdf',
                multiple: false
            },
            'txt-to-pdf': {
                title: 'TXT para PDF',
                description: 'Converta arquivos de texto simples (.txt) para PDF',
                accept: '.txt',
                multiple: false
            },
            'pdf-to-word': {
                title: 'PDF para Word',
                description: 'Converta seus documentos PDF para Word (.docx) editável',
                accept: '.pdf',
                multiple: false
            },
            'pdf-to-text': {
                title: 'PDF para Texto',
                description: 'Extraia o texto selecionável de todas as páginas do PDF',
                accept: '.pdf',
                multiple: false
            },
            'ocr-pdf': {
                title: 'OCR em PDF',
                description: 'Extraia texto de PDFs e imagens escaneadas usando reconhecimento óptico de caracteres',
                accept: '.pdf,.jpg,.jpeg,.png',
                multiple: false
            }
        };

        function formatFileSize(bytes) {
            if (!bytes || bytes === 0) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
        }

        function getFileExtension(filename) {
            return filename.slice((filename.lastIndexOf(".") - 1 >>> 0) + 2).toUpperCase() || 'FILE';
        }

        function filterCategory(category, tabBtn) {
            document.querySelectorAll('.category-tab').forEach(b => {
                b.classList.remove('active');
                b.setAttribute('aria-selected', 'false');
            });
            if (tabBtn) {
                tabBtn.classList.add('active');
                tabBtn.setAttribute('aria-selected', 'true');
            }
            const cards = document.querySelectorAll('.tools-grid .tool-card');
            cards.forEach(card => {
                if (category === 'all' || card.getAttribute('data-category') === category) {
                    card.classList.remove('card-hidden-filter');
                    card.classList.add('card-visible');
                } else {
                    card.classList.add('card-hidden-filter');
                }
            });
        }

        let hasNavigatedInApp = false;

        function navigateBack() {
            if (hasNavigatedInApp && window.history.length > 1) {
                window.history.back();
            } else {
                showHome(true);
            }
        }

        function updateDocumentSEO(view, toolName, rawTitle, rawDesc) {
            let title = 'LocalPDF.io — Ferramentas de PDF 100% Privadas, Grátis e Ilimitadas';
            let desc = 'Converta, junte, divida, comprima e edite arquivos PDF gratuitamente no seu navegador ou PC. Sem limite de tamanho, sem cadastro e com privacidade absoluta.';
            let canonicalPath = '/';

            if (view === 'tool' && toolName) {
                const tName = rawTitle || (tools[toolName] ? tools[toolName].title : toolName);
                title = `${tName} Online e Grátis — LocalPDF.io`;
                desc = rawDesc || (tools[toolName] ? tools[toolName].description : '');
                canonicalPath = `/tool/${toolName}`;
            } else if (view === 'editor') {
                title = currentLanguage === 'en' ? 'PDF Editor Online & Free — LocalPDF.io' : 'Editor de PDF Online Grátis — Organizar, Girar e Reordenar Páginas | LocalPDF.io';
                desc = currentLanguage === 'en' ? 'Reorder, rotate, duplicate and delete PDF pages visually.' : 'Edite a estrutura do seu PDF visualmente: reordene páginas, gire, duplique e exclua sem perder formatação. 100% privado e gratuito.';
                canonicalPath = '/editor';
            } else if (view === 'about') {
                title = currentLanguage === 'en' ? 'About LocalPDF.io — 100% Private PDF Suite' : 'Sobre o LocalPDF.io — Suíte de PDF com Privacidade Absoluta';
                desc = currentLanguage === 'en' ? 'Learn about LocalPDF.io: private, fast PDF processing.' : 'Conheça o LocalPDF.io: ferramentas de PDF criadas com foco total em privacidade. Modo nuvem com processamento efêmero e app desktop 100% offline.';
                canonicalPath = '/about';
            }

            document.title = title;
            const metaDesc = document.querySelector('meta[name="description"]');
            if (metaDesc) metaDesc.setAttribute('content', desc);
            const ogTitle = document.querySelector('meta[property="og:title"]');
            if (ogTitle) ogTitle.setAttribute('content', title);
            const ogDesc = document.querySelector('meta[property="og:description"]');
            if (ogDesc) ogDesc.setAttribute('content', desc);
            const canonical = document.querySelector('link[rel="canonical"]');
            if (canonical) canonical.setAttribute('href', `https://localpdf.io${canonicalPath}`);
        }

        function dismissWebBanner() {
            const banner = document.getElementById('env-banner');
            if (banner) {
                banner.classList.add('dismissed');
                try {
                    sessionStorage.setItem('localpdf_banner_dismissed', '1');
                } catch (e) {}
            }
        }

        function showTool(toolName, push = true) {
            currentTool = toolName;
            const tool = tools[toolName];
            if (!tool) return;

            document.getElementById('home-view').classList.add('hidden');
            const editorView = document.getElementById('editor-view');
            if (editorView) editorView.classList.add('hidden');
            const aboutView = document.getElementById('about-view');
            if (aboutView) aboutView.classList.add('hidden');
            document.getElementById('tool-views').classList.remove('hidden');
            const translatedTool = toolTranslations[toolName];
            const rawTitle = currentLanguage === 'en' ? (translatedTool ? translatedTool[0] : tool.title) : tool.title;
            const rawDesc = currentLanguage === 'en' ? (translatedTool ? translatedTool[1] : tool.description) : tool.description;
            document.getElementById('tool-title').innerText = rawTitle;
            document.getElementById('tool-description').innerText = rawDesc;
            document.getElementById('file-input').accept = tool.accept;
            document.getElementById('file-input').multiple = tool.multiple;
            renderToolOptions(tool.options);

            uploadedFiles = [];
            updateFileList();
            hideResult();
            updateToolPrivacyNotice();

            updateDocumentSEO('tool', toolName, rawTitle, rawDesc);

            if (push) {
                hasNavigatedInApp = true;
                const targetUrl = `/tool/${toolName}`;
                if (window.location.pathname !== targetUrl) {
                    history.pushState({ localpdf: true, view: 'tool', tool: toolName }, '', targetUrl);
                }
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function updateToolPrivacyNotice() {
            const privacyBadge = document.getElementById('tool-privacy-badge');
            if (!privacyBadge) return;
            if (isWebMode) {
                privacyBadge.className = 'tool-privacy-notice web-mode';
                privacyBadge.innerHTML = `<span>☁️</span><div><strong>${t('Modo Web Seguro (Railway):', 'Secure Web Mode (Railway):')}</strong> ${t('Seus arquivos são transmitidos com criptografia HTTPS, processados temporariamente na memória volátil do servidor e excluídos de forma definitiva imediatamente após o download. Nenhuma cópia é armazenada.', 'Your files are transmitted via HTTPS encryption, processed temporarily in server volatile memory and permanently deleted immediately after download. No copies are stored.')} <a href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" rel="noopener">${t('Baixar App Desktop para 100% offline no PC', 'Download Desktop App for 100% offline on PC')}</a></div>`;
            } else {
                privacyBadge.className = 'tool-privacy-notice local-mode';
                privacyBadge.innerHTML = `<span>🛡️</span><div><strong>${t('Processamento 100% Local:', '100% Local Processing:')}</strong> ${t('Seus arquivos são processados diretamente na memória do seu computador e nunca saem da sua máquina.', 'Your files are processed directly in your computer memory and never leave your machine.')}</div>`;
            }
        }

        function renderToolOptions(optionType) {
            const options = document.getElementById('options');
            if (optionType === 'password') {
                const pwdNote = isWebMode
                    ? t('A senha é usada apenas em memória temporária durante a geração do PDF e nunca é armazenada.', 'The password is used only in temporary memory during PDF generation and is never stored.')
                    : t('A senha é usada somente durante o processamento local.', 'The password is used only during local processing.');
                options.innerHTML = `
                    <label for="pdf-password">${t('Senha do PDF', 'PDF Password')}</label>
                    <input id="pdf-password" type="password" minlength="4" autocomplete="new-password" placeholder="${t('Digite uma senha com pelo menos 4 caracteres', 'Enter a password with at least 4 characters')}">
                    <small>${pwdNote}</small>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'watermark') {
                options.innerHTML = `
                    <label for="watermark-text">${t("Texto da marca d'água", 'Watermark text')}</label>
                    <input id="watermark-text" type="text" maxlength="80" placeholder="${t('Ex.: CONFIDENCIAL', 'e.g.: CONFIDENTIAL')}">
                    <label for="watermark-position">${t('Posição & Estilo', 'Position & Style')}</label>
                    <select id="watermark-position">
                        <option value="diagonal">${t('Diagonal 45° centralizada (Recomendado)', 'Diagonal 45° centered (Recommended)')}</option>
                        <option value="center">${t('Centro horizontal', 'Horizontal center')}</option>
                        <option value="top">${t('Cabeçalho (parte superior)', 'Header (top)')}</option>
                        <option value="bottom">${t('Rodapé (parte inferior)', 'Footer (bottom)')}</option>
                    </select>
                    <label for="watermark-color">${t('Cor da marca', 'Watermark color')}</label>
                    <select id="watermark-color">
                        <option value="gray">${t('Cinza discreto (elegante)', 'Subtle gray (elegant)')}</option>
                        <option value="red">${t('Vermelho suave (confidencial)', 'Soft red (confidential)')}</option>
                        <option value="blue">${t('Azul corporativo', 'Corporate blue')}</option>
                    </select>
                    <label for="watermark-opacity">${t('Opacidade / Transparência', 'Opacity / Transparency')}</label>
                    <select id="watermark-opacity">
                        <option value="0.22">${t('Suave (22% - texto 100% legível)', 'Light (22% - fully readable text)')}</option>
                        <option value="0.35">${t('Médio (35% - equilibrado)', 'Medium (35% - balanced)')}</option>
                        <option value="0.55">${t('Destacado (55% - visível)', 'Strong (55% - prominent)')}</option>
                    </select>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'page-numbers') {
                options.innerHTML = `
                    <label for="page-number-position">${t('Posição da numeração', 'Page numbers position')}</label>
                    <select id="page-number-position">
                        <option value="bottom-center">${t('Rodapé central', 'Bottom center')}</option>
                        <option value="bottom-right">${t('Rodapé direito', 'Bottom right')}</option>
                        <option value="top-center">${t('Cabeçalho central', 'Top center')}</option>
                        <option value="top-right">${t('Cabeçalho direito', 'Top right')}</option>
                    </select>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'unlock-password') {
                const unlockNote = isWebMode
                    ? t('A senha é usada apenas na memória temporária do servidor para remover a proteção e é descartada imediatamente.', 'The password is used only in temporary server memory to remove protection and is discarded immediately.')
                    : t('A senha é usada apenas localmente na sua máquina para remover a proteção.', 'The password is used only locally on your computer to remove protection.');
                options.innerHTML = `
                    <label for="pdf-unlock-password">${t('Senha do PDF para desbloquear', 'PDF Password to unlock')}</label>
                    <input id="pdf-unlock-password" type="password" autocomplete="current-password" placeholder="${t('Digite a senha atual do documento', 'Enter current document password')}">
                    <small>${unlockNote}</small>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'split-options') {
                options.innerHTML = `
                    <label for="split-page-range">${t('Intervalo de páginas (opcional)', 'Page range (optional)')}</label>
                    <input id="split-page-range" type="text" placeholder="${t('Ex.: 1-3, 5, 8-10 (deixe em branco para extrair todas)', 'e.g.: 1-3, 5, 8-10 (leave blank to extract all)')}">
                    <small>${t('Todas as páginas serão extraídas em arquivos individuais se este campo ficar em branco.', 'All pages will be extracted to individual files if left blank.')}</small>
                `;
                options.classList.remove('hidden');
                return;
            }
            options.innerHTML = '';
            options.classList.add('hidden');
        }

        function showHome(push = true) {
            currentTool = null;
            document.getElementById('home-view').classList.remove('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            const editorView = document.getElementById('editor-view');
            if (editorView) editorView.classList.add('hidden');
            const aboutView = document.getElementById('about-view');
            if (aboutView) aboutView.classList.add('hidden');
            uploadedFiles = [];
            hideResult();
            if (typeof resetEditor === 'function') resetEditor();

            updateDocumentSEO('home');

            if (push) {
                const targetUrl = '/';
                if (window.location.pathname !== targetUrl) {
                    history.pushState({ localpdf: true, view: 'home' }, '', targetUrl);
                }
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function showAbout(push = true) {
            currentTool = null;
            document.getElementById('home-view').classList.add('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            const editorView = document.getElementById('editor-view');
            if (editorView) editorView.classList.add('hidden');
            const aboutView = document.getElementById('about-view');
            if (aboutView) aboutView.classList.remove('hidden');
            uploadedFiles = [];
            hideResult();

            updateDocumentSEO('about');

            if (push) {
                hasNavigatedInApp = true;
                const targetUrl = '/about';
                if (window.location.pathname !== targetUrl && window.location.pathname !== '/sobre') {
                    history.pushState({ localpdf: true, view: 'about' }, '', targetUrl);
                }
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        async function updateFileList() {
            const fileList = document.getElementById('file-list');
            const convertBtn = document.getElementById('convert-btn');

            if (uploadedFiles.length === 0) {
                fileList.innerHTML = '';
                convertBtn.classList.add('hidden');
                return;
            }

            // Render items immediately with placeholder thumbs, then fill in async
            fileList.innerHTML = uploadedFiles.map((file, index) => {
                const ext = getFileExtension(file.name);
                const sizeStr = formatFileSize(file.size);
                const key = `${file.name}-${file.size}`;
                const cached = fileThumbnails[key];
                const thumbHtml = cached
                    ? `<img class="file-thumb" src="${cached}" alt="preview">`
                    : `<div class="file-thumb-placeholder" data-thumb-idx="${index}" title="${file.name}"></div>`;
                return `
                    <div class="file-item" id="file-item-${index}">
                        <div class="file-info-group">
                            ${thumbHtml}
                            <span class="file-ext-badge">${ext}</span>
                            <span class="file-name-text" title="${file.name}">${file.name}</span>
                            <span class="file-size-text">(${sizeStr})</span>
                        </div>
                        <button type="button" class="file-remove-btn" onclick="removeFile(${index})" title="${t('Remover arquivo', 'Remove file')}">
                            ✕ <span>${t('Remover', 'Remove')}</span>
                        </button>
                    </div>
                `;
            }).join('');

            // Load thumbnails asynchronously for any placeholders
            uploadedFiles.forEach(async (file, index) => {
                const key = `${file.name}-${file.size}`;
                if (fileThumbnails[key]) return;
                const thumb = await generateThumbnail(file);
                if (!thumb) return;
                const placeholder = fileList.querySelector(`[data-thumb-idx="${index}"]`);
                if (placeholder) {
                    const img = document.createElement('img');
                    img.className = 'file-thumb';
                    img.src = thumb;
                    img.alt = 'preview';
                    placeholder.replaceWith(img);
                }
            });

            let btnText = t('Processar Documento', 'Process Document');
            if (currentTool === 'merge-pdf') btnText = t('Mesclar PDFs', 'Merge PDFs');
            else if (currentTool === 'compress-pdf') btnText = t('Comprimir PDF', 'Compress PDF');
            else if (currentTool === 'split-pdf') btnText = t('Dividir PDF', 'Split PDF');
            else if (currentTool === 'protect-pdf') btnText = t('Proteger PDF', 'Protect PDF');
            else if (currentTool === 'unlock-pdf') btnText = t('Desbloquear PDF', 'Unlock PDF');
            else if (currentTool === 'pdf-to-excel') btnText = t('Extrair para Excel', 'Extract to Excel');
            else if (currentTool === 'pdf-to-jpg') btnText = t('Converter para JPG', 'Convert to JPG');
            else if (currentTool === 'pdf-to-png') btnText = t('Converter para PNG', 'Convert to PNG');
            else if (currentTool === 'watermark-pdf') btnText = t("Aplicar Marca d'Água", 'Apply Watermark');
            else if (currentTool === 'ocr-pdf') btnText = t('Executar OCR', 'Run OCR');
            else if (currentTool) btnText = t('Converter', 'Convert');

            convertBtn.textContent = btnText;
            convertBtn.disabled = false;
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

        function resetToolFlow() {
            uploadedFiles = [];
            updateFileList();
            hideResult();
            const fileInput = document.getElementById('file-input');
            if (fileInput) fileInput.value = '';
        }

        let lastSavedPath = '';

        async function openSavedFile(path) {
            const targetPath = path || lastSavedPath;
            if (!targetPath) return;
            if (window.pywebview && window.pywebview.api && window.pywebview.api.open_saved_file) {
                try {
                    const ok = await window.pywebview.api.open_saved_file(targetPath);
                    if (ok) return;
                } catch (e) {}
            }
            try {
                await fetch('/api/desktop/open-file', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path: targetPath })
                });
            } catch (e) {
                showToast(t('Não foi possível abrir o arquivo.', 'Could not open file.'), 'error');
            }
        }

        async function openSavedFolder(path) {
            const targetPath = path || lastSavedPath;
            if (!targetPath) return;
            if (window.pywebview && window.pywebview.api && window.pywebview.api.open_saved_folder) {
                try {
                    const ok = await window.pywebview.api.open_saved_folder(targetPath);
                    if (ok) return;
                } catch (e) {}
            }
            try {
                await fetch('/api/desktop/open-folder', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path: targetPath })
                });
            } catch (e) {
                showToast(t('Não foi possível abrir a pasta.', 'Could not open folder.'), 'error');
            }
        }

        async function saveDocumentResult(blob, filename, forceDownloads = false) {
            const isDesktop = !isWebMode || (window.pywebview && window.pywebview.api);

            if (isDesktop) {
                // If not forced to Downloads, try native SaveFileDialog in PyWebView
                if (!forceDownloads && window.pywebview && window.pywebview.api && window.pywebview.api.choose_save_path) {
                    try {
                        const chosenPath = await window.pywebview.api.choose_save_path(filename);
                        if (!chosenPath) {
                            // User clicked cancel on the dialog
                            return { cancelled: true, filename };
                        }
                        // Write to the chosen destination path via backend
                        const formData = new FormData();
                        formData.append('file', blob, filename);
                        formData.append('path', chosenPath);
                        const resp = await fetch('/api/desktop/write-file', {
                            method: 'POST',
                            body: formData
                        });
                        if (resp.ok) {
                            const resData = await resp.json();
                            lastSavedPath = resData.path || chosenPath;
                            return { success: true, path: lastSavedPath, filename };
                        }
                    } catch (e) {
                        console.warn('Native dialog failed, falling back to Downloads folder:', e);
                    }
                }

                // Automatic save to user's Downloads folder
                try {
                    const formData = new FormData();
                    formData.append('file', blob, filename);
                    const resp = await fetch('/api/desktop/save', {
                        method: 'POST',
                        body: formData
                    });
                    if (resp.ok) {
                        const resData = await resp.json();
                        lastSavedPath = resData.path;
                        return { success: true, path: resData.path, filename: resData.filename || filename };
                    }
                } catch (e) {
                    console.warn('Desktop save to Downloads failed:', e);
                }
            }

            // Web mode fallback (or if desktop endpoints unavailable)
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                window.URL.revokeObjectURL(url);
                if (a.parentNode) a.parentNode.removeChild(a);
            }, 60000);

            return { success: true, webDownload: true, filename };
        }

        function renderResultCard(saveRes) {
            const resultEl = document.getElementById('result');
            if (!saveRes || !resultEl) return;

            if (saveRes.cancelled) {
                resultEl.innerHTML = `
                    <div class="result-card warning">
                        <div class="result-header">
                            <h4>⚠️ ${t('Salvamento cancelado', 'Save cancelled')}</h4>
                        </div>
                        <div class="result-body">
                            <span class="result-filename">📄 ${saveRes.filename || lastDownloadedFilename}</span>
                            <p>${t('O processamento foi concluído com sucesso. Como você cancelou a seleção da pasta, escolha como deseja salvar o arquivo:', 'The processing was completed successfully. Since you cancelled folder selection, choose how you would like to save the file:')}</p>
                            <div class="result-actions">
                                <button type="button" class="btn-download-again" onclick="downloadAgain(true)">📥 ${t('Salvar em Downloads', 'Save to Downloads')}</button>
                                <button type="button" class="btn-open-folder" onclick="downloadAgain(false)">💾 ${t('Escolher pasta...', 'Choose folder...')}</button>
                                <button type="button" class="btn-reset-flow" onclick="resetToolFlow()">✨ ${t('Processar outro arquivo', 'Process another file')}</button>
                            </div>
                        </div>
                    </div>
                `;
                resultEl.classList.remove('hidden');
                return;
            }

            if (saveRes.path) {
                lastSavedPath = saveRes.path;
                showToast(t('Arquivo salvo com sucesso no seu computador!', 'File saved successfully to your computer!'), 'success');
                resultEl.innerHTML = `
                    <div class="result-card success">
                        <div class="result-header">
                            <h4>✅ ${t('Concluído e salvo no seu computador!', 'Completed and saved on your computer!')}</h4>
                        </div>
                        <div class="result-body">
                            <span class="result-filename">📄 ${saveRes.filename || lastDownloadedFilename}</span>
                            <div class="saved-path-badge">
                                <span>💾</span>
                                <span><strong>${t('Salvo em:', 'Saved to:')}</strong> ${saveRes.path}</span>
                            </div>
                            <div class="result-actions">
                                <button type="button" class="btn-open-file" onclick="openSavedFile()">📄 ${t('Abrir arquivo', 'Open file')}</button>
                                <button type="button" class="btn-open-folder" onclick="openSavedFolder()">📂 ${t('Abrir pasta', 'Open folder')}</button>
                                <button type="button" class="btn-download-again" onclick="downloadAgain(false)">💾 ${t('Salvar em outro local...', 'Save elsewhere...')}</button>
                                <button type="button" class="btn-reset-flow" onclick="resetToolFlow()">✨ ${t('Processar outro arquivo', 'Process another file')}</button>
                            </div>
                        </div>
                    </div>
                `;
                resultEl.classList.remove('hidden');
                return;
            }

            // Web download mode
            showToast(t('Arquivo processado e download iniciado!', 'File processed and download started!'), 'success');
            const successNote = isWebMode
                ? t('O arquivo foi processado com segurança em memória volátil e o download foi iniciado. O arquivo já foi excluído do servidor.', 'The file was securely processed in volatile memory and the download has started. The file has already been deleted from the server.')
                : t('O arquivo foi processado no seu computador e o download foi iniciado automaticamente.', 'The file was processed on your computer and the download started automatically.');

            resultEl.innerHTML = `
                <div class="result-card success">
                    <div class="result-header">
                        <h4>✅ ${t('Concluído com sucesso!', 'Completed successfully!')}</h4>
                    </div>
                    <div class="result-body">
                        <span class="result-filename">📄 ${lastDownloadedFilename}</span>
                        <p>${successNote}</p>
                        <div class="result-actions">
                            <button type="button" class="btn-download-again" onclick="downloadAgain(false)">📥 ${t('Baixar novamente', 'Download again')}</button>
                            <button type="button" class="btn-reset-flow" onclick="resetToolFlow()">✨ ${t('Processar outro arquivo', 'Process another file')}</button>
                        </div>
                    </div>
                </div>
            `;
            resultEl.classList.remove('hidden');
        }

        async function downloadAgain(forceDownloads = false) {
            if (!lastDownloadedBlob) return;
            const res = await saveDocumentResult(lastDownloadedBlob, lastDownloadedFilename, forceDownloads);
            renderResultCard(res);
        }

        // ── Named progress stages per tool ───────────────────────────
        const PROGRESS_STAGES = {
            default: [
                { pct: 15, pt: 'Enviando arquivo...', en: 'Uploading file...' },
                { pct: 35, pt: 'Analisando documento...', en: 'Analyzing document...' },
                { pct: 60, pt: 'Processando conteúdo...', en: 'Processing content...' },
                { pct: 80, pt: 'Otimizando resultado...', en: 'Optimizing result...' },
                { pct: 93, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'ocr-pdf': [
                { pct: 10, pt: 'Enviando arquivo...', en: 'Uploading file...' },
                { pct: 25, pt: 'Renderizando páginas para OCR...', en: 'Rendering pages for OCR...' },
                { pct: 55, pt: 'Reconhecendo texto (OCR)... pode demorar.', en: 'Recognizing text (OCR)... may take a moment.' },
                { pct: 80, pt: 'Consolidando resultado...', en: 'Consolidating result...' },
                { pct: 93, pt: 'Preparando arquivo de texto...', en: 'Preparing text file...' }
            ],
            'compress-pdf': [
                { pct: 15, pt: 'Enviando arquivo...', en: 'Uploading file...' },
                { pct: 40, pt: 'Otimizando imagens...', en: 'Optimizing images...' },
                { pct: 70, pt: 'Reestruturando PDF...', en: 'Restructuring PDF...' },
                { pct: 88, pt: 'Comprimindo streams...', en: 'Compressing streams...' },
                { pct: 94, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'merge-pdf': [
                { pct: 15, pt: 'Enviando arquivos...', en: 'Uploading files...' },
                { pct: 40, pt: 'Lendo documentos...', en: 'Reading documents...' },
                { pct: 65, pt: 'Mesclando páginas...', en: 'Merging pages...' },
                { pct: 85, pt: 'Organizando estrutura...', en: 'Organizing structure...' },
                { pct: 94, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'word-to-pdf': [
                { pct: 15, pt: 'Enviando documento...', en: 'Uploading document...' },
                { pct: 35, pt: 'Lendo estrutura do Word...', en: 'Reading Word structure...' },
                { pct: 60, pt: 'Convertendo formatação...', en: 'Converting formatting...' },
                { pct: 82, pt: 'Gerando PDF...', en: 'Generating PDF...' },
                { pct: 94, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'pdf-to-word': [
                { pct: 15, pt: 'Enviando PDF...', en: 'Uploading PDF...' },
                { pct: 35, pt: 'Analisando layout...', en: 'Analyzing layout...' },
                { pct: 62, pt: 'Extraindo texto e tabelas...', en: 'Extracting text and tables...' },
                { pct: 83, pt: 'Gerando arquivo Word...', en: 'Generating Word file...' },
                { pct: 94, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'unlock-pdf': [
                { pct: 20, pt: 'Lendo documento protegido...', en: 'Reading protected document...' },
                { pct: 50, pt: 'Verificando senha e descriptografando...', en: 'Checking password and decrypting...' },
                { pct: 80, pt: 'Gerando PDF desbloqueado...', en: 'Generating unlocked PDF...' },
                { pct: 95, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'pdf-to-excel': [
                { pct: 15, pt: 'Lendo páginas do PDF...', en: 'Reading PDF pages...' },
                { pct: 45, pt: 'Detectando e extraindo tabelas...', en: 'Detecting and extracting tables...' },
                { pct: 75, pt: 'Formatando planilha Excel (.xlsx)...', en: 'Formatting Excel (.xlsx) spreadsheet...' },
                { pct: 95, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'pdf-to-jpg': [
                { pct: 20, pt: 'Renderizando páginas...', en: 'Rendering pages...' },
                { pct: 60, pt: 'Compactando imagens JPG...', en: 'Compressing JPG images...' },
                { pct: 85, pt: 'Criando arquivo compactado...', en: 'Creating archive...' },
                { pct: 95, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'pdf-to-png': [
                { pct: 20, pt: 'Renderizando páginas em alta definição...', en: 'Rendering pages in high definition...' },
                { pct: 60, pt: 'Exportando imagens PNG...', en: 'Exporting PNG images...' },
                { pct: 85, pt: 'Criando arquivo compactado...', en: 'Creating archive...' },
                { pct: 95, pt: 'Preparando download...', en: 'Preparing download...' }
            ],
            'split-pdf': [
                { pct: 20, pt: 'Analisando páginas do PDF...', en: 'Analyzing PDF pages...' },
                { pct: 60, pt: 'Separando e extraindo páginas...', en: 'Splitting and extracting pages...' },
                { pct: 85, pt: 'Gerando arquivos...', en: 'Generating files...' },
                { pct: 95, pt: 'Preparando download...', en: 'Preparing download...' }
            ]
        };

        function startProgress(tool) {
            const progressEl = document.getElementById('progress');
            const progressBar = document.getElementById('progress-bar');
            const progressMsg = document.getElementById('progress-message');
            const progressTimer = document.getElementById('progress-timer');

            progressEl.classList.remove('hidden');
            progressSeconds = 0;
            progressTimer.textContent = '⏱️ 00:00';

            const stages = PROGRESS_STAGES[tool] || PROGRESS_STAGES.default;
            let stageIdx = 0;

            // Set initial stage
            progressBar.style.width = stages[0].pct + '%';
            progressMsg.innerHTML = `<span class="progress-stage-badge">${stageIdx + 1}/${stages.length}</span>${currentLanguage === 'en' ? stages[0].en : stages[0].pt}`;

            // Advance stages over time (distribute evenly, slower for longer tools)
            const stageDuration = tool === 'ocr-pdf' ? 4000 : 2200;
            progressInterval = setInterval(() => {
                progressSeconds++;
                const mins = String(Math.floor(progressSeconds / 60)).padStart(2, '0');
                const secs = String(progressSeconds % 60).padStart(2, '0');
                progressTimer.textContent = `⏱️ ${mins}:${secs}`;

                // Advance to next stage
                const elapsed = progressSeconds * 1000;
                const targetStage = Math.min(
                    Math.floor(elapsed / stageDuration),
                    stages.length - 1
                );
                if (targetStage > stageIdx) {
                    stageIdx = targetStage;
                }
                const stage = stages[stageIdx];
                progressBar.style.width = stage.pct + '%';
                progressMsg.innerHTML = `<span class="progress-stage-badge">${stageIdx + 1}/${stages.length}</span>${currentLanguage === 'en' ? stage.en : stage.pt}`;
            }, 1000);
        }

        function stopProgress(success) {
            if (progressInterval) {
                clearInterval(progressInterval);
                progressInterval = null;
            }
            const progressBar = document.getElementById('progress-bar');
            const progressMsg = document.getElementById('progress-message');
            if (progressBar) {
                progressBar.style.width = success ? '100%' : '0%';
            }
            if (progressMsg && success) {
                progressMsg.innerHTML = `<span class="progress-stage-badge">✓</span>${t('Concluído!', 'Done!')}`;
            }
            setTimeout(() => {
                const progressEl = document.getElementById('progress');
                if (progressEl && success) {
                    progressEl.classList.add('hidden');
                }
            }, 600);
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
                    document.getElementById('result').innerHTML = `
                        <div class="result-card error">
                            <div class="result-header"><h4>⚠️ ${t('Senha inválida', 'Invalid password')}</h4></div>
                            <div class="result-body"><p>${t('Informe uma senha com pelo menos 4 caracteres.', 'Enter a password with at least 4 characters.')}</p></div>
                        </div>
                    `;
                    document.getElementById('result').classList.remove('hidden');
                    return;
                }
                formData.append('password', passwordInput.value);
            }
            const watermarkText = document.getElementById('watermark-text');
            if (watermarkText) {
                if (!watermarkText.value.trim()) {
                    document.getElementById('result').innerHTML = `
                        <div class="result-card error">
                            <div class="result-header"><h4>⚠️ ${t('Texto obrigatório', 'Required text')}</h4></div>
                            <div class="result-body"><p>${t("Informe o texto da marca d'água.", 'Enter the watermark text.')}</p></div>
                        </div>
                    `;
                    document.getElementById('result').classList.remove('hidden');
                    return;
                }
                formData.append('watermark_text', watermarkText.value.trim());
                formData.append('watermark_position', document.getElementById('watermark-position').value);
                const wmColor = document.getElementById('watermark-color');
                if (wmColor) formData.append('watermark_color', wmColor.value);
                const wmOpacity = document.getElementById('watermark-opacity');
                if (wmOpacity) formData.append('watermark_opacity', wmOpacity.value);
            }
            const pageNumberPosition = document.getElementById('page-number-position');
            if (pageNumberPosition) {
                formData.append('page_number_position', pageNumberPosition.value);
            }
            const unlockPasswordInput = document.getElementById('pdf-unlock-password');
            if (unlockPasswordInput) {
                if (!unlockPasswordInput.value) {
                    document.getElementById('result').innerHTML = `
                        <div class="result-card error">
                            <div class="result-header"><h4>⚠️ ${t('Senha obrigatória', 'Password required')}</h4></div>
                            <div class="result-body"><p>${t('Informe a senha do documento para desbloqueá-lo.', 'Enter document password to unlock it.')}</p></div>
                        </div>
                    `;
                    document.getElementById('result').classList.remove('hidden');
                    return;
                }
                formData.append('password', unlockPasswordInput.value);
            }
            const splitPageRange = document.getElementById('split-page-range');
            if (splitPageRange && splitPageRange.value.trim()) {
                formData.append('page_range', splitPageRange.value.trim());
            }

            document.getElementById('convert-btn').disabled = true;
            document.getElementById('convert-btn').textContent = t('Processando...', 'Processing...');
            hideResult();
            startProgress(currentTool);

            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const blob = await response.blob();
                    stopProgress(true);

                    let downloadFilename = '';
                    const disposition = response.headers.get('Content-Disposition');
                    if (disposition && disposition.indexOf('filename=') !== -1) {
                        let filenamePart = disposition.split('filename=')[1].trim();
                        if (filenamePart.startsWith('"') && filenamePart.endsWith('"')) {
                            filenamePart = filenamePart.slice(1, -1);
                        }
                        downloadFilename = decodeURIComponent(filenamePart);
                    }
                    if (!downloadFilename) {
                        downloadFilename = 'localpdf_documento.pdf';
                    }

                    lastDownloadedBlob = blob;
                    lastDownloadedFilename = downloadFilename;

                    const saveRes = await saveDocumentResult(blob, downloadFilename);
                    renderResultCard(saveRes);
                } else {
                    let errorDetail = '';
                    try {
                        const errData = await response.json();
                        errorDetail = errData.error || '';
                    } catch (e) {}
                    throw new Error(errorDetail || t('Erro durante o processamento do documento.', 'Error during document processing.'));
                }
            } catch (error) {
                stopProgress(false);
                showToast(error.message || t('Falha no processamento.', 'Processing failed.'), 'error', 6000);
                document.getElementById('result').innerHTML = `
                    <div class="result-card error">
                        <div class="result-header">
                            <h4>⚠️ ${t('Falha no processamento', 'Processing failed')}</h4>
                        </div>
                        <div class="result-body">
                            <p>${error.message || t('Ocorreu um erro ao processar o arquivo. Verifique se o documento é válido e tente novamente.', 'An error occurred while processing the file. Please verify that the document is valid and try again.')}</p>
                            <div class="result-actions">
                                <button type="button" class="btn-try-again" onclick="hideResult()">${t('Tentar novamente', 'Try again')}</button>
                            </div>
                        </div>
                    </div>
                `;
                document.getElementById('result').classList.remove('hidden');
            } finally {
                const convertBtn = document.getElementById('convert-btn');
                if (convertBtn) {
                    convertBtn.disabled = false;
                    let btnText = t('Processar Documento', 'Process Document');
                    if (currentTool === 'merge-pdf') btnText = t('Mesclar PDFs', 'Merge PDFs');
                    else if (currentTool === 'compress-pdf') btnText = t('Comprimir PDF', 'Compress PDF');
                    else if (currentTool === 'split-pdf') btnText = t('Dividir PDF', 'Split PDF');
                    else if (currentTool === 'protect-pdf') btnText = t('Proteger PDF', 'Protect PDF');
                    else if (currentTool === 'watermark-pdf') btnText = t("Aplicar Marca d'Água", 'Apply Watermark');
                    else if (currentTool === 'ocr-pdf') btnText = t('Executar OCR', 'Run OCR');
                    else if (currentTool) btnText = t('Converter', 'Convert');
                    convertBtn.textContent = btnText;
                }
            }
        }

        let editorFiles = [];
        let editorPages = [];
        let editorHistory = [];
        let editorFuture = [];
        let editorDraggedIndex = null;
        let editorPendingPages = [];
        let editorPendingFiles = [];

        function showEditor(push = true) {
            currentTool = 'edit-pdf';
            document.getElementById('home-view').classList.add('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            const aboutView = document.getElementById('about-view');
            if (aboutView) aboutView.classList.add('hidden');
            document.getElementById('editor-view').classList.remove('hidden');
            resetEditor();

            updateDocumentSEO('editor');

            if (push) {
                hasNavigatedInApp = true;
                const targetUrl = '/editor';
                if (window.location.pathname !== targetUrl && window.location.pathname !== '/tool/edit-pdf') {
                    history.pushState({ localpdf: true, view: 'editor' }, '', targetUrl);
                }
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function showHomeFromEditor() {
            navigateBack();
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

        // ── Lazy loading for editor page thumbnails ──────────────────
        let _editorThumbObserver = null;
        function _setupLazyObserver() {
            if (_editorThumbObserver) _editorThumbObserver.disconnect();
            _editorThumbObserver = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        const img = entry.target;
                        if (img.dataset.lazySrc) {
                            img.src = img.dataset.lazySrc;
                            delete img.dataset.lazySrc;
                            _editorThumbObserver.unobserve(img);
                        }
                    }
                });
            }, { rootMargin: '80px' });
        }

        function renderEditorPages() {
            const container = document.getElementById('editor-pages');
            if (!editorPages.length) {
                container.innerHTML = '<div class="editor-empty">As páginas do PDF aparecerão aqui.</div>';
                return;
            }
            _setupLazyObserver();
            // Eagerly render first 6, lazy-load the rest
            container.innerHTML = editorPages.map((page, index) => {
                const eager = index < 6;
                const imgAttr = eager
                    ? `src="${page.thumbnail}"`
                    : `src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='200' height='280'%3E%3Crect width='100%25' height='100%25' fill='%23f0f2f5'/%3E%3C/svg%3E" data-lazy-src="${page.thumbnail}"`;
                return `
                <div class="editor-page${page.selected ? ' selected' : ''}" draggable="true" data-editor-index="${index}">
                    <img ${imgAttr} alt="Página ${index + 1}">
                    <div class="editor-page-number">Página ${index + 1}</div>
                    <div class="editor-page-actions">
                        <button type="button" data-action="rotate" title="Girar página">↻</button>
                        <button type="button" data-action="duplicate" title="Duplicar página">⧉</button>
                        <button type="button" data-action="delete" title="Excluir página">✕</button>
                    </div>
                </div>`;
            }).join('');
            // Register lazy images with observer
            container.querySelectorAll('img[data-lazy-src]').forEach(img => _editorThumbObserver.observe(img));

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
                const saveRes = await saveDocumentResult(blob, 'localpdf-editado.pdf');
                if (saveRes.cancelled) {
                    status.textContent = t('Salvamento cancelado.', 'Save cancelled.');
                } else if (saveRes.path) {
                    status.textContent = t(`PDF salvo em: ${saveRes.path}`, `PDF saved to: ${saveRes.path}`);
                    showToast(t('PDF editado salvo com sucesso!', 'Edited PDF saved successfully!'), 'success');
                } else {
                    status.textContent = t('PDF editado baixado com sucesso.', 'Edited PDF downloaded successfully.');
                }
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

        const isServerWebMode = {{ 'true' if is_web_mode else 'false' }};
        isWebMode = isServerWebMode || (
            window.location.hostname !== 'localhost' &&
            window.location.hostname !== '127.0.0.1' &&
            !window.location.hostname.endsWith('.local')
        );

        function updateWebModeElements() {
            if (!isWebMode) return;
            const envBanner = document.getElementById('env-banner');
            let isDismissed = false;
            try {
                isDismissed = sessionStorage.getItem('localpdf_banner_dismissed') === '1';
            } catch (e) {}

            if (envBanner) {
                if (isDismissed) {
                    envBanner.classList.add('dismissed');
                } else {
                    envBanner.classList.remove('hidden');
                    envBanner.classList.remove('dismissed');
                }
            }
            const bannerText = document.getElementById('env-banner-text');
            if (bannerText) {
                bannerText.textContent = currentLanguage === 'en'
                    ? 'Files processed in volatile memory and deleted immediately after download.'
                    : 'Arquivos processados em memória volátil e excluídos imediatamente após o download.';
            }
            const bannerCta = document.getElementById('env-banner-cta-text');
            if (bannerCta) {
                bannerCta.textContent = currentLanguage === 'en'
                    ? 'Desktop App (100% Local)'
                    : 'App Desktop (100% Local)';
            }
            const navDownload = document.getElementById('nav-desktop-download');
            if (navDownload) navDownload.classList.remove('hidden');
            const mainSubtitle = document.getElementById('main-subtitle');
            if (mainSubtitle) {
                mainSubtitle.innerHTML = currentLanguage === 'en'
                    ? 'Convert, organize and edit your documents with privacy and speed. Secure cloud processing with immediate file deletion — or <a href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" style="color:#2563eb;font-weight:700;text-decoration:underline;">download the Desktop App</a> for 100% offline use on your PC.'
                    : 'Converta, organize e edite documentos com privacidade e rapidez. Processamento seguro na nuvem com exclusão imediata dos arquivos — ou <a href="https://github.com/JoadsonRocha/localpdf.io/releases/download/1.0.0/LocalPDF.msi" target="_blank" style="color:#2563eb;font-weight:700;text-decoration:underline;">baixe o App Desktop</a> para uso 100% local no seu PC.';
            }
            const footerDesc = document.getElementById('footer-desc');
            if (footerDesc) {
                footerDesc.innerHTML = currentLanguage === 'en'
                    ? 'Free, secure and private PDF tools. In cloud mode (Railway), data is processed in ephemeral memory and deleted immediately. In desktop mode, 100% offline.'
                    : 'Ferramentas PDF gratuitas, seguras e privadas. Na nuvem (Railway), os dados são processados em memória volátil e apagados imediatamente. No app desktop, 100% offline.';
            }
        }

        updateWebModeElements();
        translatePage();

        // ── Skeleton → visible animation for tool cards ───────────────
        (function animateToolCards() {
            const cards = document.querySelectorAll('.tools-grid .tool-card');
            cards.forEach((card, i) => {
                card.classList.add('skeleton-card');
                setTimeout(() => {
                    card.classList.remove('skeleton-card');
                    card.classList.add('card-visible');
                }, 60 + i * 45);
            });
        })();

        // ── History State & Client-side Routing ──────────────────────
        (function initHistoryState() {
            const path = window.location.pathname.replace(/^[/]+|[/]+$/g, '');
            let initialViewState = { localpdf: true, view: 'home' };
            if (path === 'editor' || path === 'tool/edit-pdf' || path === 'tools/edit-pdf') {
                initialViewState = { localpdf: true, view: 'editor' };
            } else if (path === 'about' || path === 'sobre') {
                initialViewState = { localpdf: true, view: 'about' };
            } else if (path.startsWith('tool/') || path.startsWith('tools/')) {
                const name = path.replace(/^tools?[/]/, '');
                if (tools[name]) {
                    initialViewState = { localpdf: true, view: 'tool', tool: name };
                }
            }
            if (!history.state) {
                history.replaceState(initialViewState, '', window.location.href);
            }
        })();

        const initialToolFromRoute = "{{ initial_tool or '' }}";
        if (initialToolFromRoute) {
            if (initialToolFromRoute === 'edit-pdf' || initialToolFromRoute === 'editor') {
                showEditor(false);
            } else if (initialToolFromRoute === 'about' || initialToolFromRoute === 'sobre') {
                showAbout(false);
            } else if (tools[initialToolFromRoute]) {
                showTool(initialToolFromRoute, false);
            }
        }

        window.addEventListener('popstate', (event) => {
            const state = event.state;
            const path = window.location.pathname.replace(/^[/]+|[/]+$/g, '');

            if (state && state.view) {
                if (state.view === 'home') {
                    showHome(false);
                } else if (state.view === 'tool' && state.tool) {
                    showTool(state.tool, false);
                } else if (state.view === 'editor') {
                    showEditor(false);
                } else if (state.view === 'about') {
                    showAbout(false);
                }
                return;
            }

            // Fallback URL routing if state is null
            if (!path) {
                showHome(false);
            } else if (path === 'editor' || path === 'tool/edit-pdf' || path === 'tools/edit-pdf') {
                showEditor(false);
            } else if (path === 'about' || path === 'sobre') {
                showAbout(false);
            } else if (path.startsWith('tool/') || path.startsWith('tools/')) {
                const name = path.replace(/^tools?[/]/, '');
                if (tools[name]) {
                    showTool(name, false);
                } else {
                    showHome(false);
                }
            } else {
                showHome(false);
            }
        });
    </script>
</body>
</html>
"""


def is_web_environment():
    """Detects whether LocalPDF is running on Railway, a cloud container, or web host."""
    mode = os.environ.get("LOCALPDF_MODE", "").lower()
    if mode == "web":
        return True
    if mode == "local":
        return False
    if any(
        os.environ.get(k)
        for k in (
            "RAILWAY_ENVIRONMENT",
            "RAILWAY_PROJECT_ID",
            "RAILWAY_STATIC_URL",
            "RAILWAY_SERVICE_NAME",
        )
    ):
        return True
    try:
        host = request.host.split(":")[0].lower()
        if host not in ("localhost", "127.0.0.1", "0.0.0.0", "testserver") and not host.endswith(".local"):
            return True
    except Exception:
        pass
    return False


@app.after_request
def add_privacy_and_security_headers(response):
    """Adds security headers and disables browser/proxy caching for processed documents."""
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    if request.path in (
        "/convert",
        "/editor/export",
        "/editor/preview",
        "/preview-page",
        "/api/desktop/save",
        "/api/desktop/write-file",
        "/api/desktop/open-folder",
        "/api/desktop/open-file",
    ):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


@app.route("/robots.txt")
def robots_txt():
    content = (
        "User-agent: *\n"
        "Allow: /\n"
        "Disallow: /convert\n"
        "Disallow: /preview-page\n"
        "Disallow: /editor/preview\n"
        "Disallow: /editor/export\n"
        "Disallow: /healthz\n"
        "Disallow: /splash\n"
        "Sitemap: https://localpdf.io/sitemap.xml\n"
    )
    return app.response_class(content, mimetype="text/plain")


@app.route("/sitemap.xml")
def sitemap_xml():
    today = date.today().isoformat()
    urls = [
        {"loc": "https://localpdf.io/", "priority": "1.0", "changefreq": "daily"},
        {"loc": "https://localpdf.io/editor", "priority": "0.9", "changefreq": "weekly"},
        {"loc": "https://localpdf.io/about", "priority": "0.8", "changefreq": "monthly"},
    ]
    for tool_key in SEO_CONFIG.keys():
        if tool_key not in ("home", "editor", "about"):
            urls.append({
                "loc": f"https://localpdf.io/tool/{tool_key}",
                "priority": "0.9",
                "changefreq": "weekly"
            })

    xml_entries = "\n".join(
        f"  <url>\n    <loc>{u['loc']}</loc>\n    <lastmod>{today}</lastmod>\n    <changefreq>{u['changefreq']}</changefreq>\n    <priority>{u['priority']}</priority>\n  </url>"
        for u in urls
    )
    xml_content = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        f"{xml_entries}\n"
        "</urlset>"
    )
    return app.response_class(xml_content, mimetype="application/xml")


@app.route("/")
@app.route("/tool/<tool_name>")
@app.route("/tools/<tool_name>")
@app.route("/editor")
@app.route("/about")
@app.route("/sobre")
def index(tool_name=None):
    raw_path = request.path.rstrip("/")
    if raw_path in ("/editor", "/tool/edit-pdf", "/tools/edit-pdf"):
        seo_key = "editor"
        client_initial_tool = "edit-pdf"
    elif raw_path in ("/about", "/sobre"):
        seo_key = "about"
        client_initial_tool = "about"
    elif tool_name:
        seo_key = tool_name.lower()
        client_initial_tool = tool_name.lower()
    else:
        seo_key = "home"
        client_initial_tool = ""

    seo_data = get_seo_metadata(seo_key)
    web_mode = is_web_environment()
    return render_template_string(
        HTML_TEMPLATE,
        initial_tool=client_initial_tool,
        is_web_mode=web_mode,
        seo=seo_data,
    )


@app.route("/healthz")
def healthz():
    """Health-check endpoint used by the splash screen to detect when the server is ready."""
    return jsonify({"status": "ok"})


SPLASH_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LocalPDF.io — Iniciando...</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: "Segoe UI", "Avenir Next", sans-serif;
            background: linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 50%, #1e40af 100%);
            min-height: 100vh;
            display: flex; flex-direction: column;
            align-items: center; justify-content: center;
            color: white; user-select: none;
        }
        .brand { font-size: 2.4rem; font-weight: 800; letter-spacing: -0.04em; margin-bottom: 6px; }
        .brand span { color: #93c5fd; }
        .brand small { font-size: 0.85em; font-weight: 600; }
        .tagline { font-size: 0.95rem; color: rgba(255,255,255,0.7); margin-bottom: 48px; }
        .spinner-ring {
            width: 52px; height: 52px;
            border: 4px solid rgba(255,255,255,0.18);
            border-top-color: white;
            border-radius: 50%;
            animation: spin 0.9s linear infinite;
            margin-bottom: 24px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .status { font-size: 0.88rem; color: rgba(255,255,255,0.6); letter-spacing: 0.03em; }
        .dots::after {
            content: '';
            animation: dots 1.5s steps(4, end) infinite;
        }
        @keyframes dots {
            0%   { content: ''; }
            25%  { content: '.'; }
            50%  { content: '..'; }
            75%  { content: '...'; }
            100% { content: ''; }
        }
    </style>
</head>
<body>
    <div class="brand">local<span>pdf</span><small>.io</small></div>
    <p class="tagline">Privado. Local. Sem complicação.</p>
    <div class="spinner-ring"></div>
    <p class="status">Iniciando o servidor<span class="dots"></span></p>
    <script>
        (function poll() {
            fetch('/healthz', { cache: 'no-store' })
                .then(r => r.ok ? (window.location.replace('/')) : setTimeout(poll, 400))
                .catch(() => setTimeout(poll, 400));
        })();
    </script>
</body>
</html>
"""


@app.route("/splash")
def splash():
    """Animated splash screen shown while the app starts; auto-redirects to / when ready."""
    from flask import Response
    return Response(SPLASH_HTML, mimetype="text/html")


@app.route("/preview-page", methods=["POST"])
def preview_page():
    """Renders the first page of an uploaded PDF as a base64 JPEG thumbnail.
    Used by the client to show a preview in the file list before conversion.
    """
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400
    filename = secure_filename(file.filename)
    if not filename.lower().endswith(".pdf"):
        return jsonify({"error": "Apenas PDFs são suportados."}), 400
    temp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(temp_dir, filename)
        file.save(pdf_path)
        with fitz.open(pdf_path) as doc:
            if not doc.page_count:
                return jsonify({"error": "PDF sem páginas."}), 400
            page = doc.load_page(0)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(0.35, 0.35), alpha=False)
            data = base64.b64encode(pixmap.tobytes("jpeg", jpg_quality=72)).decode("ascii")
        return jsonify({"thumbnail": f"data:image/jpeg;base64,{data}"})
    except Exception:
        app.logger.exception("Falha ao gerar preview da página")
        return jsonify({"error": "Não foi possível gerar o preview."}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
# ── Desktop Integration Endpoints (Local Mode Only) ─────────────────────────
DESKTOP_ALLOWED_EXTENSIONS = ALLOWED_EXTENSIONS | {"zip"}


def is_safe_desktop_path(target_path):
    """Ensures destination path is a safe user path and not a restricted system directory."""
    if not target_path:
        return False
    norm = os.path.abspath(target_path)
    # Block system directories
    for env_var in ("WINDIR", "SystemRoot", "ProgramFiles", "ProgramFiles(x86)", "ProgramData"):
        sys_dir = os.environ.get(env_var)
        if sys_dir and norm.lower().startswith(os.path.abspath(sys_dir).lower()):
            return False
    # Validate extension
    ext = os.path.splitext(norm)[1].lstrip(".").lower()
    if ext not in DESKTOP_ALLOWED_EXTENSIONS:
        return False
    return True


@app.route("/api/desktop/save", methods=["POST"])
def desktop_save_to_downloads():
    """Saves the converted document directly to the user's Downloads directory."""
    if is_web_environment():
        return jsonify({"error": "Disponível apenas no modo desktop"}), 403

    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"error": "Arquivo inválido"}), 400

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lstrip(".").lower()
    if ext not in DESKTOP_ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Extensão não permitida: {ext}"}), 400

    downloads_dir = pathlib.Path.home() / "Downloads"
    if not downloads_dir.exists():
        downloads_dir = pathlib.Path.home()

    target = downloads_dir / filename
    stem = target.stem
    suffix = target.suffix
    c = 1
    while target.exists():
        target = downloads_dir / f"{stem} ({c}){suffix}"
        c += 1

    file.save(str(target))
    return jsonify({
        "success": True,
        "path": str(target),
        "filename": target.name,
        "directory": str(downloads_dir),
    })


@app.route("/api/desktop/write-file", methods=["POST"])
def desktop_write_to_path():
    """Writes the converted document to the specific path chosen by the user in the Save dialog."""
    if is_web_environment():
        return jsonify({"error": "Disponível apenas no modo desktop"}), 403

    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    target_path = request.form.get("path", "").strip()
    if not target_path or not is_safe_desktop_path(target_path):
        return jsonify({"error": "Caminho de salvamento inválido ou não permitido"}), 400

    file = request.files["file"]
    if not file:
        return jsonify({"error": "Arquivo inválido"}), 400

    parent_dir = os.path.dirname(target_path)
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    file.save(target_path)
    return jsonify({
        "success": True,
        "path": os.path.abspath(target_path),
        "filename": os.path.basename(target_path),
    })


@app.route("/api/desktop/open-folder", methods=["POST"])
def desktop_open_folder():
    """Highlights the saved file in Windows Explorer."""
    if is_web_environment():
        return jsonify({"error": "Disponível apenas no modo desktop"}), 403

    data = request.get_json(silent=True) or {}
    path = data.get("path", "").strip()
    if not path or not os.path.exists(path):
        return jsonify({"error": "Arquivo não encontrado"}), 404

    try:
        import subprocess

        abs_path = os.path.abspath(path)
        subprocess.Popen(["explorer.exe", f"/select,{abs_path}"])
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/desktop/open-file", methods=["POST"])
def desktop_open_file():
    """Opens the saved file in the Windows default application."""
    if is_web_environment():
        return jsonify({"error": "Disponível apenas no modo desktop"}), 403

    data = request.get_json(silent=True) or {}
    path = data.get("path", "").strip()
    if not path or not os.path.exists(path):
        return jsonify({"error": "Arquivo não encontrado"}), 404

    ext = os.path.splitext(path)[1].lstrip(".").lower()
    if ext not in DESKTOP_ALLOWED_EXTENSIONS:
        return jsonify({"error": "Extensão não permitida"}), 400

    try:
        os.startfile(os.path.abspath(path))
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/favicon.svg")
def favicon():
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    favicon_path = os.path.join(base_dir, "favicon.svg")
    if not os.path.exists(favicon_path):
        favicon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.svg")
    with open(favicon_path, "rb") as f:
        data = f.read()
    return send_file(io.BytesIO(data), mimetype="image/svg+xml")


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
        import gc
        gc.collect()


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
        first_base = (
            os.path.splitext(secure_filename(files[0].filename))[0]
            if files and files[0].filename
            else "documento"
        )
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=f"{first_base}_editado.pdf",
        )
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Falha ao exportar PDF do editor")
        return jsonify({"error": "Não foi possível exportar o PDF."}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        import gc
        gc.collect()


def format_excel_cell_value(val_data, val_formula):
    """Converte e formata valores de células do Excel para exibição profissional no PDF."""
    if val_data is not None:
        if isinstance(val_data, datetime):
            if val_data.hour == 0 and val_data.minute == 0 and val_data.second == 0:
                return val_data.strftime("%d/%m/%Y")
            return val_data.strftime("%d/%m/%Y %H:%M")
        if isinstance(val_data, date):
            return val_data.strftime("%d/%m/%Y")
        if isinstance(val_data, float):
            if val_data.is_integer():
                return str(int(val_data))
            return f"{val_data:.2f}"
        if isinstance(val_data, bool):
            return "VERDADEIRO" if val_data else "FALSO"
        s = str(val_data).strip()
        if not s.startswith("<openpyxl."):
            return s

    if val_formula is not None:
        if hasattr(val_formula, "text"):
            text = getattr(val_formula, "text", "")
            return f"={text}" if text and not text.startswith("=") else str(text)
        if isinstance(val_formula, bool):
            return "VERDADEIRO" if val_formula else "FALSO"
        s = str(val_formula).strip()
        if not s.startswith("<openpyxl."):
            return s

    return ""


def excel_to_pdf(file, temp_dir):
    """
    Converte uma planilha Excel (.xlsx) em um documento PDF estruturado e profissional.
    Suporta múltiplas abas, lê valores calculados e fórmulas, remove margens vazias,
    e renderiza tabelas estilizadas com orientação dinâmica (Paisagem / Retrato).
    """
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "planilha"
    xlsx_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(xlsx_path)

    pdf_path = os.path.join(temp_dir, f"{base_name}.pdf")

    wb_data = None
    wb_raw = None
    try:
        wb_data = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=False)
    except Exception:
        pass

    try:
        wb_raw = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=False)
    except Exception:
        pass

    workbook = wb_data or wb_raw
    if not workbook:
        raise ValueError("Não foi possível carregar o arquivo de planilha Excel.")

    sheet_names = workbook.sheetnames

    # Determinar maior número de colunas para definir orientação ideal
    max_overall_cols = 1
    for name in sheet_names:
        ws_d = wb_data[name] if wb_data and name in wb_data.sheetnames else None
        ws_r = wb_raw[name] if wb_raw and name in wb_raw.sheetnames else None
        c_d = ws_d.max_column if ws_d and ws_d.max_column else 1
        c_r = ws_r.max_column if ws_r and ws_r.max_column else 1
        max_overall_cols = max(max_overall_cols, c_d, c_r)

    use_landscape = max_overall_cols > 6
    page_size = landscape(A4) if use_landscape else A4
    page_width, _ = page_size
    margin = 28
    usable_width = page_width - (2 * margin)

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=page_size,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
    )

    styles = getSampleStyleSheet()

    # Ajuste dinâmico de fonte para caber mais colunas com legibilidade
    if max_overall_cols <= 6:
        f_size, f_leading = 9, 11
    elif max_overall_cols <= 10:
        f_size, f_leading = 8, 10
    elif max_overall_cols <= 16:
        f_size, f_leading = 7, 8.5
    else:
        f_size, f_leading = 6, 7.5

    title_style = ParagraphStyle(
        "ExcelSheetTitle",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1e40af"),
        spaceAfter=6,
        spaceBefore=0,
    )

    cell_style = ParagraphStyle(
        "ExcelCellText",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=f_size,
        leading=f_leading,
        textColor=colors.HexColor("#1e293b"),
        wordWrap="CJK",
    )

    cell_header_style = ParagraphStyle(
        "ExcelCellHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=f_size,
        leading=f_leading,
        textColor=colors.HexColor("#0f172a"),
        wordWrap="CJK",
    )

    empty_style = ParagraphStyle(
        "ExcelEmptySheet",
        parent=styles["Italic"],
        fontName="Helvetica-Oblique",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#94a3b8"),
    )

    elements = []

    for sheet_idx, sheet_name in enumerate(sheet_names):
        if sheet_idx > 0:
            elements.append(PageBreak())

        elements.append(Paragraph(f"<b>Planilha:</b> {escape(sheet_name)}", title_style))
        elements.append(Spacer(1, 4))

        ws_d = wb_data[sheet_name] if wb_data and sheet_name in wb_data.sheetnames else None
        ws_r = wb_raw[sheet_name] if wb_raw and sheet_name in wb_raw.sheetnames else None

        cur_rows = max(ws_d.max_row or 1 if ws_d else 1, ws_r.max_row or 1 if ws_r else 1)
        cur_cols = max(ws_d.max_column or 1 if ws_d else 1, ws_r.max_column or 1 if ws_r else 1)

        matrix = []
        for r in range(1, cur_rows + 1):
            row_cells = []
            for c in range(1, cur_cols + 1):
                vd = ws_d.cell(row=r, column=c).value if ws_d else None
                vr = ws_r.cell(row=r, column=c).value if ws_r else None
                row_cells.append(format_excel_cell_value(vd, vr))
            matrix.append(row_cells)

        # Remove linhas vazias nas bordas
        non_empty_rows = [r for r, row in enumerate(matrix) if any(v != "" for v in row)]
        if not non_empty_rows:
            elements.append(Paragraph("<i>(Esta planilha não contém dados preenchidos)</i>", empty_style))
            continue

        min_r, max_r = non_empty_rows[0], non_empty_rows[-1]

        # Remove colunas vazias nas bordas
        non_empty_cols = [
            c for c in range(cur_cols) if any(matrix[r][c] != "" for r in range(min_r, max_r + 1))
        ]
        if not non_empty_cols:
            elements.append(Paragraph("<i>(Esta planilha não contém dados preenchidos)</i>", empty_style))
            continue

        min_c, max_c = non_empty_cols[0], non_empty_cols[-1]

        trimmed = []
        for r in range(min_r, max_r + 1):
            trimmed.append(matrix[r][min_c : max_c + 1])

        num_cols = max_c - min_c + 1

        # Calcular larguras de coluna proporcionais
        col_lens = []
        for c in range(num_cols):
            max_l = max(len(trimmed[r][c]) for r in range(len(trimmed)))
            col_lens.append(max(max_l, 3))

        total_weight = sum(col_lens)
        col_widths = [(w / total_weight) * usable_width for w in col_lens]

        min_col_w = 20
        col_widths = [max(w, min_col_w) for w in col_widths]
        if sum(col_widths) > usable_width:
            norm_factor = usable_width / sum(col_widths)
            col_widths = [w * norm_factor for w in col_widths]

        table_rows = []
        for r_idx, row in enumerate(trimmed):
            row_cells = []
            is_first = (r_idx == 0)
            for val in row:
                st = cell_header_style if is_first else cell_style
                safe_text = escape(val) if val else "&nbsp;"
                row_cells.append(Paragraph(safe_text, st))
            table_rows.append(row_cells)

        t = Table(table_rows, colWidths=col_widths, repeatRows=1)
        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eff6ff")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        for r in range(1, len(table_rows)):
            if r % 2 == 1:
                t_style.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#f8fafc")))

        t.setStyle(TableStyle(t_style))
        elements.append(t)

    doc.build(elements)

    if wb_data:
        try:
            wb_data.close()
        except Exception:
            pass
    if wb_raw:
        try:
            wb_raw.close()
        except Exception:
            pass

    return [pdf_path]


def txt_to_pdf(file, temp_dir):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "texto"
    txt_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(txt_path)

    pdf_path = os.path.join(temp_dir, f"{base_name}.pdf")
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
        first_base = os.path.splitext(secure_filename(files[0].filename))[0] if files and files[0].filename else "localpdf"

        if tool == "pdf-to-images":
            output_files = pdf_to_images(files[0], temp_dir, img_format="png")
            zip_name = f"{first_base}_imagens.zip"
        elif tool == "pdf-to-png":
            output_files = pdf_to_images(files[0], temp_dir, img_format="png")
            zip_name = f"{first_base}_png.zip"
        elif tool == "pdf-to-jpg":
            output_files = pdf_to_images(files[0], temp_dir, img_format="jpg")
            zip_name = f"{first_base}_jpg.zip"
        elif tool == "unlock-pdf":
            output_files = unlock_pdf(files[0], temp_dir, request.form.get("password", ""))
            zip_name = f"{first_base}_desbloqueado.pdf"
        elif tool == "pdf-to-excel":
            output_files = pdf_to_excel(files[0], temp_dir)
            zip_name = f"{first_base}.xlsx"
        elif tool == "images-to-pdf":
            output_files = images_to_pdf(files, temp_dir)
            zip_name = f"{first_base}_convertido.pdf"
        elif tool == "merge-pdf":
            output_files = merge_pdfs(files, temp_dir)
            zip_name = f"{first_base}_mesclado.pdf"
        elif tool == "split-pdf":
            output_files = split_pdf(files[0], temp_dir, page_range=request.form.get("page_range", ""))
            zip_name = f"{first_base}_paginas.zip"
        elif tool == "compress-pdf":
            output_files = compress_pdf(files[0], temp_dir)
            zip_name = f"{first_base}_comprimido.pdf"
        elif tool == "protect-pdf":
            output_files = protect_pdf(files[0], temp_dir, request.form.get("password", ""))
            zip_name = f"{first_base}_protegido.pdf"
        elif tool == "watermark-pdf":
            output_files = watermark_pdf(
                files[0],
                temp_dir,
                request.form.get("watermark_text", ""),
                position=request.form.get("watermark_position", "diagonal"),
                color_style=request.form.get("watermark_color", "gray"),
                opacity=request.form.get("watermark_opacity", "0.22"),
            )
            zip_name = f"{first_base}_marca_dagua.pdf"
        elif tool == "page-numbers-pdf":
            output_files = page_numbers_pdf(
                files[0], temp_dir, request.form.get("page_number_position", "bottom-center")
            )
            zip_name = f"{first_base}_numerado.pdf"
        elif tool == "pdf-to-pdfa":
            output_files = pdf_to_pdfa(files, temp_dir)
            zip_name = f"{first_base}_pdfa.zip"
        elif tool == "word-to-pdf":
            output_files = word_to_pdf(files, temp_dir)
            zip_name = f"{first_base}.pdf"
        elif tool == "excel-to-pdf":
            output_files = excel_to_pdf(files[0], temp_dir)
            zip_name = f"{first_base}.pdf"
        elif tool == "txt-to-pdf":
            output_files = txt_to_pdf(files[0], temp_dir)
            zip_name = f"{first_base}.pdf"
        elif tool == "pdf-to-word":
            output_files = pdf_to_word(files[0], temp_dir)
            zip_name = f"{first_base}.docx"
        elif tool == "pdf-to-text":
            output_files = pdf_to_text(files[0], temp_dir)
            zip_name = f"{first_base}.txt"
        elif tool == "ocr-pdf":
            output_files = ocr_pdf(files[0], temp_dir)
            zip_name = f"{first_base}_ocr.txt"
        else:
            return jsonify({"error": "Ferramenta não suportada"}), 400

        response = build_response(output_files, temp_dir, default_zip_name=zip_name)
        return response
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except RuntimeError as error:
        return jsonify({"error": str(error)}), 400
    except Exception as error:
        app.logger.exception("Falha ao processar ferramenta %s", tool)
        return jsonify({"error": f"Não foi possível processar os arquivos: {str(error)}"}), 500
    finally:
        # Diretório temporário limpo após preparar resposta (BytesIO) evitando remoção antecipada
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        import gc
        gc.collect()


def images_to_pdf(files, temp_dir):
    """
    Combina uma ou múltiplas imagens (JPG, PNG) em um único arquivo PDF.
    """
    if not isinstance(files, list):
        files = [files]
    if not files:
        raise ValueError("Nenhuma imagem enviada para conversão.")

    first_base = os.path.splitext(secure_filename(files[0].filename))[0] if files[0].filename else "imagens"
    images = []
    for file in files:
        img_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(img_path)
        img = Image.open(img_path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)

    pdf_path = os.path.join(temp_dir, f"{first_base}_convertido.pdf")
    images[0].save(pdf_path, save_all=True, append_images=images[1:])
    return [pdf_path]


def merge_pdfs(files, temp_dir):
    """
    Combina múltiplos arquivos PDF em um único documento PDF contínuo.
    """
    if not isinstance(files, list):
        files = [files]
    if len(files) < 1:
        raise ValueError("Envie pelo menos um arquivo PDF para mesclagem.")

    first_base = os.path.splitext(secure_filename(files[0].filename))[0] if files[0].filename else "documento"
    merged_doc = fitz.open()

    for file in files:
        pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(pdf_path)
        doc = fitz.open(pdf_path)
        merged_doc.insert_pdf(doc)
        doc.close()

    output_path = os.path.join(temp_dir, f"{first_base}_mesclado.pdf")
    merged_doc.save(output_path)
    merged_doc.close()
    return [output_path]


def pdf_to_images(file, temp_dir, img_format="png"):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    output_files = []
    fmt = "jpg" if img_format.lower() in ("jpg", "jpeg") else "png"

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x resolution
        img_path = os.path.join(temp_dir, f"{base_name}_pagina_{page_num + 1}.{fmt}")
        if fmt == "jpg":
            pix.pil_save(img_path, format="JPEG", quality=90)
        else:
            pix.save(img_path)
        output_files.append(img_path)

    doc.close()
    return output_files


def parse_page_range(range_str, total_pages):
    """
    Analisa strings como '1-3, 5, 8-10' e retorna uma lista ordenada de índices 0-based válidos.
    """
    pages = set()
    parts = range_str.split(",")
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            sub = part.split("-")
            if len(sub) == 2 and sub[0].strip().isdigit() and sub[1].strip().isdigit():
                start = int(sub[0].strip())
                end = int(sub[1].strip())
                for p in range(min(start, end), max(start, end) + 1):
                    if 1 <= p <= total_pages:
                        pages.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p - 1)
    return sorted(list(pages))


def split_pdf(file, temp_dir, page_range=""):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    output_files = []

    if page_range and page_range.strip():
        selected_pages = parse_page_range(page_range.strip(), total_pages)
        if not selected_pages:
            doc.close()
            raise ValueError(
                f"Nenhuma página válida encontrada no intervalo '{page_range}'. O documento possui {total_pages} página(s)."
            )
        new_doc = fitz.open()
        for page_idx in selected_pages:
            new_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
        output_path = os.path.join(temp_dir, f"{base_name}_extraido.pdf")
        new_doc.save(output_path)
        new_doc.close()
        output_files.append(output_path)
    else:
        for page_num in range(total_pages):
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            output_path = os.path.join(temp_dir, f"{base_name}_pagina_{page_num + 1}.pdf")
            new_doc.save(output_path)
            new_doc.close()
            output_files.append(output_path)

    doc.close()
    return output_files


def unlock_pdf(file, temp_dir, password=""):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    if doc.is_encrypted:
        if not password:
            doc.close()
            raise ValueError("O documento está protegido por senha. Por favor, digite a senha para desbloqueá-lo.")
        auth = doc.authenticate(password)
        if not auth:
            doc.close()
            raise ValueError("Senha incorreta. Não foi possível desbloquear o PDF.")

    output_path = os.path.join(temp_dir, f"{base_name}_desbloqueado.pdf")
    clean_doc = fitz.open()
    clean_doc.insert_pdf(doc)
    clean_doc.save(output_path)
    clean_doc.close()
    doc.close()
    return [output_path]


def pdf_to_excel(file, temp_dir):
    """Extrai tabelas de um documento PDF para uma planilha Excel (.xlsx)."""
    if pdfplumber is None:
        raise RuntimeError("A biblioteca pdfplumber não está instalada no sistema.")

    import re
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "planilha"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove sheet inicial vazia

    has_data = False
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if tables:
                for t_idx, table in enumerate(tables, start=1):
                    sheet_name = f"Pág {page_idx} Tab {t_idx}" if len(tables) > 1 else f"Página {page_idx}"
                    ws = wb.create_sheet(title=sheet_name[:31])
                    has_data = True
                    for row in table:
                        clean_row = [
                            cell.replace("\n", " ").strip() if isinstance(cell, str) else cell
                            for cell in row
                        ]
                        ws.append(clean_row)
            else:
                text = page.extract_text()
                if text and text.strip():
                    ws = wb.create_sheet(title=f"Página {page_idx}"[:31])
                    has_data = True
                    for line in text.split("\n"):
                        cells = re.split(r"\s{2,}|\t", line.strip())
                        ws.append(cells)

    if not has_data:
        ws = wb.create_sheet(title="Dados")
        ws.append(["Nenhuma tabela ou texto detectado no documento PDF."])

    output_path = os.path.join(temp_dir, f"{base_name}.xlsx")
    wb.save(output_path)
    return [output_path]


def compress_pdf(file, temp_dir):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)

    doc = fitz.open(pdf_path)
    output_path = os.path.join(temp_dir, f"{base_name}_comprimido.pdf")
    doc.save(output_path, garbage=4, deflate=True, clean=True)
    doc.close()

    return [output_path]


def protect_pdf(file, temp_dir, password):
    """Cria uma cópia criptografada do PDF usando senha local."""
    if len(password) < 4:
        raise ValueError("A senha precisa ter pelo menos 4 caracteres.")

    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, f"{base_name}_protegido.pdf")

    with fitz.open(pdf_path) as document:
        document.save(
            output_path,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=password,
            user_pw=password,
            permissions=fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT,
        )

    return [output_path]


def watermark_pdf(file, temp_dir, text, position="diagonal", color_style="gray", opacity="0.22"):
    """
    Aplica marca d'água profissional com rotação diagonal e transparência suave,
    garantindo que o conteúdo original do PDF continue perfeitamente legível.
    """
    text = (text or "").strip()
    if not text:
        raise ValueError("Informe o texto da marca d'água.")

    try:
        opac_val = float(opacity)
        opac_val = max(0.05, min(0.9, opac_val))
    except (ValueError, TypeError):
        opac_val = 0.22

    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, f"{base_name}_marca_dagua.pdf")

    colors = {
        "gray": (0.45, 0.45, 0.48),
        "red": (0.80, 0.20, 0.20),
        "blue": (0.15, 0.35, 0.75),
    }
    col = colors.get(color_style, (0.45, 0.45, 0.48))

    with fitz.open(pdf_path) as document:
        for page in document:
            rect = page.rect
            width, height = rect.width, rect.height
            center = fitz.Point(width / 2, height / 2)

            if position in ("diagonal", "center-diagonal", "center"):
                if position == "center":
                    fontsize = min(48.0, max(20.0, width / (max(len(text), 6) * 0.55)))
                    text_len = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
                    start_point = fitz.Point(center.x - text_len / 2, center.y + fontsize / 3)
                    page.insert_text(start_point, text, fontname="helv", fontsize=fontsize, color=col, fill_opacity=opac_val, overlay=True)
                else:
                    fontsize = min(54.0, max(22.0, width / (max(len(text), 6) * 0.45)))
                    text_len = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
                    start_point = fitz.Point(center.x - text_len / 2, center.y + fontsize / 3)
                    matrix = fitz.Matrix(-45)
                    page.insert_text(start_point, text, fontname="helv", fontsize=fontsize, color=col, morph=(center, matrix), fill_opacity=opac_val, overlay=True)
            elif position == "top":
                fontsize = 16.0
                text_len = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
                start_point = fitz.Point(center.x - text_len / 2, 45)
                page.insert_text(start_point, text, fontname="helv", fontsize=fontsize, color=col, fill_opacity=min(0.7, opac_val * 1.5), overlay=True)
            elif position == "bottom":
                fontsize = 16.0
                text_len = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
                start_point = fitz.Point(center.x - text_len / 2, height - 35)
                page.insert_text(start_point, text, fontname="helv", fontsize=fontsize, color=col, fill_opacity=min(0.7, opac_val * 1.5), overlay=True)
            else:
                fontsize = min(50.0, max(22.0, width / (max(len(text), 6) * 0.45)))
                text_len = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
                start_point = fitz.Point(center.x - text_len / 2, center.y + fontsize / 3)
                matrix = fitz.Matrix(-45)
                page.insert_text(start_point, text, fontname="helv", fontsize=fontsize, color=col, morph=(center, matrix), fill_opacity=opac_val, overlay=True)

        document.save(output_path)

    return [output_path]


def page_numbers_pdf(file, temp_dir, position):
    """Adiciona numeração simples a todas as páginas do PDF."""
    positions = {"bottom-center", "bottom-right", "top-center", "top-right"}
    if position not in positions:
        raise ValueError("Posição de numeração inválida.")

    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, f"{base_name}_numerado.pdf")

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
    first_base = os.path.splitext(secure_filename(files[0].filename))[0] if files and files[0].filename else "documento"
    pdf_path = os.path.join(temp_dir, f"{first_base}.pdf")
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


def _ocr_image(img):
    try:
        return pytesseract.image_to_string(img, lang="por+eng")
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Tesseract OCR não foi encontrado. Instale o Tesseract ou configure o caminho no ambiente."
        )
    except pytesseract.TesseractError as err:
        err_msg = str(err)
        if "traineddata" in err_msg:
            for fallback_lang in ("por", "eng"):
                try:
                    return pytesseract.image_to_string(img, lang=fallback_lang)
                except Exception:
                    continue
            raise RuntimeError(
                "Dados de idioma do Tesseract (por.traineddata ou eng.traineddata) não encontrados no diretório tessdata."
            )
        raise RuntimeError(f"Erro durante processamento OCR: {err_msg}")


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
                    text = _ocr_image(img)
                extracted_text.append(f"--- Página {page_num + 1} ---\n{text}")
    elif ext in ("jpg", "jpeg", "png"):
        # Aplicar OCR diretamente na imagem
        with Image.open(input_path) as img:
            text = _ocr_image(img)
        extracted_text.append(text)
    else:
        raise RuntimeError(f"Formato não suportado para OCR: {ext}")

    # Salvar texto extraído em arquivo TXT
    base_name = os.path.splitext(filename)[0]
    txt_path = os.path.join(temp_dir, f"{base_name}_ocr.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(extracted_text))

    return [txt_path]


def build_response(output_files, temp_dir, download_name=None, default_zip_name=None):
    """Monta resposta enviando arquivos como attachment sem risco de remoção prematura do diretório temporário."""
    if len(output_files) == 1:
        file_path = output_files[0]
        filename = download_name or os.path.basename(file_path)
        with open(file_path, "rb") as f:
            data = f.read()
        return send_file(io.BytesIO(data), as_attachment=True, download_name=filename)
    else:
        zip_filename = default_zip_name or download_name or "localpdf_arquivos.zip"
        if not zip_filename.endswith(".zip"):
            zip_filename += ".zip"
        zip_path = os.path.join(temp_dir, zip_filename)
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file_path in output_files:
                zipf.write(file_path, os.path.basename(file_path))
        with open(zip_path, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data), as_attachment=True, download_name=zip_filename
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    default_host = "127.0.0.1" if os.environ.get("LOCALPDF_MODE") == "local" else "0.0.0.0"
    host = os.environ.get("LOCALPDF_HOST", default_host)
    app.run(host=host, port=port)
