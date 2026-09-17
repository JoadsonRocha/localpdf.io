import base64
import io
import json
import os
import shutil
import sys
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
        .tools-grid .tool-card { opacity: 0; transform: translateY(8px); transition: opacity 0.35s ease, transform 0.35s ease; }
        .tools-grid .tool-card.card-visible { opacity: 1; transform: translateY(0); }
        .tools-grid .tool-card.card-hidden { opacity: 0 !important; display: none !important; }

        .result-card { margin-top: 24px; padding: 22px; border-radius: 12px; text-align: left; animation: fadeIn 0.3s ease; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .result-card.success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
        .result-card.error { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
        .result-header { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
        .result-header h4 { font-size: 1.15rem; font-weight: 700; margin: 0; }
        .result-body p { font-size: 0.92rem; margin-bottom: 12px; opacity: 0.95; line-height: 1.5; }
        .result-filename { font-weight: 700; word-break: break-all; color: #0f172a; background: rgba(255,255,255,0.8); padding: 6px 12px; border-radius: 6px; display: inline-block; margin-bottom: 12px; border: 1px solid rgba(0,0,0,0.06); font-size: 0.9rem; }
        .result-actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }
        .btn-download-again { background: #16a34a; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 0.9rem; transition: background 0.2s; box-shadow: 0 4px 10px rgba(22,163,74,0.2); font-family: inherit; }
        .btn-download-again:hover { background: #15803d; }
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
        }
        .tech-pill strong {
            color: #0f172a;
        }
        @media (max-width: 700px) {
            .about-shell { padding: 22px 16px; }
            .about-hero h2 { font-size: 1.5rem; }
            .about-grid-pillars { grid-template-columns: 1fr; }
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
                <button id="language-toggle" class="language-toggle" type="button" onclick="toggleLanguage()">EN</button>
            </div>
        </nav>

        <div id="home-view">
            <div class="header">
                <h1>Trabalhe com seus PDFs sem complicação</h1>
                <p>Converta, organize e edite documentos diretamente no seu computador. Sem contas, sem nuvem e sem enviar seus arquivos para fora.</p>
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
            'Informe a senha do documento para desbloqueá-lo.': 'Enter document password to unlock it.'
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

        function showTool(toolName) {
            currentTool = toolName;
            const tool = tools[toolName];
            if (!tool) return;

            document.getElementById('home-view').classList.add('hidden');
            const editorView = document.getElementById('editor-view');
            if (editorView) editorView.classList.add('hidden');
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

            document.title = `${rawTitle} - LocalPDF.io`;
            if (window.location.pathname !== `/tool/${toolName}`) {
                history.pushState(null, '', `/tool/${toolName}`);
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function renderToolOptions(optionType) {
            const options = document.getElementById('options');
            if (optionType === 'password') {
                options.innerHTML = `
                    <label for="pdf-password">${t('Senha do PDF', 'PDF Password')}</label>
                    <input id="pdf-password" type="password" minlength="4" autocomplete="new-password" placeholder="${t('Digite uma senha com pelo menos 4 caracteres', 'Enter a password with at least 4 characters')}">
                    <small>${t('A senha é usada somente durante o processamento local.', 'The password is used only during local processing.')}</small>
                `;
                options.classList.remove('hidden');
                return;
            }
            if (optionType === 'watermark') {
                options.innerHTML = `
                    <label for="watermark-text">${t("Texto da marca d'água", 'Watermark text')}</label>
                    <input id="watermark-text" type="text" maxlength="80" placeholder="${t('Ex.: CONFIDENCIAL', 'e.g.: CONFIDENTIAL')}">
                    <label for="watermark-position">${t('Posição', 'Position')}</label>
                    <select id="watermark-position">
                        <option value="center">${t('Centro', 'Center')}</option>
                        <option value="top">${t('Parte superior', 'Top')}</option>
                        <option value="bottom">${t('Parte inferior', 'Bottom')}</option>
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
                options.innerHTML = `
                    <label for="pdf-unlock-password">${t('Senha do PDF para desbloquear', 'PDF Password to unlock')}</label>
                    <input id="pdf-unlock-password" type="password" autocomplete="current-password" placeholder="${t('Digite a senha atual do documento', 'Enter current document password')}">
                    <small>${t('A senha é usada apenas localmente na sua máquina para remover a proteção.', 'The password is used only locally on your computer to remove protection.')}</small>
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

        function showHome() {
            currentTool = null;
            document.getElementById('home-view').classList.remove('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            const editorView = document.getElementById('editor-view');
            if (editorView) editorView.classList.add('hidden');
            uploadedFiles = [];
            document.title = 'LocalPDF.io';
            if (window.location.pathname !== '/') {
                history.pushState(null, '', '/');
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

        function downloadAgain() {
            if (!lastDownloadedBlob) return;
            const url = window.URL.createObjectURL(lastDownloadedBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = lastDownloadedFilename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
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

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = downloadFilename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);

                    showToast(t('Arquivo processado e download iniciado!', 'File processed and download started!'), 'success');

                    document.getElementById('result').innerHTML = `
                        <div class="result-card success">
                            <div class="result-header">
                                <h4>✅ ${t('Concluído com sucesso!', 'Completed successfully!')}</h4>
                            </div>
                            <div class="result-body">
                                <span class="result-filename">📄 ${downloadFilename}</span>
                                <p>${t('O arquivo foi processado no seu computador e o download foi iniciado automaticamente.', 'The file was processed on your computer and the download started automatically.')}</p>
                                <div class="result-actions">
                                    <button type="button" class="btn-download-again" onclick="downloadAgain()">${t('📥 Baixar novamente', '📥 Download again')}</button>
                                    <button type="button" class="btn-reset-flow" onclick="resetToolFlow()">${t('✨ Processar outro arquivo', '✨ Process another file')}</button>
                                </div>
                            </div>
                        </div>
                    `;
                    document.getElementById('result').classList.remove('hidden');
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

        function showEditor() {
            currentTool = 'edit-pdf';
            document.getElementById('home-view').classList.add('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            document.getElementById('editor-view').classList.remove('hidden');
            resetEditor();
            document.title = (currentLanguage === 'en' ? 'Edit PDF' : 'Editar PDF') + ' - LocalPDF.io';
            if (window.location.pathname !== '/editor') {
                history.pushState(null, '', '/editor');
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function showHomeFromEditor() {
            currentTool = null;
            document.getElementById('editor-view').classList.add('hidden');
            document.getElementById('tool-views').classList.add('hidden');
            document.getElementById('home-view').classList.remove('hidden');
            resetEditor();
            document.title = 'LocalPDF.io';
            if (window.location.pathname !== '/') {
                history.pushState(null, '', '/');
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
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

        const initialToolFromRoute = "{{ initial_tool or '' }}";
        if (initialToolFromRoute) {
            if (initialToolFromRoute === 'edit-pdf' || initialToolFromRoute === 'editor') {
                showEditor();
            } else if (tools[initialToolFromRoute]) {
                showTool(initialToolFromRoute);
            }
        }

        window.addEventListener('popstate', () => {
            const path = window.location.pathname.replace(/^[/]+|[/]+$/g, '');
            if (!path) {
                showHome();
            } else if (path === 'editor' || path === 'tool/edit-pdf') {
                showEditor();
            } else if (path.startsWith('tool/')) {
                const name = path.replace('tool/', '');
                if (tools[name]) showTool(name);
            }
        });
    </script>
</body>
</html>
"""


@app.route("/")
@app.route("/tool/<tool_name>")
@app.route("/tools/<tool_name>")
@app.route("/editor")
def index(tool_name=None):
    if request.path.rstrip("/") == "/editor":
        tool_name = "edit-pdf"
    return render_template_string(HTML_TEMPLATE, initial_tool=tool_name or "")


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


def excel_to_pdf(file, temp_dir):
    base_name = os.path.splitext(secure_filename(file.filename))[0] or "planilha"
    xlsx_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(xlsx_path)

    pdf_path = os.path.join(temp_dir, f"{base_name}.pdf")
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
                request.form.get("watermark_position", "center"),
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
            shutil.rmtree(temp_dir)


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


def watermark_pdf(file, temp_dir, text, position):
    """Aplica uma marca d'água de texto em todas as páginas do PDF."""
    text = text.strip()
    if not text:
        raise ValueError("Informe o texto da marca d'água.")
    if position not in {"center", "top", "bottom"}:
        raise ValueError("Posição de marca d'água inválida.")

    base_name = os.path.splitext(secure_filename(file.filename))[0] or "documento"
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, f"{base_name}_marca_dagua.pdf")

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
