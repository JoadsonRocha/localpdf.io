# PDF In-Place Text Editor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enable visual, in-place text editing and new text insertion directly on PDF pages (WPS Office style), preserving vector layout, typography, and 100% local privacy.

**Architecture:** A dual-layer system. The backend leverages PyMuPDF (`fitz`) to extract text blocks and line spans with bounding boxes and font metadata (`/api/pdf/inspect-text`), and applies vector redaction and replacement (`/api/pdf/apply-text-edits`). The frontend provides an interactive canvas overlay where clicking any sentence turns it into an editable field, along with a toolbar to add new text, change font size/color, and save the updated vector PDF.

**Tech Stack:** Python 3.10+, Flask, PyMuPDF (`pymupdf`), HTML5 Canvas & Overlay DOM, CSS3 design tokens.

**Spec:** In-place text editing like WPS Office PDF, supporting typo correction, text updates, and text insertion on any page.

---

## Global Constraints

- Preserve 100% offline processing capability with PyMuPDF.
- Never rasterize entire pages to bitmaps for text editing; preserve vector PDF text and fonts.
- Preserve existing endpoints and test assertions (`tests/test_web_features.py` and `tests/test_web_environment.py`).
- Security: Sanitize all file uploads via `secure_filename` and enforce `.pdf` extension check.

## Review Focus

1. **Multi-page documents:** Edits on page 1 must not affect or corrupt pages 2..N.
2. **Text overlapping & glyph erasure:** The old text must be cleanly redacted so it doesn't show behind the replacement text.
3. **Special characters & Accents:** Portuguese accents (ç, á, é, í, ó, ú, ã, õ) must render without encoding corruption.
4. **Coordinate scaling:** Screen coordinate clicks on responsive canvas must map 1:1 to PDF point coordinates `[x0, y0, x1, y1]`.
5. **No text modified fallback:** If the user saves without edits, the original document remains intact.

---

### Task 1: Backend Text Inspection API (`/api/pdf/inspect-text`)

**Files:**
- Modify: `app.py`
- Test: `tests/test_pdf_text_editor.py`

- [x] Step 1: Write unit tests in `tests/test_pdf_text_editor.py` testing `/api/pdf/inspect-text` for extracting text spans with bounding boxes, font size, and color.
- [x] Step 2: Run `python -m unittest tests/test_pdf_text_editor.py` and verify it fails (endpoint not found).
- [x] Step 3: Implement `/api/pdf/inspect-text` endpoint in `app.py`:
  - Receives `file` (PDF) and optional `page` index (default 0).
  - Uses `doc[page].get_text("blocks")` and `get_text("dict")` to extract spans with normalized bounding boxes `[x0, y0, x1, y1]`, text, size, color.
  - Returns JSON with `page_width`, `page_height`, `total_pages`, `page_index`, and `spans`.
- [x] Step 4: Run unit tests and confirm they pass.

---

### Task 2: Backend Text Modification Engine (`/api/pdf/apply-text-edits`)

**Files:**
- Modify: `app.py`
- Test: `tests/test_pdf_text_editor.py`

- [x] Step 1: Write unit tests in `tests/test_pdf_text_editor.py` testing `/api/pdf/apply-text-edits`:
  - Test replacing an existing name in a PDF.
  - Test adding a new text label at coordinates `(100, 200)`.
  - Verify output PDF contains the new text and has old text redacted.
- [x] Step 2: Run unit tests and verify they fail.
- [x] Step 3: Implement `/api/pdf/apply-text-edits` in `app.py`:
  - Accepts `file` (multipart) + `edits` (JSON string containing `{ page: int, original_bbox: [x0, y0, x1, y1], new_text: str, font_size: float, color: [r, g, b] }` and `additions: [...]`).
  - Iterates over edits per page: calls `page.add_redact_annot(rect, fill=(1,1,1))` then `page.apply_redactions()`, then `page.insert_textbox(rect, new_text, fontname="helv", fontsize=font_size, color=color)`.
  - For additions: calls `page.insert_text((x, y), text, fontname="helv", fontsize=font_size, color=color)`.
  - Returns modified PDF stream with `Content-Disposition` header.
- [x] Step 4: Run unit tests and confirm they pass.

---

### Task 3: Interactive Visual Canvas in Frontend (`#editor-view`)

**Files:**
- Modify: `app.py`

- [x] Step 1: Add a "Modo Edição de Texto" toolbar and visual page canvas inside `#editor-view`:
  - Mode switcher: `[📋 Organizar Páginas]` vs `[✏️ Editar Texto do PDF]`.
  - Visual page container with high-DPI page render and relative overlay container.
  - Bounding box spans over detected text with hover highlight and click-to-edit.
  - "＋ Adicionar Texto" button to click on canvas and place a new floating text box.
  - Text formatting controls: Font size selector (10, 12, 14, 18, 24, 32), Color picker/swatches (Preto, Azul, Vermelho, Cinza), and Delete button.
- [x] Step 2: Implement frontend JavaScript logic:
  - `loadPageForTextEditing(pageIndex)`: calls `/api/pdf/inspect-text` and renders text block overlays.
  - `onSpanClick(spanEl)`: transforms span into editable input/contenteditable.
  - `addFloatingTextBox(x, y)`: creates a draggable/resizable new text element.
  - `saveTextEdits()`: bundles modifications into JSON and calls `/api/pdf/apply-text-edits`, downloading the resulting vector PDF.
- [x] Step 3: Style text editing overlay with CSS design tokens (`var(--primary)`, `var(--bg-surface)`, etc.).

---

### Task 4: Home View Tool Card, Navigation & Verification

**Files:**
- Modify: `app.py`
- Modify: `tests/test_web_features.py`

- [x] Step 1: Add "Editar Texto do PDF" tool card in `app.py` `.tools-grid` under "Organizar PDF" category with badge `<span class="tool-badge badge-new">Texto</span>`.
- [x] Step 2: Wire up route `/tool/edit-text-pdf` to open the text editor directly.
- [x] Step 3: Run full test suite: `.venv\Scripts\python.exe -m unittest discover tests`.
- [x] Step 4: Restart local server and verify end-to-end in browser.
