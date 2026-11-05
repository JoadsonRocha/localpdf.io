# Filee Engine Architecture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Refactor LocalPDF.io's monolithic conversion logic into a modular, extensible engine inspired by Filee's pipeline (`BaseConverter`, `ConversionRouter`, `EngineManager`) while preserving 100% backward compatibility with all web and desktop features.

**Architecture:** A core contract layer (`core/converter.py`), a dependency manager (`core/engine_manager.py`), and a hybrid route planner (`core/router.py`) orchestrating modular converters (`converters/documents.py`, `converters/images.py`, `converters/pdf_ops.py`, `converters/ocr.py`, `converters/pdfa.py`) and a dedicated editor service (`services/editor.py`), wired cleanly into Flask (`web/routes.py` and `app.py`).

**Tech Stack:** Python 3.11+, Flask, PyMuPDF, pdf2docx, ReportLab, openpyxl, pdfplumber, pytesseract, Ghostscript, unittest.

**Spec:** [docs/superpowers/specs/2026-11-05-filee-engine-architecture-design.md](file:///d:/FULLSTARK/localpdf.io/docs/superpowers/specs/2026-11-05-filee-engine-architecture-design.md)

## Global Constraints

- Must run on Python 3.11+ without adding new mandatory external dependencies.
- Must preserve 100% offline privacy (no cloud calls, no telemetry).
- Must retain all existing HTTP endpoint routes (`/convert`, `/editor`, `/editor/preview`, `/editor/export`, etc.).
- Output file and zip naming conventions (e.g. `_comprimido.pdf`, `_mesclado.pdf`) must remain strictly identical to preserve test assertions in `tests/test_web_features.py`.
- `windows_launcher.py` must remain intact and functional.

## Review Focus

- Route collision on identical input/output formats (`pdf -> pdf`): Dispatched explicitly by `tool_id` rather than blind format lookup.
- OCR missing Tesseract: Must silently fall back to native digital text extraction via PyMuPDF without raising 500 error.
- Multi-step route planning: A composite conversion (e.g. `docx -> pdf -> png`) must clean up intermediate temporary files and return only final outputs.
- Corrupted or password-protected input handling: Must return structured `ConversionResult(success=False, error=...)` instead of unhandled exceptions crashing the worker.
- Desktop download permission errors in test sandboxes: Desktop download routes must validate writeability or catch OS permission errors gracefully.

---

### Task 1: Core Converter Contracts (`core/converter.py`)

**Files:**
- Create: `core/__init__.py`
- Create: `core/converter.py`
- Test: `tests/test_core_converter.py`

**Interfaces:**
- Produces: `ConversionStep`, `ConversionResult`, `BaseConverter`

- [ ] **Step 1: Write the failing test**

Create `tests/test_core_converter.py` testing `ConversionStep`, `ConversionResult`, and an implementation of `BaseConverter` verifying abstract methods and execution contract.

```python
import unittest
from core.converter import ConversionStep, ConversionResult, BaseConverter

class DummyConverter(BaseConverter):
    id = "dummy"
    display_name = "Dummy Converter"
    edges = [("txt", "pdf")]

    def is_available(self) -> bool:
        return True

    def convert(self, step: ConversionStep) -> ConversionResult:
        return ConversionResult(success=True, output_files=["dummy.pdf"], suggested_filename="dummy.pdf")

class TestCoreConverter(unittest.TestCase):
    def test_dummy_converter_execution(self):
        conv = DummyConverter()
        self.assertTrue(conv.is_available())
        step = ConversionStep(input_paths=["test.txt"], output_dir="/tmp", target_format="pdf")
        res = conv.convert(step)
        self.assertTrue(res.success)
        self.assertEqual(res.output_files, ["dummy.pdf"])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_core_converter.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'core'`

- [ ] **Step 3: Implement `core/converter.py`**

Implement `ConversionStep`, `ConversionResult`, and `BaseConverter` in `core/converter.py` and empty `core/__init__.py`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_core_converter.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add core/__init__.py core/converter.py tests/test_core_converter.py
git commit -m "feat(core): add BaseConverter, ConversionStep, and ConversionResult contracts"
```

---

### Task 2: Native Engine Manager (`core/engine_manager.py`)

**Files:**
- Create: `core/engine_manager.py`
- Test: `tests/test_engine_manager.py`

**Interfaces:**
- Consumes: Nothing
- Produces: `EngineManager`, `EngineStatus` enum, `get_engine_manager()` singleton

- [ ] **Step 1: Write the failing test**

Create `tests/test_engine_manager.py` testing detection of Tesseract and Ghostscript, status reporting, and environment PATH configuration without crashing on non-Windows/missing paths.

```python
import unittest
from core.engine_manager import EngineManager, EngineStatus

class TestEngineManager(unittest.TestCase):
    def test_engine_manager_status(self):
        manager = EngineManager()
        status = manager.get_status()
        self.assertIn("tesseract", status)
        self.assertIn("ghostscript", status)
        self.assertIn(status["tesseract"]["status"], [EngineStatus.READY, EngineStatus.MISSING])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_engine_manager.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'core.engine_manager'`

- [ ] **Step 3: Implement `core/engine_manager.py`**

Implement `EngineManager` isolating detection logic from `app.py` lines 39-144, supporting system paths and `vendor/` directories, with `configure_environment()` and `get_status()`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_engine_manager.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add core/engine_manager.py tests/test_engine_manager.py
git commit -m "feat(core): add EngineManager for native Ghostscript and Tesseract detection"
```

---

### Task 3: Hybrid Conversion Router (`core/router.py`)

**Files:**
- Create: `core/router.py`
- Test: `tests/test_router.py`

**Interfaces:**
- Consumes: `core.converter.BaseConverter`, `core.converter.ConversionStep`, `core.converter.ConversionResult`
- Produces: `ConversionRouter` with `register(converter)`, `get_converter(tool_id)`, `convert(...)`, `find_route(from_ext, to_ext)`

- [ ] **Step 1: Write the failing test**

Create `tests/test_router.py` testing:
1. Registration of converters.
2. Lookup and dispatch by `tool_id` (disambiguating `pdf -> pdf` tools).
3. Lookup and dispatch by direct format `from_ext -> to_ext`.
4. Multi-hop BFS route discovery (e.g. `txt -> pdf -> png`).
5. Proper error result when route is unsupported.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_router.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'core.router'`

- [ ] **Step 3: Implement `core/router.py`**

Implement `ConversionRouter` with converter registry by `id` and format graph dictionary. Implement BFS route planner (max 2 hops) for composite routes, executing intermediate steps into temporary directories and returning the final `ConversionResult`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_router.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add core/router.py tests/test_router.py
git commit -m "feat(core): implement ConversionRouter with hybrid tool_id lookup and format graph planner"
```

---

### Task 4: Document Conversores (`converters/documents.py`)

**Files:**
- Create: `converters/__init__.py`
- Create: `converters/documents.py`
- Test: `tests/test_converters_documents.py`

**Interfaces:**
- Consumes: `core.converter.BaseConverter`, `core.converter.ConversionStep`, `core.converter.ConversionResult`
- Produces:
  - `WordToPdfConverter` (`id="word-to-pdf"`, edges `[("docx", "pdf")]`)
  - `ExcelToPdfConverter` (`id="excel-to-pdf"`, edges `[("xlsx", "pdf")]`)
  - `TxtToPdfConverter` (`id="txt-to-pdf"`, edges `[("txt", "pdf")]`)
  - `PdfToWordConverter` (`id="pdf-to-word"`, edges `[("pdf", "docx")]`)
  - `PdfToExcelConverter` (`id="pdf-to-excel"`, edges `[("pdf", "xlsx")]`)

- [ ] **Step 1: Write the failing test**

Create `tests/test_converters_documents.py` testing text to PDF, Excel to PDF, and Word to PDF converters using lightweight synthetic sample files.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_documents.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'converters.documents'`

- [ ] **Step 3: Implement `converters/documents.py`**

Extract `txt_to_pdf`, `excel_to_pdf`, `word_to_pdf`, `pdf_to_word`, and `pdf_to_excel` from `app.py` into individual `BaseConverter` classes with appropriate `suggested_filename` outputs.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_documents.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add converters/__init__.py converters/documents.py tests/test_converters_documents.py
git commit -m "feat(converters): add document converters for Word, Excel, and Text"
```

---

### Task 5: Image Conversores (`converters/images.py`)

**Files:**
- Create: `converters/images.py`
- Test: `tests/test_converters_images.py`

**Interfaces:**
- Consumes: `core.converter.BaseConverter`, `core.converter.ConversionStep`, `core.converter.ConversionResult`
- Produces:
  - `ImagesToPdfConverter` (`id="images-to-pdf"`, edges `[("png", "pdf"), ("jpg", "pdf"), ("jpeg", "pdf")]`)
  - `PdfToImagesConverter` (`id="pdf-to-images"`, edges `[("pdf", "png"), ("pdf", "jpg")]`)

- [ ] **Step 1: Write the failing test**

Create `tests/test_converters_images.py` testing conversion of a test PNG image to PDF, and a PDF back to PNG images.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_images.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'converters.images'`

- [ ] **Step 3: Implement `converters/images.py`**

Extract `images_to_pdf` and `pdf_to_images` from `app.py` into `ImagesToPdfConverter` and `PdfToImagesConverter`, preserving image DPI and resolution rendering.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_images.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add converters/images.py tests/test_converters_images.py
git commit -m "feat(converters): add image converters for Images to PDF and PDF to Images"
```

---

### Task 6: PDF Operations (`converters/pdf_ops.py`)

**Files:**
- Create: `converters/pdf_ops.py`
- Test: `tests/test_converters_pdf_ops.py`

**Interfaces:**
- Consumes: `core.converter.BaseConverter`, `core.converter.ConversionStep`, `core.converter.ConversionResult`
- Produces:
  - `MergePdfConverter` (`id="merge-pdf"`)
  - `SplitPdfConverter` (`id="split-pdf"`)
  - `CompressPdfConverter` (`id="compress-pdf"`)
  - `ProtectPdfConverter` (`id="protect-pdf"`)
  - `UnlockPdfConverter` (`id="unlock-pdf"`)
  - `WatermarkPdfConverter` (`id="watermark-pdf"`)
  - `PageNumbersConverter` (`id="page-numbers-pdf"`)

- [ ] **Step 1: Write the failing test**

Create `tests/test_converters_pdf_ops.py` validating PDF merge, split with page range, compression, password protection/unlock, and watermarking using synthetic PDF documents.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_pdf_ops.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'converters.pdf_ops'`

- [ ] **Step 3: Implement `converters/pdf_ops.py`**

Extract the operations from `app.py` lines 4321-4698 into dedicated `BaseConverter` classes, ensuring each sets its canonical `suggested_filename` (e.g. `{stem}_mesclado.pdf`, `{stem}_comprimido.pdf`, `{stem}_protegido.pdf`).

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_pdf_ops.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add converters/pdf_ops.py tests/test_converters_pdf_ops.py
git commit -m "feat(converters): add PDF operation converters (merge, split, compress, protect, watermark)"
```

---

### Task 7: OCR and PDF/A Conversores (`converters/ocr.py`, `converters/pdfa.py`)

**Files:**
- Create: `converters/ocr.py`
- Create: `converters/pdfa.py`
- Test: `tests/test_converters_ocr_pdfa.py`

**Interfaces:**
- Consumes: `core.converter.BaseConverter`, `core.engine_manager.EngineManager`
- Produces:
  - `OcrPdfConverter` (`id="ocr-pdf"`, with native digital text fallback)
  - `PdfToTextConverter` (`id="pdf-to-text"`, edges `[("pdf", "txt")]`)
  - `PdfToPdfAConverter` (`id="pdf-to-pdfa"`, Ghostscript-backed)

- [ ] **Step 1: Write the failing test**

Create `tests/test_converters_ocr_pdfa.py` testing:
1. `PdfToTextConverter` extraction from digital PDF.
2. `OcrPdfConverter` graceful fallback to digital text when Tesseract is unavailable.
3. `PdfToPdfAConverter` status check and invocation.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_ocr_pdfa.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'converters.ocr'`

- [ ] **Step 3: Implement `converters/ocr.py` and `converters/pdfa.py`**

Extract `ocr_pdf`, `pdf_to_text`, and `pdf_to_pdfa` from `app.py`. Ensure `OcrPdfConverter` preserves the full fallback message and behavior when Tesseract is absent.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_converters_ocr_pdfa.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add converters/ocr.py converters/pdfa.py tests/test_converters_ocr_pdfa.py
git commit -m "feat(converters): add OCR with smart fallback, Text extraction, and PDF/A converter"
```

---

### Task 8: Dedicated Editor Service (`services/editor.py`)

**Files:**
- Create: `services/__init__.py`
- Create: `services/editor.py`
- Test: `tests/test_services_editor.py`

**Interfaces:**
- Consumes: Nothing
- Produces: `render_pdf_page`, `editor_preview_for_file`, `insert_editor_image`, `export_editor_pages`

- [ ] **Step 1: Write the failing test**

Create `tests/test_services_editor.py` testing page rendering to thumbnail base64 and exporting pages with reordering.

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m unittest tests/test_services_editor.py`
Expected: FAIL with `ModuleNotFoundError: No module named 'services.editor'`

- [ ] **Step 3: Implement `services/editor.py`**

Extract editor functions from `app.py` lines 3663-3860 into `services/editor.py`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m unittest tests/test_services_editor.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add services/__init__.py services/editor.py tests/test_services_editor.py
git commit -m "feat(services): isolate PDF editor preview and export logic into services/editor.py"
```

---

### Task 9: Web Routes & App Refactoring (`web/routes.py` & `app.py`)

**Files:**
- Create: `web/__init__.py`
- Create: `web/routes.py`
- Modify: `app.py`
- Test: Run entire test suite `tests/`

**Interfaces:**
- Consumes: `core.router.ConversionRouter`, `core.engine_manager.EngineManager`, `converters.*`, `services.editor`
- Produces: Default instantiated `router`, Flask blueprint/route handlers, lean `app.py`

- [ ] **Step 1: Implement `web/routes.py` and factory router loader**

Implement `web/routes.py` providing a function to register all default converters into a `ConversionRouter` and route endpoints (`/convert`, `/editor/preview`, `/editor/export`, `/api/engines/status`).

- [ ] **Step 2: Refactor `app.py` to use `web/routes.py`**

Replace monolithic function bodies in `app.py` by delegating to the router and editor services, keeping all route decorators, SEO metadata, templates, and desktop endpoints intact.

- [ ] **Step 3: Run the full test suite**

Run: `.venv\Scripts\python.exe -m unittest discover tests`
Expected: All existing test cases pass (including `test_web_features.py`, `test_security.py`, `test_web_environment.py`).

- [ ] **Step 4: Commit**

```bash
git add web/__init__.py web/routes.py app.py
git commit -m "refactor(app): wire Flask endpoints to ConversionRouter and modular converters"
```

