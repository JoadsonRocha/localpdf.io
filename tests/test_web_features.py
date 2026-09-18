import os
import tempfile
import unittest
from app import app


class WebFeaturesTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_home_page_contains_interactive_tabs(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)

        # Verify category tabs
        self.assertIn('class="category-tab active"', html)
        self.assertIn("Organizar PDF", html)
        self.assertIn("Converter PDF", html)
        self.assertIn("Otimizar PDF", html)
        self.assertIn("OCR", html)
        self.assertIn("filterCategory('organizar'", html)

    def test_home_page_tool_cards_have_links_and_new_tab(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)

        # Check links with target="_blank"
        self.assertIn('href="/tool/compress-pdf"', html)
        self.assertIn('href="/tool/merge-pdf"', html)
        self.assertIn('href="/tool/pdf-to-images"', html)
        self.assertIn('href="/editor"', html)
        self.assertIn('target="_blank"', html)

        # Check data-category attributes for filtering
        self.assertIn('data-category="converter"', html)
        self.assertIn('data-category="organizar"', html)
        self.assertIn('data-category="otimizar"', html)
        self.assertIn('data-category="ocr"', html)

        # Check SVG icons are present
        self.assertIn('class="tool-icon"', html)
        self.assertIn("<svg", html)

    def test_tool_direct_routes(self):
        for tool in ("compress-pdf", "merge-pdf", "pdf-to-images", "ocr-pdf"):
            response = self.client.get(f"/tool/{tool}")
            self.assertEqual(response.status_code, 200)
            html = response.get_data(as_text=True)
            self.assertIn(f'const initialToolFromRoute = "{tool}";', html)

    def test_editor_direct_route(self):
        response = self.client.get("/editor")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('const initialToolFromRoute = "edit-pdf";', html)

    def test_favicon_route(self):
        response = self.client.get("/favicon.svg")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "image/svg+xml")

    def test_progress_box_and_file_flow_markup(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('class="progress-box hidden"', html)
        self.assertIn('id="progress-message"', html)
        self.assertIn('id="progress-timer"', html)
        self.assertIn('id="file-list"', html)
        self.assertIn('class="file-remove-btn"', html)

    def test_standardized_renaming_on_convert(self):
        import io
        import fitz
        from PIL import Image

        # Criar PDF de teste
        doc = fitz.open()
        p = doc.new_page()
        p.insert_text((50, 50), "Teste de renomeio")
        pdf_bytes = doc.tobytes()
        doc.close()

        # Teste Compressão
        resp = self.client.post(
            "/convert",
            data={
                "tool": "compress-pdf",
                "files": (io.BytesIO(pdf_bytes), "meu_relatorio.pdf"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 200)
        disp = resp.headers.get("Content-Disposition", "")
        self.assertIn("meu_relatorio_comprimido.pdf", disp)

        # Teste Mesclagem
        resp_merge = self.client.post(
            "/convert",
            data={
                "tool": "merge-pdf",
                "files": [
                    (io.BytesIO(pdf_bytes), "documento_a.pdf"),
                    (io.BytesIO(pdf_bytes), "documento_b.pdf"),
                ],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_merge.status_code, 200)
        disp_merge = resp_merge.headers.get("Content-Disposition", "")
        self.assertIn("documento_a_mesclado.pdf", disp_merge)

        # Teste Imagens para PDF
        img = Image.new("RGB", (100, 100), color="blue")
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="PNG")
        img_bytes.seek(0)

        resp_img = self.client.post(
            "/convert",
            data={
                "tool": "images-to-pdf",
                "files": [(img_bytes, "foto_viagem.png")],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_img.status_code, 200)
        disp_img = resp_img.headers.get("Content-Disposition", "")
        self.assertIn("foto_viagem_convertido.pdf", disp_img)

    def test_convert_error_returns_json_message(self):
        import io
        resp = self.client.post(
            "/convert",
            data={
                "tool": "protect-pdf",
                "password": "12",  # Senha curta demais (< 4)
                "files": (io.BytesIO(b"%PDF-1.4 dummy"), "doc.pdf"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 400)
        data = resp.get_json()
        self.assertIsNotNone(data)
        self.assertIn("error", data)

    def test_about_routes(self):
        for route in ("/about", "/sobre"):
            resp = self.client.get(route)
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Sobre o LocalPDF.io", html)
            self.assertIn("id=\"about-view\"", html)
            self.assertIn("initialToolFromRoute = \"about\"", html)

    def test_new_tools_conversion(self):
        import io
        import fitz

        # Criar PDF de teste com 3 páginas
        doc = fitz.open()
        for i in range(3):
            p = doc.new_page()
            p.insert_text((50, 50), f"Pagina {i + 1}")
            p.insert_text((50, 100), "Col1   Col2   Col3")
            p.insert_text((50, 120), "Val1   Val2   Val3")
        pdf_bytes = doc.tobytes()
        doc.close()

        # 1. PDF para JPG
        resp_jpg = self.client.post(
            "/convert",
            data={"tool": "pdf-to-jpg", "files": [(io.BytesIO(pdf_bytes), "relatorio.pdf")]},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_jpg.status_code, 200)
        self.assertIn("relatorio_jpg.zip", resp_jpg.headers.get("Content-Disposition", ""))

        # 2. PDF para PNG
        resp_png = self.client.post(
            "/convert",
            data={"tool": "pdf-to-png", "files": [(io.BytesIO(pdf_bytes), "documento.pdf")]},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_png.status_code, 200)
        self.assertIn("documento_png.zip", resp_png.headers.get("Content-Disposition", ""))

        # 3. PDF para Excel
        resp_excel = self.client.post(
            "/convert",
            data={"tool": "pdf-to-excel", "files": [(io.BytesIO(pdf_bytes), "financeiro.pdf")]},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_excel.status_code, 200)
        self.assertIn("financeiro.xlsx", resp_excel.headers.get("Content-Disposition", ""))

        # 4. Dividir por intervalo de páginas (apenas página 2)
        resp_split = self.client.post(
            "/convert",
            data={
                "tool": "split-pdf",
                "page_range": "2",
                "files": [(io.BytesIO(pdf_bytes), "livro.pdf")],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_split.status_code, 200)
        self.assertIn("livro_extraido.pdf", resp_split.headers.get("Content-Disposition", ""))

        # 5. Proteger e Desbloquear PDF
        resp_prot = self.client.post(
            "/convert",
            data={
                "tool": "protect-pdf",
                "password": "senhaSegura123",
                "files": [(io.BytesIO(pdf_bytes), "contrato.pdf")],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_prot.status_code, 200)
        protected_bytes = resp_prot.data

        resp_unlock = self.client.post(
            "/convert",
            data={
                "tool": "unlock-pdf",
                "password": "senhaSegura123",
                "files": [(io.BytesIO(protected_bytes), "contrato_protegido.pdf")],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_unlock.status_code, 200)
        self.assertIn("contrato_protegido_desbloqueado.pdf", resp_unlock.headers.get("Content-Disposition", ""))

        # 6. Marca d'água profissional com rotação e opacidade
        resp_wm = self.client.post(
            "/convert",
            data={
                "tool": "watermark-pdf",
                "watermark_text": "CONFIDENCIAL",
                "watermark_position": "diagonal",
                "watermark_color": "gray",
                "watermark_opacity": "0.22",
                "files": [(io.BytesIO(pdf_bytes), "minuta.pdf")],
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp_wm.status_code, 200)
        self.assertIn("minuta_marca_dagua.pdf", resp_wm.headers.get("Content-Disposition", ""))

    def test_homepage_javascript_and_cards_visibility(self):
        """Ensures that the homepage HTML has visible cards by default and valid JavaScript."""
        import re
        import subprocess

        resp = self.client.get("/")
        self.assertEqual(resp.status_code, 200)
        html = resp.get_data(as_text=True)

        # Card visibility default check
        self.assertIn(".tools-grid .tool-card { opacity: 1;", html)

        # Extract <script> content and test syntax via node if available
        scripts = re.findall(
            r'<script(?![^>]*type=["\']application/ld\+json["\'])[^>]*>(.*?)</script>',
            html,
            re.DOTALL,
        )
        self.assertTrue(len(scripts) > 0, "Nenhum script encontrado na página principal")

        for script in scripts:
            if not script.strip():
                continue
            with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
                tf.write(script)
                tf_path = tf.name
            try:
                # If node is available on system, check syntax
                node_proc = subprocess.run(
                    ["node", "--check", tf_path],
                    capture_output=True,
                    text=True,
                )
                self.assertEqual(
                    node_proc.returncode,
                    0,
                    f"Erro de sintaxe JavaScript detectado:\n{node_proc.stderr}",
                )
            except FileNotFoundError:
                # Node not installed in testing environment, fallback to basic syntax pattern check
                self.assertNotIn(r"replace(/\/g", script)
            finally:
                try:
                    os.remove(tf_path)
                except OSError:
                    pass


if __name__ == "__main__":
    unittest.main()

