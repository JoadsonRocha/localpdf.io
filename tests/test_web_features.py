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


if __name__ == "__main__":
    unittest.main()
