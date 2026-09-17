import os
import unittest
from app import app, is_web_environment


class WebEnvironmentTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_environment_detection_flags(self):
        # Test explicit LOCALPDF_MODE
        orig_mode = os.environ.get("LOCALPDF_MODE")
        orig_railway = os.environ.get("RAILWAY_ENVIRONMENT")
        try:
            os.environ["LOCALPDF_MODE"] = "web"
            self.assertTrue(is_web_environment())

            os.environ["LOCALPDF_MODE"] = "local"
            self.assertFalse(is_web_environment())

            del os.environ["LOCALPDF_MODE"]
            os.environ["RAILWAY_ENVIRONMENT"] = "production"
            self.assertTrue(is_web_environment())
        finally:
            if orig_mode is not None:
                os.environ["LOCALPDF_MODE"] = orig_mode
            elif "LOCALPDF_MODE" in os.environ:
                del os.environ["LOCALPDF_MODE"]

            if orig_railway is not None:
                os.environ["RAILWAY_ENVIRONMENT"] = orig_railway
            elif "RAILWAY_ENVIRONMENT" in os.environ:
                del os.environ["RAILWAY_ENVIRONMENT"]

    def test_security_and_cache_headers_on_convert(self):
        import io
        import fitz

        doc = fitz.open()
        p = doc.new_page()
        p.insert_text((50, 50), "Teste headers")
        pdf_bytes = doc.tobytes()
        doc.close()

        resp = self.client.post(
            "/convert",
            data={"tool": "compress-pdf", "files": (io.BytesIO(pdf_bytes), "doc.pdf")},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers.get("X-Content-Type-Options"), "nosniff")
        self.assertEqual(resp.headers.get("X-Frame-Options"), "SAMEORIGIN")
        self.assertIn("no-store", resp.headers.get("Cache-Control", ""))
        self.assertIn("no-cache", resp.headers.get("Pragma", ""))

    def test_about_page_contains_web_vs_desktop_differentiator(self):
        resp = self.client.get("/about")
        self.assertEqual(resp.status_code, 200)
        html = resp.get_data(as_text=True)

        # Verify the comparison section and modalities exist
        self.assertIn("Versão Web (Railway) vs.", html)
        self.assertIn("Aplicativo Desktop (Windows .msi)", html)
        self.assertIn("Versão Web (Nuvem Railway)", html)
        self.assertIn("LocalPDF Web (Railway)", html)
        self.assertIn("Processamento Efêmero", html)
        self.assertIn("Zero Persistência", html)
        self.assertIn("LocalPDF.msi", html)

    def test_web_mode_template_flag(self):
        orig_mode = os.environ.get("LOCALPDF_MODE")
        try:
            os.environ["LOCALPDF_MODE"] = "web"
            resp_web = self.client.get("/")
            html_web = resp_web.get_data(as_text=True)
            self.assertIn("const isServerWebMode = true;", html_web)

            os.environ["LOCALPDF_MODE"] = "local"
            resp_local = self.client.get("/")
            html_local = resp_local.get_data(as_text=True)
            self.assertIn("const isServerWebMode = false;", html_local)
        finally:
            if orig_mode is not None:
                os.environ["LOCALPDF_MODE"] = orig_mode
            elif "LOCALPDF_MODE" in os.environ:
                del os.environ["LOCALPDF_MODE"]


if __name__ == "__main__":
    unittest.main()
