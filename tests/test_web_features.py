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


if __name__ == "__main__":
    unittest.main()
