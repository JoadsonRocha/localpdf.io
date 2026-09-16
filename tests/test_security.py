import io
import unittest

import fitz

from app import allowed_file, app


class SecurityTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.original_max_content_length = app.config["MAX_CONTENT_LENGTH"]

    def tearDown(self):
        app.config["MAX_CONTENT_LENGTH"] = self.original_max_content_length

    def make_pdf(self):
        document = fitz.open()
        page = document.new_page()
        page.insert_text((40, 80), "security test")
        data = document.tobytes()
        document.close()
        return data

    def test_rejects_unsupported_extensions(self):
        self.assertFalse(allowed_file("payload.exe"))
        self.assertFalse(allowed_file("payload.pdf.exe"))

    def test_rejects_missing_files(self):
        response = self.client.post(
            "/convert",
            data={"tool": "compress-pdf"},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)

    def test_rejects_unknown_tool(self):
        response = self.client.post(
            "/convert",
            data={
                "files": (io.BytesIO(b"data"), "file.pdf"),
                "tool": "unknown",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)

    def test_enforces_upload_limit(self):
        app.config["MAX_CONTENT_LENGTH"] = 4
        response = self.client.post(
            "/convert",
            data={
                "files": (io.BytesIO(b"too large"), "file.pdf"),
                "tool": "compress-pdf",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 413)

    def test_rejects_weak_password_without_internal_error(self):
        response = self.client.post(
            "/convert",
            data={
                "files": (io.BytesIO(self.make_pdf()), "..\\..\\safe.pdf"),
                "tool": "protect-pdf",
                "password": "123",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertNotIn(b"Traceback", response.data)
        self.assertNotIn(b"site-packages", response.data)

    def test_editor_rejects_non_pdf_sources(self):
        response = self.client.post(
            "/editor/preview",
            data={"files": (io.BytesIO(b"data"), "file.exe")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)

    def test_editor_rejects_malformed_page_data(self):
        response = self.client.post(
            "/editor/export",
            data={
                "files": (io.BytesIO(self.make_pdf()), "file.pdf"),
                "pages": "{",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)


if __name__ == "__main__":
    unittest.main()
