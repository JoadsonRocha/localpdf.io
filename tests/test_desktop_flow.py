import base64
import io
import os
import pathlib
import tempfile
import unittest
from unittest.mock import MagicMock

from app import app
from windows_launcher import DesktopAPI


class DesktopFlowTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.orig_mode = os.environ.get("LOCALPDF_MODE")

    def tearDown(self):
        if self.orig_mode is not None:
            os.environ["LOCALPDF_MODE"] = self.orig_mode
        elif "LOCALPDF_MODE" in os.environ:
            del os.environ["LOCALPDF_MODE"]

    def test_desktop_endpoints_blocked_in_web_mode(self):
        os.environ["LOCALPDF_MODE"] = "web"

        res = self.client.post("/api/desktop/save")
        self.assertEqual(res.status_code, 403)

        res = self.client.post("/api/desktop/write-file")
        self.assertEqual(res.status_code, 403)

        res = self.client.post("/api/desktop/open-folder")
        self.assertEqual(res.status_code, 403)

        res = self.client.post("/api/desktop/open-file")
        self.assertEqual(res.status_code, 403)

    def test_desktop_save_to_downloads_in_local_mode(self):
        os.environ["LOCALPDF_MODE"] = "local"
        test_content = b"%PDF-1.4 test document content"

        res = self.client.post(
            "/api/desktop/save",
            data={"file": (io.BytesIO(test_content), "test_doc.pdf")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertTrue(data["success"])
        saved_path = data["path"]
        self.assertTrue(os.path.exists(saved_path))
        self.assertTrue(saved_path.endswith(".pdf"))

        # Clean up created test file
        try:
            os.remove(saved_path)
        except OSError:
            pass

    def test_desktop_save_rejects_disallowed_extension(self):
        os.environ["LOCALPDF_MODE"] = "local"
        res = self.client.post(
            "/api/desktop/save",
            data={"file": (io.BytesIO(b"binary"), "malicious.exe")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 400)
        self.assertIn("Extensão não permitida", res.get_json()["error"])

    def test_desktop_write_to_path_success_and_protection(self):
        os.environ["LOCALPDF_MODE"] = "local"
        with tempfile.TemporaryDirectory() as tmpdir:
            target_path = os.path.join(tmpdir, "chosen_output.pdf")
            test_content = b"%PDF-1.4 custom path content"

            res = self.client.post(
                "/api/desktop/write-file",
                data={
                    "file": (io.BytesIO(test_content), "document.pdf"),
                    "path": target_path,
                },
                content_type="multipart/form-data",
            )
            self.assertEqual(res.status_code, 200)
            self.assertTrue(os.path.exists(target_path))
            with open(target_path, "rb") as f:
                self.assertEqual(f.read(), test_content)

            # Test rejecting system directories
            windir = os.environ.get("WINDIR", "C:\\Windows")
            sys_path = os.path.join(windir, "bad.pdf")
            res_bad = self.client.post(
                "/api/desktop/write-file",
                data={
                    "file": (io.BytesIO(test_content), "document.pdf"),
                    "path": sys_path,
                },
                content_type="multipart/form-data",
            )
            self.assertEqual(res_bad.status_code, 400)

    def test_desktop_api_class(self):
        api = DesktopAPI()
        downloads_dir = api.get_downloads_dir()
        self.assertTrue(os.path.isdir(downloads_dir))

        # Test choose_save_path with mock window
        mock_win = MagicMock()
        mock_win.create_file_dialog.return_value = ["C:\\test\\chosen.pdf"]
        api.set_window(mock_win)

        chosen = api.choose_save_path("meu_arquivo.pdf")
        self.assertEqual(chosen, os.path.abspath("C:\\test\\chosen.pdf"))

        # Test cancel dialog
        mock_win.create_file_dialog.return_value = None
        self.assertIsNone(api.choose_save_path("meu_arquivo.pdf"))

        # Test save_file with base64 data
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "base64_saved.pdf")
            raw_data = b"%PDF-1.4 base64 direct test"
            b64_str = base64.b64encode(raw_data).decode("ascii")

            res = api.save_file("doc.pdf", b64_str, target_path=out_file)
            self.assertTrue(res["success"])
            self.assertTrue(os.path.exists(out_file))
            with open(out_file, "rb") as f:
                self.assertEqual(f.read(), raw_data)


if __name__ == "__main__":
    unittest.main()
