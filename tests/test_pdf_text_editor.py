import io
import json
import unittest
import fitz

from app import app


class PdfTextEditorTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

        # Create a sample PDF with known text
        doc = fitz.open()
        page = doc.new_page(width=600, height=800)
        page.insert_text((50, 100), "Titulo do Documento", fontsize=18, color=(0, 0, 0))
        page.insert_text((50, 150), "Nome do Cliente: Joao da Silva", fontsize=12, color=(0.2, 0.2, 0.2))
        self.sample_pdf_bytes = doc.tobytes()
        doc.close()

    def test_inspect_text_endpoint(self):
        data = {
            "file": (io.BytesIO(self.sample_pdf_bytes), "sample.pdf"),
            "page": 0,
        }
        res = self.client.post(
            "/api/pdf/inspect-text",
            data=data,
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 200)
        json_data = res.get_json()
        self.assertTrue(json_data["success"])
        self.assertEqual(json_data["total_pages"], 1)
        self.assertEqual(json_data["page_index"], 0)
        self.assertIn("page_width", json_data)
        self.assertIn("page_height", json_data)
        self.assertIn("page_image", json_data)

        # Check that spans contain our text
        spans = json_data["spans"]
        self.assertTrue(len(spans) >= 2)
        texts = [s["text"].strip() for s in spans]
        self.assertTrue(any("Titulo do Documento" in t for t in texts))
        self.assertTrue(any("Joao da Silva" in t for t in texts))

        # Check span properties
        for s in spans:
            self.assertIn("bbox", s)
            self.assertEqual(len(s["bbox"]), 4)
            self.assertIn("size", s)
            self.assertIn("color", s)

    def test_apply_text_edits_endpoint(self):
        # First inspect to get Joao da Silva bbox
        inspect_res = self.client.post(
            "/api/pdf/inspect-text",
            data={"file": (io.BytesIO(self.sample_pdf_bytes), "sample.pdf"), "page": 0},
            content_type="multipart/form-data",
        )
        self.assertEqual(inspect_res.status_code, 200)
        spans = inspect_res.get_json()["spans"]
        joao_span = [s for s in spans if "Joao da Silva" in s["text"]][0]

        edits = [
            {
                "page": 0,
                "original_bbox": joao_span["bbox"],
                "new_text": "Nome do Cliente: Carlos Eduardo Pereira",
                "font_size": 12.0,
                "color": [0, 0, 0],
            }
        ]
        additions = [
            {
                "page": 0,
                "x": 50,
                "y": 250,
                "text": "Observacao: Documento aprovado sem ressalvas.",
                "font_size": 11.0,
                "color": [0.1, 0.5, 0.1],
            }
        ]

        res = self.client.post(
            "/api/pdf/apply-text-edits",
            data={
                "file": (io.BytesIO(self.sample_pdf_bytes), "meu_contrato.pdf"),
                "edits": json.dumps(edits),
                "additions": json.dumps(additions),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.mimetype, "application/pdf")
        self.assertIn("meu_contrato_editado.pdf", res.headers.get("Content-Disposition", ""))

        # Verify edited text in output PDF
        out_doc = fitz.open(stream=res.data, filetype="pdf")
        page_text = out_doc[0].get_text()
        out_doc.close()

        self.assertIn("Carlos Eduardo Pereira", page_text)
        self.assertIn("Documento aprovado sem ressalvas", page_text)
        self.assertNotIn("Joao da Silva", page_text)


if __name__ == "__main__":
    unittest.main()
