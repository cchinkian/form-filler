from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from reportlab.pdfgen import canvas
import pypdf

import config_loader
import pdf_engine


def main():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        forms = root / "forms"
        output = root / "filled"
        form_folder = forms / "form_a"
        form_folder.mkdir(parents=True)

        pdf_path = form_folder / "Any Clean Name.pdf"
        c = canvas.Canvas(str(pdf_path), pagesize=(595, 842))
        c.drawString(50, 800, "Sample Form A")
        c.drawString(50, 760, "Name:")
        c.drawString(50, 730, "IC:")
        c.drawString(50, 700, "CIF:")
        c.drawString(50, 670, "Amount:")
        c.drawString(50, 640, "Date:")
        c.save()

        settings = {
            "forms_folder": str(forms),
            "output_folder": str(output),
            "default_font_size": 10,
        }
        forms_config = {
            "_shared_fields": {},
            "form_a": {
                "name": "Form A",
                "template_subfolder": "form_a",
                "fields": [
                    {"name": "name", "source": "data", "page": 1, "x": 100, "y": 760, "required": True},
                    {"name": "ic_number", "source": "data", "page": 1, "x": 100, "y": 730, "required": True},
                    {"name": "cif_no", "source": "data", "page": 1, "x": 100, "y": 700},
                    {"name": "amount", "source": "data", "page": 1, "x": 100, "y": 670, "required": True},
                    {"name": "date", "source": "session", "session_key": "date", "page": 1, "x": 100, "y": 640},
                ],
            },
        }
        app = {"id": "form_a", "name": "Form A", "forms": ["form_a"]}
        client = {
            "name": "Test Client",
            "ic_number": "900101101234",
            "cif_no": "CIF123",
            "amount": "10000",
        }

        paths, warnings = pdf_engine.fill_bundle(
            app,
            forms_config,
            client,
            output,
            settings,
            config_loader.find_template,
            session={"date": "19/05/2026", "rm_branch": "KJG"},
        )

        assert len(paths) == 1
        assert paths[0].exists()
        assert warnings == []

        text = pypdf.PdfReader(str(paths[0])).pages[0].extract_text() or ""
        assert "Test Client" in text
        assert "10000" in text
        assert "19/05/2026" in text
        print(f"OK: {paths[0]}")


if __name__ == "__main__":
    main()
