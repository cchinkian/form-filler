import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from reportlab.pdfgen import canvas
import pypdf

import package_engine


def _make_pdf(path: Path, title: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=(595, 842))
    c.drawString(50, 800, title)
    c.drawString(50, 760, "Name:")
    c.drawString(50, 730, "IC:")
    c.drawString(50, 700, "Amount:")
    c.drawString(50, 670, "Product:")
    c.drawString(50, 640, "Date:")
    c.save()


def _write_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def test_package_engine_generates_ordered_combined_pdf() -> None:
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        forms_root = root / "forms"
        output_root = root / "output"
        fixture_root = root / "json"

        _make_pdf(forms_root / "first_source.pdf", "FIRST_SOURCE_TEMPLATE")
        _make_pdf(forms_root / "second_source.pdf", "SECOND_SOURCE_TEMPLATE")
        fixture_root.mkdir()

        _write_json(
            fixture_root / "procedures.json",
            [
                {
                    "ProcedureCode": "PROC_SMOKE",
                    "DisplayName": "Two Form Smoke",
                    "Active": True,
                }
            ],
        )
        _write_json(
            fixture_root / "source_forms.json",
            [
                {
                    "SourceFormCode": "SF_FIRST",
                    "DisplayName": "First Source",
                    "PDFFilePath": "first_source.pdf",
                    "MappingKey": "first_mapping",
                    "Active": True,
                },
                {
                    "SourceFormCode": "SF_SECOND",
                    "DisplayName": "Second Source",
                    "PDFFilePath": "second_source.pdf",
                    "MappingKey": "second_mapping",
                    "Active": True,
                },
            ],
        )
        _write_json(
            fixture_root / "procedure_items.json",
            [
                {
                    "ProcedureCode": "PROC_SMOKE",
                    "StepNo": 3,
                    "ItemType": "SourceForm",
                    "SourceFormCode": "SF_SECOND",
                    "BlankPageCount": 0,
                },
                {
                    "ProcedureCode": "PROC_SMOKE",
                    "StepNo": 1,
                    "ItemType": "SourceForm",
                    "SourceFormCode": "SF_FIRST",
                    "BlankPageCount": 0,
                },
                {
                    "ProcedureCode": "PROC_SMOKE",
                    "StepNo": 2,
                    "ItemType": "BlankPage",
                    "SourceFormCode": "",
                    "BlankPageCount": 1,
                },
            ],
        )
        _write_json(
            fixture_root / "forms.json",
            {
                "_shared_fields": {},
                "first_mapping": {
                    "name": "First Mapping",
                    "fields": [
                        {
                            "name": "name",
                            "ExcelColumnOrField": "name",
                            "source": "data",
                            "page": 1,
                            "x": 100,
                            "y": 760,
                            "required": True,
                        },
                        {
                            "name": "amount",
                            "ExcelColumnOrField": "amount",
                            "source": "data",
                            "page": 1,
                            "x": 100,
                            "y": 700,
                            "required": True,
                        },
                        {
                            "name": "date",
                            "source": "session",
                            "session_key": "date",
                            "page": 1,
                            "x": 100,
                            "y": 640,
                        },
                    ],
                },
                "second_mapping": {
                    "name": "Second Mapping",
                    "fields": [
                        {
                            "name": "ic_number",
                            "ExcelColumnOrField": "ic_number",
                            "source": "data",
                            "page": 1,
                            "x": 100,
                            "y": 730,
                            "required": True,
                        },
                        {
                            "name": "product_type",
                            "ExcelColumnOrField": "product_type",
                            "source": "data",
                            "page": 1,
                            "x": 100,
                            "y": 670,
                        },
                        {
                            "name": "date",
                            "source": "session",
                            "session_key": "date",
                            "page": 1,
                            "x": 100,
                            "y": 640,
                        },
                    ],
                },
            },
        )

        procedure = json.loads((fixture_root / "procedures.json").read_text(encoding="utf-8"))[0]
        source_forms = {
            row["SourceFormCode"]: row
            for row in json.loads((fixture_root / "source_forms.json").read_text(encoding="utf-8"))
        }
        procedure_items = json.loads((fixture_root / "procedure_items.json").read_text(encoding="utf-8"))
        forms_config = json.loads((fixture_root / "forms.json").read_text(encoding="utf-8"))

        customer_row = {
            "name": "Avery Tan",
            "ic_number": "900101101234",
            "cis": "CIS-987654",
            "amount": "25000",
            "product_type": "UT Subscription",
        }

        result = package_engine.generate_package(
            procedure=procedure,
            source_forms=source_forms,
            procedure_items=procedure_items,
            forms_config=forms_config,
            client=customer_row,
            settings={
                "forms_folder": str(forms_root),
                "output_folder": str(output_root),
                "default_font_size": 10,
            },
            output_root=output_root,
            session={"date": "19/05/2026", "rm_branch": "KJG"},
        )

        output_path = result["output_path"]
        assert result["status"] == "Success"
        assert result["warnings"] == []
        assert output_path.exists()
        assert list(output_root.glob("*.pdf")) == [output_path]
        assert "CIS" not in output_path.name.upper()
        assert customer_row["cis"] not in output_path.name

        reader = pypdf.PdfReader(str(output_path))
        assert len(reader.pages) == 3
        page_texts = [(page.extract_text() or "") for page in reader.pages]

        assert "FIRST_SOURCE_TEMPLATE" in page_texts[0]
        assert "Avery Tan" in page_texts[0]
        assert "25000" in page_texts[0]
        assert "19/05/2026" in page_texts[0]
        assert page_texts[1].strip() == ""
        assert "SECOND_SOURCE_TEMPLATE" in page_texts[2]
        assert "900101101234" in page_texts[2]
        assert "UT Subscription" in page_texts[2]
        assert "19/05/2026" in page_texts[2]


def main() -> None:
    test_package_engine_generates_ordered_combined_pdf()
    print("OK: procedure package smoke")


if __name__ == "__main__":
    main()
