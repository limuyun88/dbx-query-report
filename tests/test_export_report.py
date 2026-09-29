import copy
import html
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "export_report.py"
FIXTURE = ROOT / "tests" / "fixtures" / "report.json"
SPEC = importlib.util.spec_from_file_location("export_report", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
BASE = json.loads(FIXTURE.read_text(encoding="utf-8"))
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class ReportExportTests(unittest.TestCase):
    def test_end_to_end_export_and_saved_values(self):
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "report"
            result = subprocess.run(
                [sys.executable, str(SCRIPT), "--input", str(FIXTURE), "--out-dir", str(output)],
                text=True, capture_output=True, check=False
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual({p.name for p in output.iterdir()}, {"report.xlsx", "report.html", "report.sql"})
            self.assertEqual((output / "report.sql").read_text(encoding="utf-8"),
                             "\n\n".join(BASE["sql"]) + "\n")
            page = (output / "report.html").read_text(encoding="utf-8")
            self.assertIn("示例商品乙", page)
            self.assertIn("850.10", page)
            self.assertIn("15.01%", page)
            self.assertIn("无符合条件记录", MODULE.build_html(
                MODULE.validate({**copy.deepcopy(BASE), "datasets": [
                    {**copy.deepcopy(BASE["datasets"][0]), "rows": []}
                ]})
            ))
            with zipfile.ZipFile(output / "report.xlsx") as book:
                self.assertIsNone(book.testzip())
                names = set(book.namelist())
                self.assertIn("xl/worksheets/sheet2.xml", names)
                sheet1 = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
                sheet2 = ET.fromstring(book.read("xl/worksheets/sheet2.xml"))
                self.assertEqual(sheet1.find(f".//{MAIN}c[@r='C7']/{MAIN}v").text, "1000.25")
                self.assertEqual(sheet1.find(f".//{MAIN}c[@r='C8']/{MAIN}v").text, "850.10")
                self.assertEqual(sheet2.find(f".//{MAIN}c[@r='D8']/{MAIN}v").text, "52.05")
                self.assertIsNone(sheet1.find(f".//{MAIN}c[@r='D7']/{MAIN}v"))

    def test_text_is_inert_in_excel_and_html(self):
        data = copy.deepcopy(BASE)
        attack = '=HYPERLINK("https://example.invalid","x")<script>alert(1)</script>'
        data["datasets"][1]["rows"][0]["name"] = attack
        validated = MODULE.validate(data)
        page = MODULE.build_html(validated)
        self.assertNotIn("<script>alert", page)
        self.assertIn(html.escape(attack), page)
        with zipfile.ZipFile(__import__("io").BytesIO(MODULE.build_xlsx(validated))) as book:
            sheet = ET.fromstring(book.read("xl/worksheets/sheet2.xml"))
            cell = sheet.find(f".//{MAIN}c[@r='C7']")
            self.assertEqual(cell.attrib["t"], "inlineStr")
            self.assertIsNone(cell.find(f"{MAIN}f"))
            self.assertEqual(cell.find(f".//{MAIN}t").text, attack)

    def test_rejects_truncation_bad_numbers_and_existing_output(self):
        for mutation in (
            lambda d: d.update(complete=False),
            lambda d: d["datasets"][0]["rows"][0].update(sales="NaN"),
            lambda d: d["datasets"][0]["rows"][0].update(customers="1.5"),
            lambda d: d["datasets"][0]["rows"][0].update(sales="0.1234567890123456"),
            lambda d: d["datasets"][0]["rows"][0].pop("sales"),
        ):
            data = copy.deepcopy(BASE)
            mutation(data)
            with self.assertRaises(ValueError):
                MODULE.validate(data)
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "existing"
            output.mkdir()
            with self.assertRaises(ValueError):
                MODULE.export(MODULE.validate(copy.deepcopy(BASE)), output)

    def test_zero_negative_and_null(self):
        data = copy.deepcopy(BASE)
        rows = data["datasets"][0]["rows"]
        rows[0]["sales"] = "0"
        rows[1]["sales"] = "-20.25"
        result = MODULE.build_html(MODULE.validate(data))
        self.assertIn("-20.25", result)
        self.assertIn("—", result)
        self.assertIn("#b34747", result)

    def test_two_months_ten_products_each_remain_complete(self):
        data = copy.deepcopy(BASE)
        template = data["datasets"][1]
        template["rows"] = [
            {
                "month": month,
                "rank": rank,
                "name": f"模拟商品{month[-2:]}-{rank}",
                "sales": str(1000 - rank * 12 + (100 if month.endswith("02") else 0)),
            }
            for month in ("2026-01", "2026-02")
            for rank in range(1, 11)
        ]
        checked = MODULE.validate(data)
        with zipfile.ZipFile(__import__("io").BytesIO(MODULE.build_xlsx(checked))) as book:
            sheet = ET.fromstring(book.read("xl/worksheets/sheet2.xml"))
            for row_number, expected in ((7, "2026-01"), (16, "2026-01"),
                                         (17, "2026-02"), (26, "2026-02")):
                self.assertEqual(
                    sheet.find(f".//{MAIN}c[@r='A{row_number}']/{MAIN}is/{MAIN}t").text,
                    expected,
                )
        page = MODULE.build_html(checked)
        self.assertEqual(page.count('class="barrow"'), 22)
        self.assertIn("模拟商品02-10", page)


if __name__ == "__main__":
    unittest.main()
