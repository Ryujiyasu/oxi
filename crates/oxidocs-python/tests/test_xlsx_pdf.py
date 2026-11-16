# SPDX-License-Identifier: MIT OR Apache-2.0
"""Native PDF integration: PYTHONPATH=python OXIDOCS_BINARY=... python this_file.py."""
from pathlib import Path
import subprocess
import tempfile
import unittest
import zipfile

import oxidocs

ROOT = Path(__file__).resolve().parents[3]


class SpreadsheetPdf(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="oxi-xlsx-pdf-")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def workbook(self, extra="", names=""):
        path = self.folder / "input.xlsx"
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')
            package.writestr("xl/workbook.xml", '''<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Invoice" sheetId="1" r:id="r1"/></sheets><definedNames>''' + names + "</definedNames></workbook>")
            package.writestr("xl/_rels/workbook.xml.rels", '''<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>''')
            package.writestr("xl/worksheets/sheet1.xml", '''<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>日本語の帳票</t></is></c><c r="B1"><f>2+3</f><v>5</v></c></row></sheetData>''' + extra + "</worksheet>")
        return path

    def test_cli_and_python_convert_a_workbook(self):
        source = self.workbook('<pageSetup paperSize="9" orientation="landscape"/>')
        result = oxidocs.xlsx_to_pdf(source, self.folder / "python.pdf", strict=True, timeout=60)
        self.assertEqual(result.pages, 1)
        self.assertTrue(result.output.read_bytes().startswith(b"%PDF-"))
        self.assertEqual(result.bytes, result.output.stat().st_size)
        process = subprocess.run([oxidocs.binary_path(), "xlsx-to-pdf", "--strict", "--quiet", str(source), str(self.folder / "cli.pdf")], capture_output=True, timeout=60)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertTrue((self.folder / "cli.pdf").read_bytes().startswith(b"%PDF-"))

    def test_strict_preserves_output_and_server_survives_refusal(self):
        source = self.workbook('<headerFooter><oddHeader>Not drawn</oddHeader></headerFooter>')
        output = self.folder / "kept.pdf"
        output.write_bytes(b"existing destination")
        with oxidocs.Converter(threads=1) as converter:
            with self.assertRaises(oxidocs.UnsupportedContent) as caught:
                converter.convert(source, output, strict=True, timeout=60)
            self.assertIn("header-footer", caught.exception.kinds)
            self.assertEqual(output.read_bytes(), b"existing destination")
            self.assertEqual(converter.convert(source, output, timeout=60).pages, 1)

    def test_workbook_content_detection_and_mixed_formats(self):
        source = self.workbook()
        renamed = source.with_suffix(".bin")
        source.rename(renamed)
        with oxidocs.Converter(threads=2) as converter:
            self.assertEqual(converter.convert(renamed, self.folder / "book.pdf", timeout=60).pages, 1)
            slides = ROOT / "tests/fixtures/multi_slide.pptx"
            self.assertEqual(converter.convert(slides, self.folder / "slides.pdf", timeout=60).pages, 3)
            doc = ROOT / "tests/fixtures/basic_test.docx"
            self.assertGreaterEqual(converter.convert(doc, self.folder / "doc.pdf", timeout=60).pages, 1)

    def test_invalid_print_area_is_an_error_and_server_keeps_running(self):
        source = self.workbook(names='<definedName name="_xlnm.Print_Area" localSheetId="0">Invoice!$A$1:$XFD$1048576</definedName>')
        with oxidocs.Converter(threads=1) as converter:
            with self.assertRaises(oxidocs.ConversionError):
                converter.convert(source, self.folder / "bad.pdf", timeout=60)
            self.assertFalse((self.folder / "bad.pdf").exists())
            self.workbook()
            self.assertEqual(converter.convert(source, self.folder / "ok.pdf", timeout=60).pages, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
