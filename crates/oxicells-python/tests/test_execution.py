# SPDX-License-Identifier: MIT OR Apache-2.0
"""Contract tests against the actual compiled OxiCells engine, never a mock."""
import decimal
import json
import math
import os
from pathlib import Path
import struct
import subprocess
import tempfile
import unittest
import zipfile

import oxicells as ox


@unittest.skipUnless(os.environ.get("OXICELLS_TEST_FIXTURE"), "run the cells_python Cargo integration test")
class NativeWorkbook(unittest.TestCase):
    def setUp(self):
        self.fixture = Path(os.environ["OXICELLS_TEST_FIXTURE"])
        self.wb = ox.open(self.fixture, timeout=60)
        self.addCleanup(self.wb.close)

    def test_state_and_compile_do_not_run_auto_open(self):
        self.assertTrue(self.wb.has_macros)
        self.assertNotEqual(self.wb.run("PythonBridge.ReadCell"), 999)
        self.assertEqual(self.wb.run("PythonBridge.Step", 2), 2)
        self.wb.compile()
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 5)

    def test_exact_scalar_values_and_invalid_arguments_keep_state(self):
        self.wb.run("PythonBridge.Step", 2)
        values = [ox.Byte(255), ox.Integer(-32768), ox.Long(2147483647),
                  ox.LongLong(9007199254740993), ox.Single(1.25), ox.Date(1.5),
                  ox.Currency("-922337203685477.5808"),
                  decimal.Decimal("7.9228162514264337593543950335"),
                  "日本語", "\ud800", ox.BStr(b"A\x00B"), ox.EMPTY, None]
        for value in values:
            with self.subTest(value=repr(value)):
                actual = self.wb.run("PythonBridge.Echo", value)
                self.assertEqual(actual, value)
                self.assertIs(type(actual), type(value))
        for value in [ox.Integer(32768), decimal.Decimal("1e-29")]:
            with self.assertRaises(ox.OxiCellsError):
                self.wb.run("PythonBridge.Echo", value)
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 5)

    def test_nonfinite_values_preserve_bits(self):
        value = struct.unpack(">d", bytes.fromhex("7ff8000000000001"))[0]
        actual = self.wb.run("PythonBridge.Echo", value)
        self.assertTrue(math.isnan(actual))
        self.assertEqual(struct.pack(">d", actual).hex(), "7ff8000000000001")
        actual = self.wb.run("PythonBridge.Echo", ox.Single(float("inf")))
        self.assertIs(type(actual), ox.Single)
        self.assertEqual(actual, float("inf"))

    def test_arrays_retain_vba_bounds_and_accept_python_tables(self):
        values = ((1, 2), (3, 4))
        result = self.wb.run("PythonBridge.Echo", values)
        self.assertEqual(result.dimensions, ((0, 2), (0, 2)))
        self.assertEqual(result.at(1, 0), 3)
        typed = ox.Array(tuple(ox.Integer(v) for v in (7, 8, 9, 10)),
                         ((-2, 2), (1, 2)), ox.Integer(0), False)
        # Assignment to the Variant function result makes its array copy dynamic.
        self.assertEqual(self.wb.run("PythonBridge.Echo", typed),
                         ox.Array(typed.values, typed.dimensions, typed.element_default, True))
        bounds = self.wb.run("PythonBridge.Bounds", [[1, 2], [3, 4]])
        self.assertEqual(bounds.values, (8204, 0, 1, 0, 1, 4))
        with self.assertRaises(IndexError): result.at(2, 0)

    def test_runtime_errors_preserve_writes_and_session(self):
        self.wb.run("PythonBridge.Step", 2)
        with self.assertRaises(ox.MacroError) as raised:
            self.wb.run("PythonBridge.FailAfterWrite")
        self.assertTrue(str(raised.exception))
        self.assertEqual(self.wb.run("PythonBridge.ReadCell"), 9)
        # An unhandled VBA runtime error ends execution and resets module storage.
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 3)

    def test_objects_are_owned_and_release_runs_destructors(self):
        box = self.wb.run("PythonBridge.MakeBox")
        self.assertEqual(box.invoke("Touch"), 17)
        self.assertEqual(self.wb.run("PythonBridge.CallBox", box), 17)
        with ox.open(self.fixture, timeout=60) as other:
            with self.assertRaises(ox.OxiCellsError): other.run("PythonBridge.CallBox", box)
            self.assertEqual(other.run("PythonBridge.Step", 4), 4)
        box.release()
        self.assertEqual(self.wb.run("PythonBridge.ReadCell"), 41)
        with self.assertRaises(ox.OxiCellsError): box.invoke("Touch")

    def test_save_reopen_and_live_project(self):
        self.wb.run("PythonBridge.Step", 23)
        with tempfile.TemporaryDirectory(prefix="oxicells-save-") as folder:
            xlsx = Path(folder) / "結果.xlsx"
            self.wb.save(xlsx)
            with zipfile.ZipFile(xlsx) as archive:
                self.assertNotIn("xl/vbaProject.bin", archive.namelist())
            with ox.open(xlsx, timeout=60) as plain:
                self.assertFalse(plain.has_macros)
            self.assertEqual(self.wb.run("PythonBridge.Step", 1), 24)
            self.wb.save(xlsx)  # Replace an existing destination through a staged save.
            xlsm = Path(folder) / "結果.xlsm"
            self.wb.save(xlsm)
            with ox.open(xlsm, timeout=60) as restored:
                self.assertTrue(restored.has_macros)
                self.assertEqual(restored.run("PythonBridge.ReadCell"), 24)
                self.assertEqual(restored.run("PythonBridge.Step", 1), 1)

    def test_failed_os_save_preserves_the_destination_and_session(self):
        self.wb.run("PythonBridge.Step", 2)
        with tempfile.TemporaryDirectory(prefix="oxicells-save-error-") as folder:
            destination = Path(folder) / "existing.xlsx"
            destination.mkdir()
            marker = destination / "keep.txt"
            marker.write_text("keep", encoding="utf-8")
            with self.assertRaises(ox.OxiCellsError): self.wb.save(destination)
            self.assertEqual(marker.read_text(encoding="utf-8"), "keep")
            self.assertEqual(sorted(p.name for p in Path(folder).iterdir()), ["existing.xlsx"])
            self.assertEqual(self.wb.run("PythonBridge.Step", 3), 5)

    def test_bad_save_and_closed_session_are_explicit(self):
        self.wb.run("PythonBridge.Step", 2)
        with self.assertRaises(ox.OxiCellsError): self.wb.save(self.fixture.with_suffix(".bad"))
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 5)
        self.wb.close()
        self.wb.close()
        with self.assertRaises(ox.OxiCellsError): self.wb.run("PythonBridge.Step", 1)

    def test_bad_protocol_request_does_not_consume_the_workbook(self):
        requests = [
            {"id": 1, "op": "open", "path": str(self.fixture), "seed": 1},
            {"id": 2, "op": "compile", "args": [{"type": "long", "value": 1}]},
            {"id": 3, "op": "run", "procedure": "PythonBridge.Step", "args": [{"type": "long", "value": 5}]},
            {"id": 4, "op": "close"},
        ]
        process = subprocess.run([ox.binary_path(), "serve"],
            input="\n".join(json.dumps({"currentTime": 45000.0, **r}) for r in requests) + "\n",
            text=True, encoding="utf-8", capture_output=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        self.assertEqual(process.returncode, 0, process.stderr)
        answers = [json.loads(line) for line in process.stdout.splitlines()]
        self.assertFalse(answers[1]["ok"])
        self.assertEqual(answers[2]["outcome"]["result"], {"type": "long", "value": 5})


if __name__ == "__main__": unittest.main(verbosity=2)
