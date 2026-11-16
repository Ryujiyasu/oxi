# SPDX-License-Identifier: MIT OR Apache-2.0
"""Contract tests against the actual compiled OxiCells engine, never a mock."""
import decimal
import json
import math
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

import oxicells as ox


class EngineDiagnostics(unittest.TestCase):
    def launch(self, source):
        # Inject a real child-process failure, exercising the actual pipes and
        # reader threads without introducing a deliberate panic into the engine.
        real_popen = subprocess.Popen
        def start(_command, **options):
            return real_popen([sys.executable, "-c", source], **options)
        return patch("oxicells.subprocess.Popen", side_effect=start)

    def test_engine_exit_keeps_a_bounded_stderr_tail(self):
        source = ("import sys; sys.stdin.readline(); "
                  "sys.stderr.buffer.write(b'FIRST_DIAGNOSTIC\\n'+b'x'*65536+"
                  "'\\npanic fixture: 日本語\\n'.encode('utf-8')); "
                  "sys.stderr.flush(); sys.exit(37)")
        with self.launch(source), self.assertRaises(ox.OxiCellsError) as raised:
            ox.open("unused.xlsx", binary=sys.executable, timeout=10)
        message = str(raised.exception)
        self.assertIn("exit code 37", message)
        self.assertIn("panic fixture: 日本語", message)
        self.assertNotIn("FIRST_DIAGNOSTIC", message)
        self.assertLess(len(message), 17000)

    def test_stderr_flood_does_not_block_valid_protocol_responses(self):
        source = ("import json,sys\n"
                  "for line in sys.stdin:\n"
                  " r=json.loads(line)\n"
                  " sys.stderr.buffer.write(b'x'*65536);sys.stderr.flush()\n"
                  " outcome={'protocol':1,'sheets':[],'hasMacros':False} if r['op']=='open' else {}\n"
                  " print(json.dumps({'id':r['id'],'ok':True,'outcome':outcome}),flush=True)\n")
        with self.launch(source), ox.open("unused.xlsx", binary=sys.executable, timeout=10) as wb:
            self.assertFalse(wb.has_macros)
            self.assertEqual(wb.sheets, ())


class ExecutionEnvironmentValidation(unittest.TestCase):
    def test_invalid_environment_fails_before_starting_an_engine(self):
        for options, error in [
            ({"error_language": 932}, TypeError),
            ({"error_language": "fr-FR"}, ValueError),
            ({"error_language": ""}, ValueError),
            ({"vba_help_file": Path("help.chm")}, TypeError),
            ({"host_ansi_code_page": True}, TypeError),
            ({"host_ansi_code_page": 932.0}, TypeError),
            ({"host_ansi_code_page": "932"}, TypeError),
            ({"host_ansi_code_page": 65001}, ValueError),
        ]:
            with self.subTest(options=options), patch("oxicells.subprocess.Popen") as start:
                with self.assertRaises(error): ox.open("unused.xlsx", **options)
                start.assert_not_called()


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

    def test_inspection_reports_capabilities_without_execution_or_state_reset(self):
        self.wb.run("PythonBridge.Step", 2)
        report = self.wb.inspect_macros()
        self.assertTrue(report.has_macros)
        self.assertIn("PythonBridge.Auto_Open", report.auto_run)
        self.assertTrue(report.analysis_complete, report.modules)
        self.assertFalse(report.compatibility_verified)
        self.assertEqual(report.unparsed_lines, 0)
        self.assertIn("PythonBridge", [module["name"] for module in report.modules])
        self.assertTrue(any(signal["capability"] == "starts_program" for signal in report.external_calls))
        self.assertTrue(any("MSXML2.XMLHTTP" in signal["what"] for signal in report.external_calls))
        self.assertTrue(any(item["module"] == "PythonBridge" and item["line"] > 0
                            for item in report.unresolved_calls))
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 5)
        standalone = ox.inspect_macros(self.fixture, timeout=60)
        self.assertEqual(standalone.auto_run, report.auto_run)
        self.assertEqual(standalone.external_calls, report.external_calls)

    def test_macro_free_inspection_reports_absence(self):
        with tempfile.TemporaryDirectory(prefix="oxicells-inspection-") as folder:
            path = Path(folder) / "plain.xlsx"
            self.wb.save(path)
            report = ox.inspect_macros(path, timeout=60)
            self.assertFalse(report.has_macros)
            self.assertEqual(report.modules, ())
            self.assertEqual(report.signals, ())
            self.assertEqual(report.auto_run, ())
            self.assertTrue(report.analysis_complete)
            self.assertFalse(report.compatibility_verified)

    def test_inspection_exposes_parameters_and_callable_arity(self):
        report = self.wb.inspect_macros()
        modules = {module["name"]: module for module in report.modules}
        procedures = {procedure["name"]: procedure for procedure in modules["PythonBridge"]["procedures"]}
        step = procedures["Step"]
        self.assertEqual(step["kind"], "Function")
        self.assertEqual(step["visibility"], "Public")
        self.assertEqual(step["return_type"], "Long")
        self.assertEqual(step["required_arguments"], 1)
        self.assertEqual(step["parameters"], [{"name":"delta","mode":"ByVal","type":"Long",
            "optional":False,"is_array":False,"has_default":False}])
        self.assertEqual(procedures["Auto_Open"]["kind"], "Sub")
        self.assertEqual(procedures["Auto_Open"]["parameters"], [])
        self.assertEqual(procedures["ReadCell"]["required_arguments"], 0)
        self.assertIsNone(modules["PythonBridge"]["parse_error"])
        destructor = next(p for p in modules["BridgeBox"]["procedures"] if p["name"]=="Class_Terminate")
        self.assertEqual(destructor["visibility"], "Private")
        self.assertEqual(self.wb.run("PythonBridge.Step", 3), 3)

    def test_code_page_override_matches_unicode_packages_and_keeps_errors_actionable(self):
        fixtures = Path(os.environ["OXICELLS_TEST_VBA_FIXTURES"])
        cases = json.loads((fixtures/"unicode_modules.json").read_text(encoding="utf-8"))
        failing_defaults = {"unicode_mixed", "unicode_workbook", "unicode_chart"}
        with tempfile.TemporaryDirectory(prefix="oxicells-code-page-") as folder:
            for index, case in enumerate(cases):
                path = fixtures/case["package_fixture"]
                with self.subTest(case=case["name"]):
                    if case["name"] in failing_defaults:
                        report = ox.inspect_macros(path, timeout=60)
                        errors = [m["parse_error"] for m in report.modules if m["parse_error"]]
                        self.assertTrue(errors, report.modules)
                        self.assertTrue(any("source_code_page=" in e["message"] for e in errors))
                        self.assertTrue(all("procedures" in m for m in report.modules))
                        with ox.open(path, timeout=60) as wrong:
                            with self.assertRaises(ox.OxiCellsError) as raised:
                                wrong.run(case["procedure"])
                            self.assertIn("source_code_page=", str(raised.exception))
                    report = ox.inspect_macros(path, source_code_page=932, timeout=60)
                    self.assertTrue(report.analysis_complete, report.modules)
                    with ox.open(path, source_code_page=932, timeout=60) as wb:
                        expected = case["expected"].replace(case["host_context"]["book_name"], path.name)
                        if case["name"] == "unicode_workbook":
                            # Actual saved-file Excel observations differ from
                            # this JSON's explicitly bound source-only context.
                            with self.assertRaises(ox.MacroError) as raised:
                                wb.run(case["procedure"])
                            self.assertIn("ActiveX component cannot create object", str(raised.exception))
                            self.assertIn("日本語台帳", str(raised.exception))
                        else:
                            self.assertEqual(wb.run(case["procedure"]), expected)
                        output = Path(folder)/f"unicode-{index}.xlsx"
                        wb.save(output)
                    with zipfile.ZipFile(output) as package:
                        self.assertIsNone(package.testzip())
                        self.assertNotIn("xl/vbaProject.bin", package.namelist())

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

    def test_finite_double_bits_survive_json_in_both_directions(self):
        for encoded in ["3fefffff7fffffff", "3fefffff80000000", "3fefffff80000001",
                        "400fffff80000000", "bfefffff80000000", "8000000000000000"]:
            with self.subTest(bits=encoded):
                payload = bytes.fromhex(encoded)
                value = struct.unpack(">d", payload)[0]
                signed = int.from_bytes(payload, "big", signed=True)
                # LongLong responses use exact decimal strings on the wire,
                # isolating incoming floating-point conversion from outgoing.
                self.assertEqual(self.wb.run("PythonBridge.DoubleBits", value),
                                 ox.LongLong(signed))
                # Send an exact integer to isolate outgoing numeric conversion.
                result = self.wb.run("PythonBridge.FromBits", ox.LongLong(signed))
                self.assertEqual(struct.pack(">d", result), payload)

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

    def test_input_functions_read_the_mounted_package_and_preserve_odd_bstr_bytes(self):
        for mode in range(4):
            for count in [1, 2]:
                with self.subTest(mode=mode, count=count):
                    answer = self.wb.run("PythonBridge.InputHeader", mode, count)
                    self.assertIsInstance(answer, ox.Array)
                    self.assertEqual(answer.dimensions, ((0, 8),))
                    byte_result = mode < 2
                    byte_length = count if byte_result else count * 2
                    self.assertEqual(answer.values[:4], ("String", 8, byte_length, byte_length // 2))
                    if byte_result and count == 1:
                        self.assertEqual(answer.values[4], ox.BStr(b"P"))
                    elif byte_result:
                        self.assertEqual(answer.values[4].encode("utf-16-le"), b"PK")
                    else:
                        self.assertEqual(answer.values[4], "PK"[:count])
                    self.assertEqual(answer.values[5:], (count + 1, False, count))

    def test_execution_environment_persists_across_calls(self):
        help_file = "C:\\example\\日本語\\VbLR6.chm"
        # Default StrConv uses the user locale; LCID 2048 selects the system
        # ANSI code page. These can differ in the same Windows environment.
        cases = [
            ({"error_language":"ja-JP", "host_ansi_code_page":932},
             "プロシージャの呼び出し、または引数が不正です。", "3F|82A0"),
            ({"error_language":"EN-us", "host_ansi_code_page":1252},
             "Invalid procedure call or argument", "3F|3F"),
        ]
        for options, description, encoded in cases:
            with self.subTest(options=options), ox.open(
                    self.fixture, vba_help_file=help_file, timeout=60, **options) as book:
                for _ in range(2):
                    value = book.run("PythonBridge.Environment")
                    self.assertIsInstance(value, ox.Array)
                    self.assertEqual(value.dimensions, ((0, 7),))
                    self.assertEqual(value.values[0], 5)
                    self.assertEqual(value.values[1], description)
                    self.assertTrue(value.values[2])
                    self.assertEqual(value.values[3], help_file)
                    self.assertEqual(value.values[5:], (encoded, 0))

    def test_protocol_rejects_invalid_environment_without_consuming_the_book(self):
        invalid = [{"errorLanguage":1}, {"errorLanguage":"fr-FR"},
                   {"vbaHelpFile":False}, {"hostAnsiCodePage":True},
                   {"hostAnsiCodePage":932.0}, {"hostAnsiCodePage":65001}]
        requests = [{"op":"open", "path":str(self.fixture), "seed":1, **v}
                    for v in invalid]
        requests += [{"op":"open", "path":str(self.fixture), "seed":1},
                     {"op":"run", "procedure":"PythonBridge.Step",
                      "args":[{"type":"long", "value":5}]}, {"op":"close"}]
        process = subprocess.run([ox.binary_path(), "serve"],
            input="\n".join(json.dumps({"id":i, "currentTime":45000.0, **r})
                            for i, r in enumerate(requests)) + "\n",
            text=True, encoding="utf-8", capture_output=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        self.assertEqual(process.returncode, 0, process.stderr)
        answers = [json.loads(line) for line in process.stdout.splitlines()]
        self.assertEqual(len(answers), len(requests))
        for answer in answers[:len(invalid)]: self.assertFalse(answer["ok"], answer)
        self.assertTrue(answers[-3]["ok"], answers[-3])
        self.assertEqual(answers[-2]["outcome"]["result"], {"type":"long", "value":5})

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
