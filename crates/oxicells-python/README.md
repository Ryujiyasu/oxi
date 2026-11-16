# OxiCells Python: workbook VBA

An experimental Python API for running **supported VBA** with Oxi's native engine, without Microsoft Excel. This package is under development and has not been released to PyPI. Complete Excel VBA compatibility has not been established.

```python
import oxicells

with oxicells.open("report.xlsm", timeout=60) as wb:
    wb.compile()                         # Check the project without running a macro.
    result = wb.run("Module1.MonthlyReport", 2026, 11)
    wb.save("result.xlsx")                # Write a macro-free workbook.
```

Opening a workbook does not execute its macros. Calls on the same `Workbook` preserve module variables, cell changes and returned object references. Saving as `.xlsx` or `.xltx` removes VBA from the output package; `.xlsm` and `.xltm` are macro-enabled formats. A new process that reopens a saved workbook starts with fresh module variables.

`compile()` reports project compilation errors. It does **not** certify that every Excel feature used by the project is implemented. Runtime errors raise `MacroError`, with `kind`, `line` and `details`; writes completed before an error remain in the live workbook. Protocol and file errors raise `OxiCellsError`. When the optional request timeout expires, the process is terminated and the workbook must be reopened.

Macro file operations use the engine's virtual `C:\work` namespace. The input package is mounted there under its basename. Arbitrary files on the Python host are not automatically mounted. Excel UI, external COM components and all VBA features are not yet covered by this API.

## Inspect source without running macros

```python
report = oxicells.inspect_macros("received.xlsm")
print(report.auto_run)          # Qualified automatic entrypoint names found in source.
print(report.external_calls)    # File, process, network, DLL and other external-call evidence.
print(report.unresolved_calls)  # Computed CreateObject/GetObject targets.
```

`wb.inspect_macros()` reads the already loaded project's source without running or resetting its macros. Reports include modules, procedures, API names, declarations and capability signals with exported-module line numbers. A signal's `source_line` points into the combined exported source; zero module lines indicate evidence without an exact position.

Parse errors and unread constructs remain visible. `analysis_complete` means the parser read the source; it does not mean dynamic targets were resolved, that the code is safe, or that its runtime features are implemented. `compatibility_verified` remains false. Automatic entrypoints describe Office behavior; opening or inspecting a workbook in OxiCells does not invoke them.

Procedure entries expose `kind`, effective `visibility`, and `parameters`. Each parameter includes its name, declared `type`, `mode` (`ByRef`, `ByVal` or `ParamArray`), `optional`, `is_array`, and `has_default`. `required_arguments` counts parameters that must be supplied; a `ParamArray` may be empty. Functions and Property Get procedures also report their `return_type`. A module that cannot be parsed retains `parse_error` and an empty `procedures` list.

```python
for module in report.modules:
    for procedure in module["procedures"]:
        print(module["name"], procedure["name"], procedure["kind"],
              procedure["visibility"], procedure["required_arguments"])
```

### VBA source code pages

The default reads the code page recorded in the VBA project. Some workbooks record a code page that differs from their source bytes. If the actual source encoding is known, pass `source_code_page` to `open()` or `inspect_macros()`; for Japanese Windows source in CP932:

```python
report = oxicells.inspect_macros("received.xlsm", source_code_page=932)
with oxicells.open("received.xlsm", source_code_page=932) as wb:
    result = wb.run("集計.Probe")
```

A wrong code page can corrupt names and produce a syntax error. The engine preserves the original error and adds a code-page hint when non-ASCII source fails to parse. Reopen the workbook with the appropriate override. This option does not change the encoding of Python strings.

Unexpected engine exits include the exit code and a bounded tail of engine stderr in `OxiCellsError`, so startup failures and panics retain their diagnostic text.

## Values and objects

Python integers use VBA Long or LongLong according to their range. `Byte`, `Integer`, `Long`, `LongLong`, `Single`, `Date` (OLE date serial) and `Currency` wrappers allow explicit subtypes. Python `decimal.Decimal` values retain their exact 96-bit VBA Decimal representation within its supported range and scale. `EMPTY`, `MISSING`, `NOTHING`, `NULL_STRING` and Python `None` (VBA Null) are distinct.

Lists and rectangular nested lists become VBA Variant arrays. Returned `Array` values expose `dimensions` as `(lower_bound, length)` pairs, flat row-major `values`, and `at(*indices)` using the original VBA subscripts. Explicit `Array` arguments can preserve typed elements and negative lower bounds. Strings preserve unpaired UTF-16 units; `BStr` represents raw BSTR bytes.

Returned `Object` values belong to their workbook session. Use `invoke("Method", ...)`, `get("Property")`, `set("Property", value)` and `release()`. Closing the workbook releases retained references. An object from another workbook session is rejected.

## Build and verify from the repository

From `crates/oxicells-python`:

```sh
python -m pip install "maturin>=1.16,<2"
python -m maturin build --release
python -m pip install /path/to/the/built/oxicells-wheel.whl
```

For source development, put `python/` on `PYTHONPATH` and set `OXICELLS_BINARY` to the built `oxicells` executable. A platform wheel contains that executable and the Python facade; no `oxidocs` executable is installed by this package.

From the repository root, the integration test builds a temporary macro workbook and runs the Python contract tests against the actual native executable:

```sh
cargo test -p oxidocs-cli --features vba-execution --bin oxicells --test cells_python -- --nocapture
```

The bundled native engine is MPL-2.0; the Python binding code is MIT OR Apache-2.0.

## 日本語

Excel 本体を使わず、Oxi のエンジンで**対応済みの VBA** を実行する実験的な Python API です。開発中で、PyPI には未公開です。Excel VBA の全機能との完全一致を保証する段階ではありません。

```python
import oxicells

with oxicells.open("帳票.xlsm", timeout=60) as wb:
    wb.compile()
    結果 = wb.run("Module1.月次集計", 2026, 11)
    wb.save("結果.xlsx")
```

開くだけではマクロを実行しません。同じ `Workbook` への呼び出しでは、変数・セル変更・返されたオブジェクトを引き継ぎます。`.xlsx` と `.xltx` への保存では、出力ファイルから VBA を取り除きます。`.xlsm` と `.xltm` はマクロ付きの保存形式です。保存したファイルを別のセッションで開くと、モジュール変数は初期化されます。

`compile()` は実行せずにプロジェクトのコンパイルを確認します。使用する Excel 機能がすべて対応済みであることを判定するものではありません。VBA の実行エラーは `MacroError`、通信やファイルのエラーは `OxiCellsError` で伝えます。実行エラーより前に行われたセル変更は残ります。指定した制限時間を超えるとプロセスを終了するため、再度開く必要があります。

マクロのファイル操作には仮想の `C:\work` を使い、入力ファイルだけを配置します。Python を動かす PC のほかのファイルを自動では読み込みません。Excel の画面操作や外部 COM など、まだ扱えない機能があります。整数・小数の型や配列の添字も保持して受け渡します。

`oxicells.inspect_macros("受領ファイル.xlsm")` または `wb.inspect_macros()` で、マクロを動かさずにソースを調べられます。`auto_run` はソースにある自動実行の入口、`external_calls` は外部呼び出しの兆候、`unresolved_calls` は呼び出し先が実行時まで決まらない箇所です。モジュール名とエクスポートしたソースの行番号も返します。読み取れなかったコードや解析エラーは隠しません。

`analysis_complete` はソースを解析できたという意味です。安全性や、Oxi で使われる全機能を実行できることの保証ではなく、`compatibility_verified` は false のままです。調査のためにマクロを実行したり、実行中の変数をリセットしたりはしません。


### プロシージャー情報と文字コード

`report.modules` の各モジュールには、解析に失敗した場合も `procedures` のリストがあります。解析できないモジュールでは空のリストとなり、`parse_error` に原因が残ります。各プロシージャーには `kind`（Sub・Function・Property Get/Let/Set）、有効な `visibility`、`parameters`、`required_arguments` を返します。引数ごとの名前・宣言型 `type`・渡し方 `mode`・省略可否 `optional`・配列かどうか `is_array`・既定値の有無 `has_default` を調べられます。ParamArray は空でも呼び出せます。Function と Property Get では戻り値の型 `return_type` も返します。

VBA ソースは、標準ではプロジェクトに記録されたコードページで読みます。記録と実際の文字コードが異なるファイルでは、文字化けした名前や構文エラーになることがあります。実際のソースが日本語 Windows の CP932 だと分かっている場合は、次のように指定します。

```python
report = oxicells.inspect_macros(受領ファイル.xlsm, source_code_page=932)
with oxicells.open(受領ファイル.xlsm, source_code_page=932) as wb:
    result = wb.run(集計.Probe)
```

非 ASCII 文字を含むソースを解析できなかった場合は、元のエラーに `source_code_page` の案内を添えます。適切な文字コードを指定して開き直してください。この指定は Python 文字列の文字コードを変更するものではありません。

エンジンが異常終了したときは、終了コードと長さを制限した stderr の末尾を `OxiCellsError` に含めます。起動失敗や panic の原因を Python 側でも確認できます。
