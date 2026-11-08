# Oxi for Gotenberg

A [Gotenberg](https://gotenberg.dev) 9 module that converts `.docx` to PDF with the Oxi engine. Without it, Gotenberg converts them with LibreOffice. A request Oxi cannot fully serve goes to LibreOffice exactly as before, so turning the module on changes the engine behind `/forms/libreoffice/convert` and nothing else.

- **Same API.** Existing clients keep their requests. Each response says which engine produced it.
- **No office suite on the docx path.** Oxi is a single Rust binary: no LibreOffice profile and no UNO. Conversion stays on your own server, like the rest of Gotenberg.
- **Fast.** The median document takes 0.08 s, against 0.89 s through LibreOffice on the same machine. See [Measured](#measured).
- **Fallback, not guesswork.** Oxi hands a document to LibreOffice when it holds content Oxi knows it does not draw, or when Oxi's pages show less than 90% of the document's own text.

## Quick start

```sh
docker build -t gotenberg-oxi .
docker run --rm -p 3000:3000 gotenberg-oxi

curl -s -D - -o out.pdf \
  -F files=@report.docx \
  http://localhost:3000/forms/libreoffice/convert | grep -i '^gotenberg-engine'
# Gotenberg-Engine: oxi
```

The image is the official `gotenberg/gotenberg:9.1.0` with two changes: its `gotenberg` binary is rebuilt with this module, and `oxidocs` is installed beside it. The image sets `OXI_ENABLE=true`. Chromium, LibreOffice, the PDF engines and every other route behave as in the official image.

## Which engine converts a request

Oxi takes a request on `/forms/libreoffice/convert` when every file in it is a `.docx` (stage files aside) and no LibreOffice-only option is set. Every other request goes to LibreOffice, and the response says why:

```
Gotenberg-Engine: oxi | libreoffice
Gotenberg-Engine-Reason: <reason>          (only when LibreOffice converted)
```

| Reason | When |
|---|---|
| `requested` | The client sent `Gotenberg-Oxi: off`. |
| `webhook` | The request asks for a webhook callback. |
| `option: <field>` | A LibreOffice option is set to anything but its default, such as `landscape`, `nativePdfa` or `paperWidth`. |
| `unknown-field: <field>` | The form has a field the route does not know. LibreOffice then rejects it with its usual 400. |
| `invalid-form` | The `stages` field does not parse. LibreOffice then reports the error as usual. |
| `no-docx`, `other-files` | No `.docx` in the request, or files of other types along with the `.docx`. |
| `unsupported-content: <kinds>` | The document holds content Oxi does not draw (for example `smartart` with no saved drawing, or `chart` for chartex). The same applies if Oxi draws less than 90% of the document's own text (`undrawn-text (84% drawn)`). |
| `oxi-timeout`, `oxi-failed` | Oxi ran out of time or failed. The module logs the error. |

The `stages` post-processing field (merge, split, PDF/A, PDF/UA, metadata, watermark, encryption…) runs on Oxi's PDFs exactly as on LibreOffice's. So does `downloadFrom`. Output files are named as on the LibreOffice route (`report.docx.pdf`).

## Configuration

| Flag | Environment | Default | Meaning |
|---|---|---|---|
| `--oxi-enable` | `OXI_ENABLE` | `false` (`true` in this image) | Convert `.docx` with Oxi |
| `--oxi-bin-path` | `OXI_BIN_PATH` | `oxidocs` | The oxidocs binary |
| `--oxi-threads` | `OXI_THREADS` | `2` | Documents converted in parallel |
| `--oxi-timeout` | `OXI_TIMEOUT` | `0` | Time limit for one document, retry included. `0` gives it 90% of what is left of the request's `--api-timeout`. A document that is slow for Oxi (in practice, ten thousand paragraphs or more) is slow for LibreOffice too, so falling back early saves nothing. |
| `--oxi-fallback` | `OXI_FALLBACK` | `true` | When `false`, answer 422 (unsupported content) or 500 (failure) instead of falling back |

Oxi runs as one `oxidocs serve` process (JSON lines over stdin/stdout). It starts on the first `.docx`. If the process crashes or a document runs out of time, the module restarts it. A document that was in the process when it went down is retried once.

## Measured

On 2026-11-08 we ran 4,988 `.docx` files through this image, at Oxi `7774ac9c`. The files come from [docx-corpus](https://docxcorp.us): 4,000 English and 988 Japanese, taken from the end of each type's manifest. Each file was sent twice, once as deployed and once with `Gotenberg-Oxi: off` (LibreOffice). Four requests ran at a time, with `API_TIMEOUT=170s`.

| | Oxi (as deployed) | LibreOffice only |
|---|---|---|
| Answered | 4,988 / 4,988 | 4,982 / 4,988 (6 rejected with 400) |
| Converted by Oxi | 4,977 (99.8%) | — |
| Fell back to LibreOffice | 11: 10 `undrawn-text`, 1 `chart` | — |
| Median / p99 / max per document | 0.08 s / 13.7 s / 60 s | 0.89 s / 27.3 s / 119 s |
| Total, English (4,000) | 2,885 s | 8,056 s |
| Total, Japanese (988) | 169 s | 1,308 s |

The six documents LibreOffice could not convert were all converted on the deployed path.

This run measures robustness and speed, not fidelity. Fidelity against Microsoft Word, compared with LibreOffice and other engines on documents no engine was tuned on, is in the main [README](../../README.md#layout-accuracy-vs-microsoft-word).

## Testing

```sh
# Go tests against a real oxidocs build
OXI_TEST_BIN=/path/to/oxidocs \
OXI_TEST_DOCX=$PWD/e2e/samples/simple.docx \
OXI_TEST_UNSUPPORTED_DOCX=$PWD/e2e/samples/unsupported.docx \
go test ./...

# End-to-end against a running image
e2e/e2e.sh http://localhost:3000
```

The e2e script checks the routing above: Oxi takes plain `.docx`, `stages` merge and PDF/A run on Oxi's output, and options, unsupported content, other file types and the opt-out header go to LibreOffice. It also times 20 conversions each way. The sample documents come from other open-source projects' test suites; [e2e/samples/SOURCES.txt](e2e/samples/SOURCES.txt) lists them.

## License

MIT OR Apache-2.0, the same as Oxi's language bindings. The Oxi engine itself is MPL-2.0. Gotenberg is MIT.
