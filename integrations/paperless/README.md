# Oxi .docx parser for Paperless-ngx

A [Paperless-ngx](https://docs.paperless-ngx.com) v3 parser plugin. It converts `.docx` documents inside the Paperless container, with the [Oxi](https://github.com/Ryujiyasu/oxi) engine. It does not use Tika or Gotenberg.

Out of the box, Paperless-ngx sends Office documents to two extra services. Tika extracts their text. Gotenberg, which runs LibreOffice, makes the PDF that Paperless displays and archives. With this plugin, `.docx` needs neither:

- **PDF:** made by the bundled `oxidocs` engine (installed from PyPI with the plugin). When your output type is PDF/A (the default), the container's own `ocrmypdf` converts it, the same way Paperless makes its other archives.
- **Text, date and metadata:** read straight from the document's XML, so search sees whole words in reading order.
- **Fallback:** a document holding content Oxi does not draw (some charts, SmartArt saved without its drawing, or less than 90% of its text drawn) goes to the built-in Tika parser when Tika is enabled. Without Tika, it is converted as well as Oxi can.
- **Other formats:** `.doc`, `.xlsx`, `.pptx`, `.odt` and the rest stay with Tika as before.

The plugin itself makes no network requests.

## Install with Docker

1. Next to your `docker-compose.yml`, create this `Dockerfile`:

   ```dockerfile
   FROM ghcr.io/paperless-ngx/paperless-ngx:latest
   RUN pip install --no-cache-dir paperless-ngx-oxidocs
   ```

   Leave the image's user as it is: its `/init` starts as root and drops privileges itself.

2. Point the `webserver` service at it instead of the upstream image:

   ```yaml
   services:
     webserver:
       build: .
       # ...keep the rest of your existing config (env, volumes, depends_on, ...)
   ```

3. Rebuild and restart:

   ```bash
   docker compose build webserver && docker compose up -d
   ```

At startup the log shows `Loaded third-party parser 'Oxi DOCX Parser' ...`. A consumed `.docx` logs `Oxi converted report.docx: 3 pages in 120 ms`.

You can keep the Tika and Gotenberg services for the formats this plugin doesn't handle and for the fallback. If you only store `.docx` among Office formats, you can turn Tika off (`PAPERLESS_TIKA_ENABLED=false`) and drop both services.

Supported image architectures: `amd64` and `arm64`, the two Paperless publishes.

## Install on bare metal

Install into the Python environment Paperless runs in (often a venv under `/opt/paperless`), then restart its services:

```bash
/opt/paperless/venv/bin/pip install paperless-ngx-oxidocs
```

PDF/A output needs `ocrmypdf` on the PATH, which a Paperless install already has.

## How good is the PDF?

The engine was measured on 50 English and 50 Japanese documents. The documents were frozen before any engine rendered them, and each output was scored against Microsoft Word's own PDF. On the same Linux machine, Oxi beat LibreOffice, which is what Gotenberg runs, on every measure:

| | English SSIM | page count matches Word | Japanese SSIM | page count matches Word |
|---|---|---|---|---|
| Oxi | 0.884 | 47 / 50 | 0.797 | 45 / 50 |
| LibreOffice | 0.841 | 39 / 50 | 0.769 | 26 / 50 |

Across 4,988 documents, the median conversion took 0.08 s against 0.89 s for LibreOffice. The method is in [integrations/gotenberg](../gotenberg#measured).

## Checking it

`e2e/run.sh` builds Paperless-ngx 3.4.2 with this plugin. It then consumes three `.docx` files with no Tika and no Gotenberg, and checks what Paperless stored for each: page count, text, and archive PDF. `pytest tests/` checks the text extraction on its own.

## Notes

- Paperless-ngx does not support third-party plugins: report problems with this plugin at [Ryujiyasu/oxi](https://github.com/Ryujiyasu/oxi/issues), not to Paperless-ngx.
- This plugin was written with AI assistance (Claude).
- License: MIT OR Apache-2.0. The `oxidocs` engine it installs is MPL-2.0.
