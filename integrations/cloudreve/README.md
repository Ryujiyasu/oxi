# Oxi thumbnails for Cloudreve

[简体中文](README.zh-CN.md)

[Cloudreve](https://github.com/cloudreve/cloudreve) makes thumbnails of
documents by running LibreOffice: `soffice --headless --convert-to png
--outdir DIR FILE`. Oxi's `oxidocs` answers that same command line for
**.docx, .pptx and .xlsx**, laying the first page out the way Word, PowerPoint
and Excel do, and hands every other file to LibreOffice. Nothing in Cloudreve
changes: no plugin, no setting.

Per Word document in the container, oxidocs took 0.13 s and 100 MB where
LibreOffice took 2.0 s and 228 MB. Fifteen thumbnails at once all came out.
That is the case where LibreOffice's shared profile makes Cloudreve's docs
advise one or two workers.

## With Docker

```bash
docker compose up -d --build      # in this directory
```

The [Dockerfile](Dockerfile) starts from the official `cloudreve/cloudreve`
image and adds:

- oxidocs, the static release pinned by version and SHA-256, in `/opt/oxidocs`
- `/usr/local/bin/soffice` → oxidocs. It comes before `/usr/bin` on PATH, so
  it is the `soffice` Cloudreve's default setting runs.
- `OXIDOCS_SOFFICE_FALLBACK=/usr/bin/soffice`, so .doc, .odt, .pdf, .txt and
  the rest still go to LibreOffice.

To use it with an existing installation, build this image and run it on the
same data volume. If the LibreOffice path under *Settings → Media processing*
was changed from `soffice` to a full path, set it to `/usr/local/bin/soffice`.

## Without Docker

Download `oxidocs-<arch>-unknown-linux-musl.tar.gz` from the
[releases](https://github.com/Ryujiyasu/oxi/releases) (`oxidocs-v*`), unpack
it to a directory of its own, run `oxidocs fonts-index` once there (it writes
the fonts it carries next to itself), and link it as `soffice`:

```bash
ln -s /opt/oxidocs/oxidocs /usr/local/bin/soffice
export OXIDOCS_SOFFICE_FALLBACK=/usr/bin/soffice   # in Cloudreve's environment
```

Or point Cloudreve's LibreOffice path at the link. Without LibreOffice
installed, leave the fallback unset and cut the LibreOffice extension list
to `docx,docm,dotx,pptx,ppsx,potx,xlsx,xlsm,xltx`. Any other extension would
fail, and Cloudreve stops trying a file whose thumbnail failed once.

## Check

`e2e/run.sh` builds the image, starts Cloudreve, uploads a .docx, .pptx,
.xlsx and .txt, and fetches their thumbnails. LibreOffice is reached through
a logging wrapper, and the check fails if any of the three Office files went
to it, or if the text file did not.
