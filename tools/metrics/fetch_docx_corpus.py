"""Fetch a stratified English sample of superdoc-dev/docx-corpus (ODC-BY).

736K real-world .docx from Common Crawl, classified by type/topic/language.
Manifest API: https://api.docxcorp.us/manifest?type=<type>&lang=<lang>
Files: https://docxcorp.us/documents/<sha256>.docx

Usage:
  python fetch_docx_corpus.py fetch [n_per_type] [lang]   # default 10, en
  python fetch_docx_corpus.py harness                     # Oxi crash/page pass

Quarantine on download: valid zip, has word/document.xml, no VBA project
(vbaProject.bin), size <= MAX_MB. Files land in
pipeline_data/docx_corpus/<lang>/<type>/<sha16>.docx (deterministic: the
first N manifest entries per type).
"""
import os, sys, json, zipfile, io, urllib.request

UA = {"User-Agent": "oxi-corpus-fetch/1.0 (+https://gitlab.com/Ryujiyasu/oxi)"}


def _get(url, timeout):
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read()

TYPES = ["legal", "forms", "reports", "policies", "educational",
         "correspondence", "technical", "administrative", "creative", "reference"]
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                    "pipeline_data", "docx_corpus")
MAX_MB = 4.0


def fetch(n_per_type=10, lang="en"):
    os.makedirs(ROOT, exist_ok=True)
    summary = {}
    for t in TYPES:
        outdir = os.path.join(ROOT, lang, t)
        os.makedirs(outdir, exist_ok=True)
        url = f"https://api.docxcorp.us/manifest?type={t}&lang={lang}"
        try:
            manifest = _get(url, 60).decode().split()
        except Exception as e:
            print(f"{t}: manifest FAILED {e}")
            continue
        got, tried = 0, 0
        for doc_url in manifest:
            if got >= n_per_type:
                break
            tried += 1
            if tried > n_per_type * 5:
                break
            name = doc_url.rsplit("/", 1)[-1]
            dest = os.path.join(outdir, name[:16] + ".docx")
            if os.path.exists(dest):
                got += 1
                continue
            try:
                data = _get(doc_url, 120)
            except Exception:
                continue
            if len(data) > MAX_MB * 1024 * 1024 or len(data) < 1000:
                continue
            # quarantine
            try:
                z = zipfile.ZipFile(io.BytesIO(data))
                names = z.namelist()
                if "word/document.xml" not in names:
                    continue
                if any("vbaProject" in n for n in names):
                    continue
            except Exception:
                continue
            open(dest, "wb").write(data)
            got += 1
        summary[t] = got
        print(f"{t}: {got} docs")
    json.dump(summary, open(os.path.join(ROOT, "_fetch_summary.json"), "w"))


def harness():
    import subprocess, glob, tempfile
    exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                       "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
    rows = []
    tmp = tempfile.mkdtemp(prefix="dcx_")
    for f in sorted(glob.glob(os.path.join(ROOT, "*", "*", "*.docx"))):
        rel = os.path.relpath(f, ROOT)
        try:
            r = subprocess.run([exe, f, os.path.join(tmp, "o")],
                               capture_output=True, text=True, timeout=120)
            out = (r.stdout or "") + (r.stderr or "")
            pages = None
            for line in out.splitlines():
                if line.startswith("Parsed ") and " pages" in line:
                    pages = int(line.split()[1])
                    break
            status = "ok" if (r.returncode == 0 and pages) else f"rc={r.returncode}"
            if "panicked" in out:
                status = "PANIC: " + out.split("panicked", 1)[1][:80].strip()
        except subprocess.TimeoutExpired:
            status, pages = "TIMEOUT", None
        rows.append({"doc": rel, "pages": pages, "status": status})
        flag = "" if status == "ok" else "   <-- " + status
        print(f"{rel}: pages={pages}{flag}")
    for g in glob.glob(os.path.join(tmp, "o_p*.png")):
        os.remove(g)
    json.dump(rows, open(os.path.join(ROOT, "_harness.json"), "w"), indent=1)
    bad = [r for r in rows if r["status"] != "ok"]
    print(f"\n{len(rows)} docs, {len(rows)-len(bad)} ok, {len(bad)} failed")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "fetch"
    if mode == "fetch":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 10
        lang = sys.argv[3] if len(sys.argv) > 3 else "en"
        fetch(n, lang)
    else:
        harness()
