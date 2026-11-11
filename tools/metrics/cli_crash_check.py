# -*- coding: utf-8 -*-
"""Convert many documents with the CLI and report every one that fails.

  python tools/metrics/cli_crash_check.py <oxidocs.exe> [--jobs=4] [--timeout=300]
         [--part=word/media/*.emf] [--out=DIR] <docx or glob> ...

Each document is converted with `docx-to-pdf` (OXI_IGNORE_LRPB=1, as every
benchmark). A non-zero exit, a timeout or a missing PDF is a failure; the last
lines of its output are printed. --part keeps only documents whose package holds
a part matching the pattern (fnmatch, case-insensitive), e.g. the documents with
a metafile. Paths are handled in Python, not through a shell loop, so a carriage
return or a backslash in a path cannot turn every conversion into "file not
found" (the 2026-11-11 WMF check lost a whole run that way).
Exit status: 0 when nothing failed, 1 otherwise.
"""
import fnmatch, glob, os, subprocess, sys, tempfile, zipfile
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main(argv):
    jobs, timeout, parts, out, exe, pats = 4, 300, [], None, None, []
    for a in argv:
        if a.startswith("--jobs="): jobs = max(1, int(a.split("=", 1)[1]))
        elif a.startswith("--timeout="): timeout = int(a.split("=", 1)[1])
        elif a.startswith("--part="): parts.append(a.split("=", 1)[1].lower())
        elif a.startswith("--out="): out = a.split("=", 1)[1]
        elif exe is None: exe = a
        else: pats.append(a)
    if not exe or not pats:
        print(__doc__); return 2
    docs = sorted({os.path.abspath(p) for pat in pats for p in (glob.glob(pat) or [pat])})
    if parts:
        def has_part(p):
            try:
                names = zipfile.ZipFile(p).namelist()
            except Exception:
                return False
            return any(fnmatch.fnmatch(n.lower(), pat) for n in names for pat in parts)
        docs = [d for d in docs if has_part(d)]
    out = out or tempfile.mkdtemp(prefix="oxi_crash_")
    os.makedirs(out, exist_ok=True)
    env = dict(os.environ, OXI_IGNORE_LRPB="1")

    def run(d):
        pdf = os.path.join(out, os.path.splitext(os.path.basename(d))[0] + ".pdf")
        try:
            r = subprocess.run([exe, "docx-to-pdf", d, pdf], capture_output=True, timeout=timeout,
                               env=env, encoding="utf-8", errors="replace")
        except subprocess.TimeoutExpired:
            return d, f"timeout > {timeout}s", ""
        if r.returncode != 0:
            return d, f"rc={r.returncode}", (r.stdout + r.stderr)[-400:]
        if not os.path.exists(pdf) or os.path.getsize(pdf) == 0:
            return d, "no PDF", (r.stdout + r.stderr)[-400:]
        return None

    with ThreadPoolExecutor(jobs) as ex:
        fails = [f for f in ex.map(run, docs) if f]
    for d, why, tail in fails:
        print(f"FAIL {why} {d}")
        for line in tail.strip().splitlines()[-2:]:
            print("    " + line)
    print(f"docs {len(docs)} failures {len(fails)} (PDFs in {out})")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
