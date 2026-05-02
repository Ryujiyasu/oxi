"""Ra2 §19.8 baseline impact survey: count baseline docs where any <w:tblPr>
has <w:tblpPr> appearing before <w:tblStyle> (ECMA-376 CT_TblPrBase order
violation that Word silently drops).

Per spec §19.8: when this order is violated, Word ignores tblpPr entirely
and the table renders inline at anchor_bottom (slope=0). The TODO was to
count baseline docs affected.

Approach: parse each .docx as a zip, read document.xml, find every
<w:tblPr> block, check whether tblpPr precedes tblStyle within the
block. Report counts.
"""
import os
import sys
import zipfile
import re
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"


# Match a <w:tblPr ...>...</w:tblPr> block (non-greedy, allow attributes).
TBLPR_RE = re.compile(rb"<w:tblPr(?:\s[^>]*)?>(.*?)</w:tblPr>", re.DOTALL)
# Self-closing tblPr (rare; no children, can't violate order)
TBLPR_SELF = re.compile(rb"<w:tblPr(?:\s[^>]*)?/>")
# Inside a tblPr block, find positions of these elements (self-closing or open).
TBLSTYLE_RE = re.compile(rb"<w:tblStyle\b")
TBLPPR_RE   = re.compile(rb"<w:tblpPr\b")


def survey_doc(path):
    """Return (n_tblpr, n_with_both, n_violations) for one docx."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            try:
                xml = zf.read("word/document.xml")
            except KeyError:
                return None
    except (zipfile.BadZipFile, OSError) as e:
        return ("err", str(e))

    n_tblpr = 0
    n_with_both = 0
    n_violations = 0
    violations = []
    for m in TBLPR_RE.finditer(xml):
        n_tblpr += 1
        body = m.group(1)
        ts = TBLSTYLE_RE.search(body)
        tp = TBLPPR_RE.search(body)
        if ts and tp:
            n_with_both += 1
            if tp.start() < ts.start():
                n_violations += 1
                # Capture the snippet for inspection (first 200 chars)
                snippet = body[:200].decode("utf-8", errors="replace")
                violations.append(snippet)
    return (n_tblpr, n_with_both, n_violations, violations)


def main():
    paths = sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx")))
    print(f"Surveying {len(paths)} baseline docx files...")
    print()

    total_tblpr = 0
    total_both = 0
    total_violations = 0
    docs_with_violation = []
    err_count = 0

    for path in paths:
        name = os.path.basename(path)
        result = survey_doc(path)
        if result is None:
            continue
        if isinstance(result[0], str) and result[0] == "err":
            err_count += 1
            continue
        n_tblpr, n_both, n_viol, _viols = result
        total_tblpr += n_tblpr
        total_both += n_both
        total_violations += n_viol
        if n_viol > 0:
            docs_with_violation.append((name, n_viol, n_tblpr))

    print(f"Total <w:tblPr> blocks examined: {total_tblpr}")
    print(f"  with both tblStyle and tblpPr: {total_both}")
    print(f"  with tblpPr BEFORE tblStyle (Word silently drops): {total_violations}")
    print(f"Read errors: {err_count}")
    print()
    print(f"Docs with at least one order violation: {len(docs_with_violation)}/{len(paths)}")
    if docs_with_violation:
        for name, n_viol, n_tblpr in docs_with_violation[:20]:
            print(f"  {name}: {n_viol}/{n_tblpr} tblPr block(s) violate order")
    else:
        print("  (none — preventive fix only)")


if __name__ == "__main__":
    main()
