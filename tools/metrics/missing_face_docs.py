# -*- coding: utf-8 -*-
"""Which gate documents name a face that is missing on this machine?

Word caches a missing face's substitute per process (memory note
word_missing_face_substitute_process_state): such a document's truth depends on
what the measuring Word opened before it.  List them, with each missing face's
fontTable declaration (charset / altName), so their truth can be re-measured in a
fresh Word process.

    python tools/metrics/missing_face_docs.py            # prints a summary, writes pipeline_data/missing_face_docs.json
"""
import collections, json, os, re, sys, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")
import face_census as F  # noqa: E402
import regen_word_truth as R  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
RFONTS = re.compile(r'<w:rFonts [^>]*>')
ATTR = re.compile(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"')


def faces_named(z):
    out = set()
    for n in z.namelist():
        if n.startswith("word/") and n.endswith(".xml") and "fontTable" not in n:
            try:
                for tag in RFONTS.findall(z.read(n).decode("utf-8", "replace")):
                    out |= set(ATTR.findall(tag))
            except Exception:
                pass
    return out


def declarations(z):
    try:
        ft = z.read("word/fontTable.xml").decode("utf-8", "replace")
    except KeyError:
        return {}
    out = {}
    for m in re.finditer(r'<w:font w:name="([^"]+)"([\s\S]*?)</w:font>', ft):
        cs = re.search(r'<w:charset w:val="([0-9A-Fa-f]+)"', m.group(2))
        alt = re.search(r'<w:altName w:val="([^"]+)"', m.group(2))
        out[m.group(1)] = {"charset": cs.group(1) if cs else None, "alt": alt.group(1) if alt else None}
    return out


def main():
    names, _ = F.build_inventory()
    have = set(names)
    res = {}
    by_face = collections.Counter()
    for did, path, truth in R.work_list():
        try:
            z = zipfile.ZipFile(path)
        except Exception:
            continue
        missing = sorted(f for f in faces_named(z) if f.strip().lower() not in have and not f.startswith("+"))
        if missing:
            dec = declarations(z)
            res[did] = {f: dec.get(f) for f in missing}
            for f in missing:
                by_face[f] += 1
    (REPO / "pipeline_data" / "missing_face_docs.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    print("gate documents naming a missing face:", len(res))
    for f, n in by_face.most_common(40):
        print(f"  {n:4d}  {f}")


if __name__ == "__main__":
    main()
