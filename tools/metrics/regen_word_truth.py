# -*- coding: utf-8 -*-
"""Re-measure the pagination gate's Word truth on THIS machine.

2026-10-16 decision: the gate truth is this machine's Word (see the memory note
truth_machine_is_new_pc).  The truth files live in subset_gate.TRUTH_DIRS; each
is re-measured in place with measure_pagination_word.measure_doc after:

  1. a one-time backup of every truth directory to <dir>_oldpc_20261002/
  2. a PRE-OPEN pass: every document is opened and closed once, so Word fetches
     the Office cloud fonts it names before anything is measured (Word downloads
     a cloud font when a document first asks for it, and an export made in that
     same open draws the text in a fallback)
  3. the measurement, stamped {"machine": "newpc", "word_build": ...}

The CloudFonts listing before and after the pre-open pass and Word's language
settings are written to pipeline_data/truth_env_<date>.json.  Resumable: a truth
file already stamped newpc is skipped.

    python tools/metrics/regen_word_truth.py            # all gate documents
    python tools/metrics/regen_word_truth.py --limit=5  # first five (smoke test)
"""
import datetime, json, os, shutil, sys, time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import feature_census as FC  # noqa: E402
import subset_gate as SG  # noqa: E402
import measure_pagination_word as MW  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
TAG = "newpc"


def local_path(p):
    p = str(p)
    for old in ("oxi-main",):
        p = p.replace("\\" + old + "\\", "\\oxi\\").replace("/" + old + "/", "/oxi/")
    return p


def work_list():
    rows = FC.load(); out = []
    for did, r in rows.items():
        if "error" in r:
            continue
        t = SG.truth_for(did)
        if t is None:
            continue
        p = local_path(r["path"])
        if os.path.exists(p):
            out.append((did, p, Path(t)))
    return out


def cloud_listing():
    d = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "FontCache" / "4" / "CloudFonts"
    return sorted(str(p.relative_to(d)) for p in d.rglob("*") if p.is_file()) if d.exists() else []


def backup(dirs):
    for d in dirs:
        b = d.parent / (d.name + "_oldpc_20261002")
        if d.exists() and not b.exists():
            shutil.copytree(d, b)
            print("backup", d, "->", b.name, flush=True)


def is_new(t):
    try:
        return json.loads(t.read_text(encoding="utf-8")).get("machine") == TAG
    except Exception:
        return False


def main():
    limit = None
    for a in sys.argv[1:]:
        if a.startswith("--limit="):
            limit = int(a.split("=", 1)[1])
    items = work_list()
    print("gate documents with truth:", len(items), flush=True)
    backup(sorted({t.parent for _, _, t in items}))
    todo = [it for it in items if not is_new(it[2])]
    if limit:
        todo = todo[:limit]
    print("to measure:", len(todo), flush=True)
    import win32com.client
    env = {"date": datetime.datetime.now().isoformat(timespec="seconds"), "cloud_before": cloud_listing()}
    word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
    try:
        ls = word.Application.LanguageSettings
        env["word_build"] = f"{word.Version}.{word.Build}"
        env["editing"] = {k: bool(ls.LanguagePreferredForEditing(v)) for k, v in (("ja-JP", 1041), ("en-US", 1033), ("zh-CN", 2052), ("ko-KR", 1042))}
        env["ui_lang"] = ls.LanguageID(2)
        # pre-open pass: let Word fetch the cloud fonts the documents name
        for i, (did, p, t) in enumerate(todo):
            try:
                d = word.Documents.Open(p, ReadOnly=True, AddToRecentFiles=False); d.Close(0)
            except Exception as e:
                print("  preopen FAIL", did, e, flush=True)
            if i % 50 == 0:
                print(f"  preopen {i}/{len(todo)}", flush=True)
        env["cloud_after_preopen"] = cloud_listing()
        (REPO / "pipeline_data" / f"truth_env_{datetime.datetime.now():%Y%m%d_%H%M}.json").write_text(json.dumps(env, ensure_ascii=False, indent=1), encoding="utf-8")
        print("cloud fonts added by pre-open:", sorted(set(env["cloud_after_preopen"]) - set(env["cloud_before"])), flush=True)
        for i, (did, p, t) in enumerate(todo):
            t0 = time.time()
            try:
                res = MW.measure_doc(word, p)
            except Exception as e:
                print(f"  FAIL {did}: {e}", flush=True)
                try:
                    word.Quit()
                except Exception:
                    pass
                word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
                continue
            res["machine"] = TAG; res["word_build"] = env["word_build"]
            t.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  [{i + 1}/{len(todo)}] {did} {res['n_pages']}p {time.time() - t0:.1f}s", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
