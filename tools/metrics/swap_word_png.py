# -*- coding: utf-8 -*-
"""Swap the freshly-regenerated word_png_new in as the reference set, preserving
the originals as a backup and keeping the OLD reference for any doc that failed
to regenerate (Word COM timeout on heavy docs).

  word_png        -> word_png_backup_<date>   (full backup, reversible)
  word_png_new/<d>-> word_png/<d>             (per-doc, only where a fresh render exists)
  backup/<d>      -> word_png/<d>             (restore old for docs missing a fresh render)

Run AFTER tools/metrics/regen_word_png.py finishes. Idempotent-ish (refuses if
the backup already exists).
"""
import os, sys, shutil
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(r"c:\Users\ryuji\oxi-main") / "pipeline_data"
OLD = REPO / "word_png"
NEW = REPO / "word_png_new"
BACKUP = REPO / "word_png_backup_20260627"


def main():
    if not NEW.exists():
        print("no word_png_new — run regen first"); return
    if BACKUP.exists():
        print(f"backup {BACKUP} already exists — aborting (manual check needed)"); return
    # the reference set = doc dirs that have a docx + a real page in OLD
    old_docs = {d.name for d in OLD.iterdir() if d.is_dir() and (d / "page_0001.png").exists()}
    new_docs = {d.name for d in NEW.iterdir() if d.is_dir() and (d / "page_0001.png").exists()}
    regenerated = old_docs & new_docs
    kept_old = old_docs - new_docs
    print(f"OLD ref docs: {len(old_docs)} | regenerated: {len(regenerated)} | kept-old (regen failed): {len(kept_old)}")
    if kept_old:
        print("  kept-old:", sorted(kept_old)[:20])
    # 1. full backup of OLD
    print("backing up word_png -> word_png_backup_20260627 ...")
    shutil.move(str(OLD), str(BACKUP))
    OLD.mkdir(parents=True, exist_ok=True)
    # 2. move regenerated dirs from NEW into word_png
    for d in regenerated:
        shutil.move(str(NEW / d), str(OLD / d))
    # 3. restore old dirs for failed regens
    for d in kept_old:
        shutil.copytree(str(BACKUP / d), str(OLD / d))
    # 4. also restore any extra OLD dirs (the _pN duplicates etc.) so nothing is lost
    for d in BACKUP.iterdir():
        if d.is_dir() and not (OLD / d.name).exists():
            shutil.copytree(str(d), str(OLD / d.name))
    print(f"DONE. word_png now has {sum(1 for x in OLD.iterdir() if x.is_dir())} dirs. "
          f"Backup at {BACKUP}. Leftover word_png_new dirs: "
          f"{sum(1 for x in NEW.iterdir() if x.is_dir()) if NEW.exists() else 0}")


if __name__ == "__main__":
    main()
