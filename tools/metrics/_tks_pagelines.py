# -*- coding: utf-8 -*-
# Dump per-line Y + gap for Word page(s) and Oxi page(s) side by side.
# Usage: python _tks_pagelines.py WORD:47,48 OXI:47,48,49 [KEY=VAL...]
import os, sys, json, subprocess
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
RENDERER=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX=os.path.abspath('tools/golden-test/documents/docx/tokyoshugyo_000599795.docx')
def norm(s): return s.replace("　"," ").rstrip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
def word_lines(pi):
    pg=W["pages"][pi-1]; ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=80 and g["y"]<800: ys[round(g["y"],1)].append(g)
    out=[]
    for y0 in sorted(ys):
        gg=sorted(ys[y0],key=lambda g:g["x"]); out.append((y0,norm("".join(g["char"] for g in gg))))
    return out
def render(env):
    e=dict(os.environ); e.update(env)
    subprocess.run([RENDERER,DOCX,'C:/tmp/tks_pl','96','--dump-layout=C:/tmp/tks_pl.json'],env=e,capture_output=True)
    return json.load(open('C:/tmp/tks_pl.json',encoding="utf-8"))
def oxi_lines(od,pn):
    for pg in od["pages"]:
        if pg["page"]==pn:
            rows=defaultdict(str)
            for el in pg["elements"]:
                if el.get("type")=="text" and el.get("text","").strip() and round(el["y"],1)>=80:
                    rows[round(el["y"],1)]+=el["text"]
            return [(y,norm(rows[y])) for y in sorted(rows)]
    return []
def show(lines,tag):
    prev=None
    for y,t in lines:
        gap = f"{y-prev:5.2f}" if prev is not None else "  -  "
        print(f"  {tag} y={y:7.2f} gap={gap}  {t[:34]}")
        prev=y
wp=[int(a) for a in next((x.split(":")[1] for x in sys.argv[1:] if x.startswith("WORD:")),"").split(",") if a]
op=[int(a) for a in next((x.split(":")[1] for x in sys.argv[1:] if x.startswith("OXI:")),"").split(",") if a]
env={}
for a in sys.argv[1:]:
    if "=" in a and ":" not in a: k,v=a.split("=",1); env[k]=v
for p in wp:
    print(f"--- WORD page {p} ---"); show(word_lines(p),"W")
od=render(env)
for p in op:
    print(f"--- OXI page {p} (env={env}) ---"); show(oxi_lines(od,p),"O")
