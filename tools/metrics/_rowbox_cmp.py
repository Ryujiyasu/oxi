# -*- coding: utf-8 -*-
"""Compare Oxi row pitches (default vs OXI_ROWBOX2=1) against Word PDF truth
on the rowbox_sweep specimen (16 configs, 1 page each)."""
import os, sys, json, subprocess, tempfile

ROOT = r"c:\Users\ryuji\oxi-main"
SWEEP = os.path.join(os.environ.get("TEMP", "."), "rowbox_sweep")
DOCX = os.path.join(SWEEP, "rowbox_sweep.docx")
PDF = os.path.join(SWEEP, "rowbox_sweep.pdf")
RENDERER = os.path.join(ROOT, r"tools\oxi-gdi-renderer\target\release\oxi-gdi-renderer.exe")

CONFIGS = ["sz4_4L","sz4_1L","sz4_2L","sz0_4L","sz0_1L","sz8_4L","sz8_1L","sz12_1L",
           "sz24_1L","sz4mar_1L","sz4mar_4L","cell4_1L","cell4_4L","sz4_fs9_1L",
           "sz4_atl_1L","sz4_ex_1L"]

def merge(ys, tol=1.2):
    out = []
    for y in sorted(ys):
        if out and abs(out[-1] - y) < tol:
            continue
        out.append(y)
    return out

def pitches(ys):
    return [round(ys[i+1] - ys[i], 2) for i in range(len(ys) - 1)]

def word_pitches():
    import fitz
    d = fitz.open(PDF)
    res = {}
    for i, tag in enumerate(CONFIGS):
        if i >= len(d):
            break
        ys = set()
        for dr in d[i].get_drawings():
            for it in dr["items"]:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.2 and abs(p1.x - p2.x) > 60:
                        ys.add(round(p1.y, 2))
                elif it[0] == "re":
                    rr = it[1]
                    if rr.height < 2.0 and rr.width > 60:
                        ys.add(round(rr.y0, 2))
        res[tag] = pitches(merge(ys))
    return res

def oxi_pitches(env_extra):
    td = tempfile.mkdtemp()
    dump = os.path.join(td, "dump.json")
    env = dict(os.environ)
    env.update(env_extra)
    subprocess.run([RENDERER, DOCX, os.path.join(td, "pg"), "96",
                    "--dump-layout=" + dump, "--exclude=text,shading,box,image,clip"],
                   env=env, check=True, capture_output=True)
    data = json.load(open(dump, encoding="utf-8"))
    res = {}
    for i, tag in enumerate(CONFIGS):
        if i >= len(data["pages"]):
            break
        page = data["pages"][i]
        ys = set()
        for el in page["elements"]:
            if el["type"] == "border" and el["w"] > 60 and el["h"] < 2.0:
                ys.add(round(el["y"], 2))
        res[tag] = pitches(merge(ys))
    return res

w = word_pitches()
a = oxi_pitches({"OXI_ROWBOX2": ""})  # empty string still "is_ok"? -- no: set explicitly below
# NOTE: env var present with empty value -> std::env::var returns Ok("") -> is_ok() true.
# So default A must NOT contain the key at all.
a = oxi_pitches({})
os.environ.pop("OXI_ROWBOX2", None)
b_env = {"OXI_ROWBOX2": "1"}
b = oxi_pitches(b_env)

print(f"{'config':12s} {'Word':32s} | {'Oxi default':32s} | {'Oxi ROWBOX2':32s}")
for tag in CONFIGS:
    def fmt(p):
        return ",".join(f"{x:.2f}" for x in p[:4])
    print(f"{tag:12s} {fmt(w.get(tag, [])):32s} | {fmt(a.get(tag, [])):32s} | {fmt(b.get(tag, [])):32s}")
