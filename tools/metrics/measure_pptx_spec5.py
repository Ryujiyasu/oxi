# -*- coding: utf-8 -*-
"""Spec #5 (paragraph alignment) wave-1 oracle: open repro in PowerPoint COM,
record paragraph Alignment + shape/frame geometry, export PDF."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import win32com.client

base = r"pipeline_data\pptx_probes\spec5_align"
pptx_path = os.path.abspath(os.path.join(base, "spec5_align.pptx"))
pdf_path = os.path.abspath(os.path.join(base, "spec5_align.pdf"))

app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(pptx_path, WithWindow=False)

sh = pres.Slides(1).Shapes(1)
tf = sh.TextFrame
data = {
    "shape": {"left": sh.Left, "top": sh.Top, "width": sh.Width, "height": sh.Height},
    "frame": {
        "margin_left": tf.MarginLeft, "margin_right": tf.MarginRight,
        "margin_top": tf.MarginTop, "margin_bottom": tf.MarginBottom,
        "word_wrap": tf.WordWrap, "vertical_anchor": tf.VerticalAnchor,
        "auto_size": getattr(tf, "AutoSize", None),
    },
    "paragraphs": [],
}
tr = tf.TextRange
n = tr.Paragraphs().Count
for i in range(1, n + 1):
    p = tr.Paragraphs(i)
    data["paragraphs"].append({
        "text": p.Text,
        "alignment": p.ParagraphFormat.Alignment,
    })

with open(os.path.join(base, "spec5_truth.json"), "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

pres.SaveAs(pdf_path, 32)
pres.Close()
app.Quit()
print("truth saved, paragraphs:", n)
print("shape:", data["shape"])
