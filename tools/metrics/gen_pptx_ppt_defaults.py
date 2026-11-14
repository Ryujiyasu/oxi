# -*- coding: utf-8 -*-
"""Author a probe deck with PowerPoint's OWN defaults, then export its PDF.

Everything on these slides is made the way a user makes it in the UI, so the
XML is PowerPoint's: inserted shapes carry `p:style` (fillRef/lnRef/fontRef),
theme colours with a brightness carry `lumMod`/`lumOff`, a new table carries
the default `tableStyleId`, connectors are `p:cxnSp`, an overfull body
placeholder gets `normAutofit`, and the title layout's subtitle inherits its
bullets from the master.

NEVER run this while the renderer is producing PNGs
(`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/gen_pptx_ppt_defaults.py
Writes pipeline_data/pptx_probes/ppt_defaults/probe_ppt_defaults.{pptx,pdf}
"""
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "ppt_defaults"
OUT.mkdir(parents=True, exist_ok=True)
PPTX = OUT / "probe_ppt_defaults.pptx"
PDF = OUT / "probe_ppt_defaults.pdf"

# Office constants
ppLayoutTitle, ppLayoutText, ppLayoutBlank = 1, 2, 12
msoShapeRectangle, msoShapeOval, msoShapeRoundedRectangle = 1, 9, 5
msoConnectorStraight, msoConnectorElbow = 1, 2
msoArrowheadTriangle = 2
msoThemeColorText1, msoThemeColorBackground1, msoThemeColorAccent1 = 13, 14, 5
msoThemeColorAccent2 = 6
msoAutoSizeTextToFitShape = 2

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    # A window is needed for PowerPoint to lay the text out, which is when it
    # computes (and later saves) normAutofit/@fontScale.
    pres = app.Presentations.Add(WithWindow=True)
    try:
        # --- slide 1: inserted shapes, theme colours with brightness, connectors
        s = pres.Slides.Add(1, ppLayoutBlank)
        r = s.Shapes.AddShape(msoShapeRectangle, 40, 40, 200, 100)
        r.TextFrame.TextRange.Text = "Default rectangle"
        o = s.Shapes.AddShape(msoShapeOval, 300, 40, 160, 100)
        o.TextFrame.TextRange.Text = "Oval"
        rr = s.Shapes.AddShape(msoShapeRoundedRectangle, 520, 40, 160, 100)
        rr.TextFrame.TextRange.Text = "Rounded"

        def tb(x, y, text, theme, bright):
            t = s.Shapes.AddTextbox(1, x, y, 300, 30)
            t.TextFrame.TextRange.Text = text
            f = t.TextFrame2.TextRange.Font.Fill.ForeColor
            f.ObjectThemeColor = theme
            f.Brightness = bright
            t.TextFrame.TextRange.Font.Size = 20
            return t

        tb(40, 170, "Text 1, lighter 50%", msoThemeColorText1, 0.5)
        tb(40, 205, "Background 1, darker 50%", msoThemeColorBackground1, -0.5)
        tb(40, 240, "Accent 1, darker 50%", msoThemeColorAccent1, -0.5)
        tb(40, 275, "Accent 2, lighter 40%", msoThemeColorAccent2, 0.4)

        # A shape filled with a lighter accent (lumMod/lumOff on the FILL).
        lf = s.Shapes.AddShape(msoShapeRectangle, 400, 180, 140, 60)
        lf.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        lf.Fill.ForeColor.Brightness = 0.8
        lf.TextFrame.TextRange.Text = "Light fill"

        a = s.Shapes.AddShape(msoShapeRectangle, 400, 300, 100, 50)
        b = s.Shapes.AddShape(msoShapeRectangle, 580, 380, 100, 50)
        c = s.Shapes.AddConnector(msoConnectorElbow, 0, 0, 10, 10)
        c.ConnectorFormat.BeginConnect(a, 4)
        c.ConnectorFormat.EndConnect(b, 1)
        c.Line.EndArrowheadStyle = msoArrowheadTriangle
        c.RerouteConnections()
        st = s.Shapes.AddConnector(msoConnectorStraight, 40, 340, 300, 400)
        st.Line.EndArrowheadStyle = msoArrowheadTriangle

        # --- slide 2: a default table
        s2 = pres.Slides.Add(2, ppLayoutBlank)
        tbl = s2.Shapes.AddTable(4, 3, 60, 60, 600, 200).Table
        for ri in range(1, 5):
            for ci in range(1, 4):
                tbl.Cell(ri, ci).Shape.TextFrame.TextRange.Text = f"R{ri}C{ci}"

        # --- slide 3: title slide (subtitle must not grow a bullet)
        s3 = pres.Slides.Add(3, ppLayoutTitle)
        s3.Shapes(1).TextFrame.TextRange.Text = "Probe title"
        s3.Shapes(2).TextFrame.TextRange.Text = "Subtitle without bullet"

        # --- slide 4: an overfull body that PowerPoint shrinks to fit
        s4 = pres.Slides.Add(4, ppLayoutText)
        s4.Shapes(1).TextFrame.TextRange.Text = "Autofit body"
        body = s4.Shapes(2)
        body.TextFrame.TextRange.Text = "\r".join(
            f"Bullet line number {i} with some extra words to wrap" for i in range(1, 15)
        )
        body.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
        pres.Windows(1).View.GotoSlide(4)
        app.ActiveWindow.ViewType = 1  # ppViewSlide forces a layout pass
        import time
        time.sleep(2)

        pres.SaveAs(str(PPTX))
        pres.SaveAs(str(PDF), 32)
        print("wrote", PPTX, PDF)
    finally:
        pres.Close()
finally:
    app.Quit()
