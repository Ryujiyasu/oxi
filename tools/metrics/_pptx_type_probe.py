# -*- coding: utf-8 -*-
"""Empirically resolve the MsoShapeType numeric values that THIS PowerPoint
installation reports for Shape.Type, so the truth harness can label shapes
without relying on a possibly-wrong hand-written enum.

For each shape it prints: Type  | HasTable HasTextFrame HasText | AutoShapeType
| PlaceholderType | and a semantic tag derived from properties (not the enum).
Run: python _pptx_type_probe.py <input.pptx>
"""
import sys

sys.stdout.reconfigure(encoding="utf-8")


def tag(sh):
    try:
        if sh.HasTable:
            return "TABLE"
    except Exception:
        pass
    try:
        if sh.PlaceholderFormat:
            return "PLACEHOLDER"
    except Exception:
        pass
    try:
        if sh.HasSmartArt:
            return "SMARTART"
    except Exception:
        pass
    try:
        if sh.HasChart:
            return "CHART"
    except Exception:
        pass
    try:
        if sh.Type == 6:
            return "GROUP"
    except Exception:
        pass
    try:
        if sh.HasTextFrame:
            return "TEXTFRAME"
    except Exception:
        pass
    return "?"


def walk(sh, depth):
    t = int(sh.Type)
    try:
        htf = bool(sh.HasTextFrame)
        ht = bool(sh.TextFrame.HasText) if htf else False
    except Exception:
        htf, ht = "?", "?"
    try:
        has_tbl = bool(sh.HasTable)
    except Exception:
        has_tbl = "?"
    try:
        ast = int(sh.AutoShapeType)
    except Exception:
        ast = "?"
    try:
        ptype = int(sh.PlaceholderFormat.Type)
    except Exception:
        ptype = "?"
    print(
        "%sType=%-3d tag=%-11s HasTable=%-5s HasTextFrame=%-6s HasText=%-5s "
        "AutoShapeType=%-4s PlaceholderType=%s  %s"
        % ("  " * depth, t, tag(sh), has_tbl, htf, ht, ast, ptype, sh.Name)
    )
    try:
        if t == 6 and sh.GroupItems.Count > 0:
            for i in range(1, int(sh.GroupItems.Count) + 1):
                walk(sh.GroupItems(i), depth + 1)
    except Exception:
        pass


def main(in_path):
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(in_path, True, False, False)
        try:
            for si in range(1, int(pres.Slides.Count) + 1):
                slide = pres.Slides(si)
                print("--- slide %d ---" % si)
                for i in range(1, int(slide.Shapes.Count) + 1):
                    walk(slide.Shapes(i), 0)
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    finally:
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main(sys.argv[1])
