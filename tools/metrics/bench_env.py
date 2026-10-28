# -*- coding: utf-8 -*-
"""Benchmark mode for Word's saved page-break hint (<w:lastRenderedPageBreak/>).

Word writes that marker wherever its own render started a page, so a file
Word saved carries part of the answer every Oxi benchmark is scored against.
The Oxi parser reads it by default (several layout rules follow it); other
engines get no such hint. A benchmark against Word therefore runs BLIND by
default: OXI_IGNORE_LRPB=1 makes the parser drop the marker, and child
processes (the renderers) inherit it.

    import bench_env
    mode = bench_env.apply()          # "blind" unless OXI_BENCH_WITH_LRPB=1

Set OXI_BENCH_WITH_LRPB=1 to measure the hinted mode, so both numbers can be
published side by side. Renderers built before OXI_IGNORE_LRPB existed ignore
the variable and always run hinted.
"""
from __future__ import annotations

import os

IGNORE = "OXI_IGNORE_LRPB"
WITH = "OXI_BENCH_WITH_LRPB"


def mode() -> str:
    return "hinted" if os.environ.get(WITH, "") not in ("", "0") else "blind"


def apply() -> str:
    """Set the parser switch for this process and its children; return the mode."""
    m = mode()
    if m == "blind":
        os.environ[IGNORE] = "1"
    else:
        os.environ.pop(IGNORE, None)
    print(f"[bench_env] lastRenderedPageBreak: {m} "
          f"({IGNORE}={os.environ.get(IGNORE, 'unset')})", flush=True)
    return m
