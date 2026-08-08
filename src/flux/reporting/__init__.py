"""
Excel reporting layer.

    styling       palette, type, number formats, sheet chrome
    formulas      Excel formula builders, the lever cells and the column layout
    rows          the report row every sheet writes the same way
    demo_pack     the full multi-entity showcase, from generated data
    client_pack   the pack built from an ingested client file
    pptx_pack     the management deck: the same numbers, seven slides

Both packs write live formulas over an input sheet rather than static values, so
an edited input recalculates the whole workbook. Every reporting sheet in either
pack carries the same columns, built by `rows` from the layout in `formulas`, so
the two packs cannot report the same figures two different ways.
"""

from __future__ import annotations

from .demo_pack import (build_demo_pack, demo_analysis_blocks,
                        demo_ytd_detail)
from .client_pack import build_client_pack
from .pptx_pack import build_pptx_pack

__all__ = ["build_demo_pack", "build_client_pack", "build_pptx_pack",
           "demo_analysis_blocks", "demo_ytd_detail"]
