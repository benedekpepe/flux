"""
Excel reporting layer.

    styling       palette, type, number formats, sheet chrome
    formulas      Excel formula builders and the materiality levers
    demo_pack     the full multi-entity showcase, from generated data
    client_pack   the pack built from an ingested client file

Both packs write live formulas over an input sheet rather than static values, so
an edited input recalculates the whole workbook.
"""

from __future__ import annotations

from .demo_pack import build_demo_pack
from .client_pack import build_client_pack

__all__ = ["build_demo_pack", "build_client_pack"]
