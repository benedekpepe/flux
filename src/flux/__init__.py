"""
Flux - management reporting engine.

Turns a general ledger into a management-ready P&L with budget and prior-year
variance analysis, favourable/unfavourable classification, materiality flagging
and consolidation by legal entity, department and cost centre.

Layers:
    coa             chart of accounts, org hierarchy, P&L structure
    engine          variance computation and materiality
    commentary      narrative generation from the engine output
    ingest          arbitrary client files -> the internal schema
    synthetic_data  seeded generator for the demo company
    reporting       Excel packs (live formulas), styling and formula helpers
"""

from __future__ import annotations

__version__ = "0.5.0"
__author__ = "Benedek Péter - B.P. Studio"

from .coa import CHART_OF_ACCOUNTS, PNL_STRUCTURE
from .engine import (
    MaterialityRule,
    build_report,
    leaf_variances,
    department_variances,
    cost_centre_variances,
    entity_variances,
)
from .commentary import generate_commentary, line_comments

__all__ = [
    "CHART_OF_ACCOUNTS",
    "PNL_STRUCTURE",
    "MaterialityRule",
    "build_report",
    "leaf_variances",
    "department_variances",
    "cost_centre_variances",
    "entity_variances",
    "generate_commentary",
    "line_comments",
    "__version__",
]
