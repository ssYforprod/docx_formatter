from __future__ import annotations

from docx_formatter.core.units import cm, mm, pt


def test_cm_matches_observed_twips_conversion() -> None:
    # Зафиксировано разведкой (scratchpad/explore_oxml.py): Cm(1.25) -> 709
    # твипов, не 720 — так считает python-docx (EMU-based округление), и
    # именно это число будет в w:ind/@w:firstLine для отступа ГОСТ 1.25 см.
    assert cm(1.25).twips == 709


def test_pt_and_mm_expose_docx_shared_conversions() -> None:
    assert pt(14).twips == 280
    assert mm(20).cm == 2.0
