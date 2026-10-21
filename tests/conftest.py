from __future__ import annotations

from pathlib import Path

import pytest

from docx_formatter.core.profile import Profile, load_profile

REPO_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def gost_profile() -> Profile:
    profile, _ = load_profile(REPO_ROOT / "profiles" / "gost_7_32_2017.json")
    return profile
