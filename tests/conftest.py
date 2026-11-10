import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))


@pytest.fixture(autouse=True)
def isolated_user_profiles(tmp_path, monkeypatch):
    """Профілі користувача й стан GUI — у тимчасовій папці, а не в домашній."""
    monkeypatch.setenv("DOCX_VALIDATOR_PROFILES", str(tmp_path / "user_profiles"))
    yield


@pytest.fixture
def settings():
    from lpnu_docx_validator.settings import default_settings
    return default_settings()


@pytest.fixture
def check(tmp_path, settings):
    """check(builder_or_path, **overrides) → ValidationResult."""
    from lpnu_docx_validator.checks import validate

    def run(doc, s=None):
        path = doc if isinstance(doc, Path) else doc.save(tmp_path / "doc.docx")
        return validate(path, s or settings)
    return run


def rules_of(result, sev=None):
    return [i.rule for i in result.issues if sev is None or i.severity == sev]
