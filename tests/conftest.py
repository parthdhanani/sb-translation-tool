import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import fixtures  # noqa: E402


@pytest.fixture
def pptx_path(tmp_path):
    return fixtures.make_pptx(tmp_path / "storyboard_es.pptx")


@pytest.fixture
def docx_path(tmp_path):
    return fixtures.make_docx(tmp_path / "translation.docx")


@pytest.fixture
def story_path(tmp_path):
    return fixtures.make_story(tmp_path / "course.story")


@pytest.fixture
def app_module(tmp_path, monkeypatch):
    import app as app_mod

    sessions = tmp_path / "sessions"
    sessions.mkdir()
    monkeypatch.setattr(app_mod, "SESSIONS", sessions)
    app_mod.app.config["TESTING"] = True
    return app_mod


@pytest.fixture
def client(app_module):
    return app_module.app.test_client()
