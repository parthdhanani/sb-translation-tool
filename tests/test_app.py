"""Flask test-client coverage of both workflows and input validation."""
import io
import zipfile

from docx import Document
from fixtures import STEM_ES

import story_parser as sp


def _files(**paths):
    return {k: (io.BytesIO(p.read_bytes()), p.name) for k, p in paths.items()}


def test_index(client):
    r = client.get("/")
    assert r.status_code == 200 and b"<html" in r.data


def test_word_workflow_end_to_end(client, pptx_path, docx_path, app_module):
    r = client.post("/scan", data=_files(pptx=pptx_path, word=docx_path), content_type="multipart/form-data")
    assert r.status_code == 200, r.get_json()
    d = r.get_json()
    assert len(d["pptx_questions"]) == 1 and len(d["word_questions"]) == 1
    sid = d["session_id"]

    r = client.post("/apply", json={"session_id": sid, "mappings": [{"word_slide_id": "Q1", "pptx_id": 1}]})
    assert r.status_code == 200 and r.get_json()["results"][0]["filled"] == 7

    r = client.get(f"/download/{sid}")
    assert r.status_code == 200
    assert "updated_translation.docx" in r.headers["Content-Disposition"]
    cells = [c.text for row in Document(io.BytesIO(r.data)).tables[0].rows for c in row.cells]
    assert STEM_ES in cells


def test_story_workflow_end_to_end(client, pptx_path, story_path):
    r = client.post("/scan-story", data=_files(pptx=pptx_path, story=story_path),
                    content_type="multipart/form-data")
    assert r.status_code == 200, r.get_json()
    d = r.get_json()
    assert d["story_questions"][0]["slide_title"] == "Planets quiz"
    sid = d["session_id"]

    r = client.post("/apply-story", json={"session_id": sid, "mappings": [
        {"story_slide_file": "story/slides/slide1.xml", "pptx_id": 1}]})
    assert r.status_code == 200 and r.get_json()["results"][0]["changes"] == 4

    r = client.get(f"/download-story/{sid}")
    assert r.status_code == 200
    with zipfile.ZipFile(io.BytesIO(r.data)) as z:
        assert STEM_ES in z.read("story/slides/slide1.xml").decode()


def test_missing_files_rejected(client, pptx_path):
    r = client.post("/scan", data=_files(pptx=pptx_path), content_type="multipart/form-data")
    assert r.status_code == 400
    r = client.post("/scan-story", data={}, content_type="multipart/form-data")
    assert r.status_code == 400


def test_non_zip_upload_rejected_and_cleaned(client, pptx_path, tmp_path, app_module):
    bogus = tmp_path / "notes.docx"
    bogus.write_text("not a zip")
    r = client.post("/scan", data=_files(pptx=pptx_path, word=bogus), content_type="multipart/form-data")
    assert r.status_code == 400 and "valid file" in r.get_json()["error"]
    assert list(app_module.SESSIONS.iterdir()) == []


def test_oversized_story_archive_rejected(client, pptx_path, story_path, app_module, monkeypatch):
    monkeypatch.setattr(sp, "MAX_UNPACKED_B", 100)
    r = client.post("/scan-story", data=_files(pptx=pptx_path, story=story_path),
                    content_type="multipart/form-data")
    assert r.status_code == 500 and "unreasonable size" in r.get_json()["error"]
    assert list(app_module.SESSIONS.iterdir()) == []


def test_session_id_validation(client):
    assert client.post("/apply", json={"session_id": "../../etc"}).status_code == 400
    assert client.post("/apply-story", json={"session_id": None}).status_code == 400
    assert client.get("/download/not-a-uuid").status_code == 404
    assert client.get("/download-story/..%2F..%2Fetc").status_code == 404
    unknown = "12345678-1234-1234-1234-123456789abc"
    assert client.post("/apply", json={"session_id": unknown}).status_code == 404
    assert client.get(f"/download/{unknown}").status_code == 404


def test_apply_without_json_body(client):
    for url in ("/apply", "/apply-story"):
        r = client.post(url, data="garbage", content_type="text/plain")
        assert r.status_code == 400 and r.get_json()["error"] == "Invalid session."


def test_busy_when_session_cap_reached(client, pptx_path, docx_path, app_module, monkeypatch):
    monkeypatch.setattr(app_module, "MAX_SESSIONS", 1)
    (app_module.SESSIONS / "live").mkdir()
    r = client.post("/scan", data=_files(pptx=pptx_path, word=docx_path), content_type="multipart/form-data")
    assert r.status_code == 503


def test_apply_with_non_object_json(client):
    for body in ("null", "[]", '"x"'):
        r = client.post("/apply", data=body, content_type="application/json")
        assert r.status_code == 400
