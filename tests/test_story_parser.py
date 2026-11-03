"""Storyline .story parsing, write-back, and the unpacked-size guard."""
import zipfile

import pytest
from fixtures import OPTS_EN, OPTS_ES, STEM_EN, STEM_ES

import story_parser as sp

SLIDE = "story/slides/slide1.xml"


def test_get_story_questions(story_path):
    qs = sp.get_story_questions(str(story_path))
    assert len(qs) == 1  # slide2 has no <choices>
    q = qs[0]
    assert q["slide_file"] == SLIDE and q["slide_title"] == "Planets quiz"
    assert q["stem"] == STEM_EN and q["q_type"] == "MCQ" and q["n_correct"] == 1
    assert [o["text"] for o in q["options"]] == OPTS_EN
    assert [o["correct"] for o in q["options"]] == [True, False, False]


def test_apply_story_mappings_writes_new_file(story_path, tmp_path):
    original = story_path.read_bytes()
    pptx_q = {"id": 1, "stem": STEM_ES, "options": OPTS_ES}
    out = tmp_path / "out" / "course.story"
    results = sp.apply_story_mappings(str(story_path), [pptx_q],
                                      [{"story_slide_file": SLIDE, "pptx_id": 1}], str(out))
    assert results == [{"slide_file": SLIDE, "changes": 4}]
    assert story_path.read_bytes() == original  # input never modified

    with zipfile.ZipFile(out) as z:
        assert sorted(z.namelist()) == sorted(zipfile.ZipFile(story_path).namelist())
        assert z.read("story/media/readme.txt") == b"untouched"
        xml = z.read(SLIDE).decode()
    for en, es in zip(OPTS_EN, OPTS_ES):
        assert f"<plain>{es}</plain>" in xml and f"<plain>{en}</plain>" not in xml
        assert f'Span Text="{es}"' in xml
    assert f"<plain>{STEM_ES}</plain>" in xml and f'typeName="{STEM_ES}"' in xml
    # opN labels and correct/incorrect flags are preserved.
    assert all(f'typeName="op{i}"' in xml for i in (1, 2, 3))
    assert xml.count('correct="true"') == 1 and xml.count('correct="false"') == 2

    reparsed = sp.get_story_questions(str(out))[0]
    assert reparsed["stem"] == STEM_ES and [o["text"] for o in reparsed["options"]] == OPTS_ES


def test_apply_story_mappings_reports_missing_pptx(story_path, tmp_path):
    results = sp.apply_story_mappings(str(story_path), [],
                                      [{"story_slide_file": SLIDE, "pptx_id": 5}], str(tmp_path / "o.story"))
    assert results[0]["error"] and results[0]["changes"] == 0


def test_xml_encoding_helpers():
    assert sp._xml_attr_encode('a & "b" <c>\n') == "a &amp; &quot;b&quot; &lt;c&gt; "
    assert sp._xml_text_encode("x < y & z") == "x &lt; y &amp; z"


@pytest.mark.parametrize("fn", ["get", "apply"])
def test_zip_bomb_guard_rejects_oversized_archive(story_path, tmp_path, monkeypatch, fn):
    # Declared unpacked size of the fixture is a few KB; lower the 1 GB cap below it
    # instead of building a real 1 GB archive.
    total = sum(i.file_size for i in zipfile.ZipFile(story_path).infolist())
    monkeypatch.setattr(sp, "MAX_UNPACKED_B", total - 1)
    with pytest.raises(ValueError, match="unreasonable size"):
        if fn == "get":
            sp.get_story_questions(str(story_path))
        else:
            sp.apply_story_mappings(str(story_path), [], [], str(tmp_path / "o.story"))


def test_zip_bomb_guard_allows_at_limit(story_path, monkeypatch):
    total = sum(i.file_size for i in zipfile.ZipFile(story_path).infolist())
    monkeypatch.setattr(sp, "MAX_UNPACKED_B", total)
    assert len(sp.get_story_questions(str(story_path))) == 1


def test_default_limit_is_1_gib():
    assert sp.MAX_UNPACKED_B == 1024 ** 3
