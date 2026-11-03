"""PPTX + Word (SingleTable export) parsing and write-back."""
from docx import Document
from fixtures import OPTS_EN, OPTS_ES, STEM_EN, STEM_ES

import parser as sb_parser


def test_pptx_questions_strip_option_letters(pptx_path):
    qs = sb_parser.get_pptx_questions(str(pptx_path))
    assert len(qs) == 1  # single-line "Welcome" slide is skipped
    q = qs[0]
    assert q["id"] == 1 and q["stem"] == STEM_ES and q["options"] == OPTS_ES


def test_word_questions_single_table(docx_path):
    qs = sb_parser.get_word_questions(str(docx_path))
    assert len(qs) == 1  # "Intro" slide has no answer states
    q = qs[0]
    assert q["slide_id"] == "Q1" and q["type"] == "MCQ"
    assert q["stem"] == STEM_EN and q["options"] == OPTS_EN


def test_apply_mappings_fills_translation_column(docx_path, pptx_path, tmp_path):
    doc, results = sb_parser.apply_mappings(
        str(docx_path), str(pptx_path), [{"word_slide_id": "Q1", "pptx_id": 1}])
    # stem + 3 options x (normal + selected state)
    assert results == [{"slide_name": "Q1", "filled": 7, "mismatch": 0}]
    out = tmp_path / "out.docx"
    doc.save(str(out))
    rows = [[c.text for c in r.cells] for r in Document(str(out)).tables[0].rows]
    trans = {(r[1], r[2]): r[3] for r in rows}
    assert trans[("Text Normal State", STEM_EN)] == STEM_ES
    for en, es in zip(OPTS_EN, OPTS_ES):
        assert trans[("Button Normal State", en)] == es
        assert trans[("Button Selected State", en)] == es
    # Source column untouched; unrelated slide untouched.
    assert trans[("Text Normal State", "Welcome to the course")] == ""


def test_apply_mappings_ignores_unknown_ids(docx_path, pptx_path):
    _, results = sb_parser.apply_mappings(
        str(docx_path), str(pptx_path),
        [{"word_slide_id": "nope", "pptx_id": 1}, {"word_slide_id": "Q1", "pptx_id": 99}])
    assert results == []


def test_helpers():
    assert sb_parser._clean("1Hello\n\n2world") == "Hello world"
    assert sb_parser._is_skip("Submit") and sb_parser._is_skip("%question%") and sb_parser._is_skip("42")
    assert not sb_parser._is_skip("Mercury")
    left, right = sb_parser._split_proportional("uno dos tres cuatro", "one two three four", "one two")
    assert (left, right) == ("uno dos", "tres cuatro")
