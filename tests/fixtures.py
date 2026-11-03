"""Build tiny synthetic inputs at test time (no client content, no binaries in git).

- make_pptx():  translated storyboard, one MCQ slide + one non-question slide
- make_docx():  SingleTable-format Storyline translation export for the same MCQ
- make_story(): minimal .story ZIP with one MCQ slide (3 choices) + one content slide
"""
import zipfile
from pathlib import Path

from docx import Document
from pptx import Presentation
from pptx.util import Inches

STEM_EN = "Which planet is closest to the sun?"
OPTS_EN = ["Mercury", "Venus", "Mars"]
STEM_ES = "¿Qué planeta está más cerca del sol?"
OPTS_ES = ["Mercurio", "Venus (ES)", "Marte"]

SLIDE_GUID = "11111111-1111-1111-1111-111111111111"
OPT_GUIDS = [
    "aaaaaaa1-0000-0000-0000-000000000001",
    "aaaaaaa2-0000-0000-0000-000000000002",
    "aaaaaaa3-0000-0000-0000-000000000003",
]


def make_pptx(path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s1 = prs.slides.add_slide(blank)
    tb = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(3)).text_frame
    tb.text = STEM_ES
    for letter, opt in zip("ABC", OPTS_ES):
        tb.add_paragraph().text = f"{letter}. {opt}"
    s2 = prs.slides.add_slide(blank)
    s2.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = "Welcome"
    prs.save(str(path))
    return path


def make_docx(path: Path) -> Path:
    doc = Document()
    rows = [("ID \U0001f512", "Type", "Source", "Translation"),
            ("", "Scene name", "Scene 1", ""),
            ("", "Slide name", "Q1", ""),
            ("t1", "Text Normal State", STEM_EN, "")]
    for i, opt in enumerate(OPTS_EN, start=2):
        rows.append((f"b{i}", "Button Normal State", opt, ""))
        rows.append((f"b{i}", "Button Selected State", opt, ""))
    # A non-question slide: no answer-state rows, so it must be ignored.
    rows += [("", "Slide name", "Intro", ""), ("t9", "Text Normal State", "Welcome to the course", "")]
    table = doc.add_table(rows=len(rows), cols=4)
    for r, vals in zip(table.rows, rows):
        for c, v in zip(r.cells, vals):
            c.text = v
    doc.save(str(path))
    return path


def _textbox(text: str) -> str:
    return ("<textBox><text>&lt;Document&gt;&lt;Block&gt;"
            f"&lt;Span Text=\"{text}\" Style=\"FontFamily:Arial\"/&gt;"
            "&lt;/Block&gt;&lt;/Document&gt;</text>"
            f"<plain>{text}</plain></textBox>")


def quiz_slide_xml() -> str:
    opts = "".join(
        f'<pic g="{g}" typeName="op{i}">{_textbox(t)}</pic>'
        for i, (g, t) in enumerate(zip(OPT_GUIDS, OPTS_EN), start=1)
    )
    choices = "".join(
        f'<intrFreeChoice shpG="{g}"><scoringData correct="{"true" if i == 0 else "false"}"/></intrFreeChoice>'
        for i, g in enumerate(OPT_GUIDS)
    )
    return (f'<sld g="{SLIDE_GUID}"><shapeLst>'
            f'<sp g="bbbbbbbb-0000-0000-0000-000000000001" typeName="Question label">{_textbox(STEM_EN)}</sp>'
            f'{opts}</shapeLst><interactions><choices>{choices}</choices></interactions></sld>')


def make_story(path: Path) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("story/story.xml", f'<story><slide g="{SLIDE_GUID}" title="Planets quiz"/></story>')
        z.writestr("story/slides/slide1.xml", quiz_slide_xml())
        z.writestr("story/slides/slide2.xml",
                   '<sld g="22222222-2222-2222-2222-222222222222"><shapeLst>'
                   f'<sp g="cccccccc-0000-0000-0000-000000000001" typeName="Intro">{_textbox("Welcome to the course")}</sp>'
                   '</shapeLst></sld>')
        z.writestr("story/media/readme.txt", "untouched")
    return path
