"""
Storyline .story file — generic parser and modifier.

A .story file is a ZIP archive.  Each slide lives in story/slides/slideN.xml.
Articulate uses a custom XML schema (non-standard element boundaries), so we use
targeted regex rather than ElementTree.

Generic contract: nothing is hardcoded. No slide numbers, no GUIDs, no question
text. The parser works by pattern alone and handles any Storyline course file.

Key schema facts (stable across Storyline versions):
  <choices>
    <intrFreeChoice shpG="GUID" …> <scoringData correct="true|false" …/> …
  </choices>
        shpG links a choice entry to the answer-button shape whose g="GUID" matches.

  <pic g="GUID" … typeName="op1" …>   (or op2 / op3 / … opN)
        Outer answer-button shape. typeName="opN" is consistent across all
        Storyline MCQ/MRQ/TF interactions.

  Content shapes (question stem, option text at various states) use typeName to
  store the actual display text — so "typeName" doubles as the object label.
  We update typeName alongside <plain> and rich-text content so that the
  Storyline timeline always reflects the current text content.

  <textBox …>
    <text> &lt;Document…Span Text="answer text"…&gt; </text>  ← XML-escaped rich text
    <plain>answer text</plain>                                  ← plain-text twin
  </textBox>

  dragDropIntr   → marker that identifies DnD slides.
"""

import re
import zipfile
from pathlib import Path

# ── compiled patterns ────────────────────────────────────────────────────────

_TAG        = re.compile(r'<(?:pic|sp)\b([^>]+)>', re.IGNORECASE | re.DOTALL)
_ATTR_OP    = re.compile(r'typeName="(op\d+)"',     re.IGNORECASE)
_ATTR_G     = re.compile(r'\bg="([0-9a-f-]+)"',     re.IGNORECASE)
_CHOICES    = re.compile(r'<choices>(.*?)</choices>', re.DOTALL)
_CHOICE_ENT = re.compile(
    r'<intrFreeChoice\b[^>]*\bshpG="([0-9a-f-]+)"[^>]*/?>.*?'
    r'<scoringData\b[^>]*\bcorrect="(true|false)"',
    re.DOTALL | re.IGNORECASE
)
_PLAIN      = re.compile(r'<plain>(.*?)</plain>', re.DOTALL)
_BLOCK_ESC  = re.compile(r'&lt;Block&gt;.*?&lt;/Block&gt;', re.DOTALL)
_SLIDE_TITLE = re.compile(
    r'<slide\b[^>]*\bg="([0-9a-f-]+)"[^>]*\btitle="([^"]*)"', re.IGNORECASE
)


# ── low-level XML helpers ────────────────────────────────────────────────────

def _plain_text(raw: str) -> str:
    """Collapse a raw <plain> value to a single display string.
    Normalises all whitespace runs to a single space so the result matches
    the typeName attribute value (which stores text without extra whitespace)."""
    return re.sub(r'\s+', ' ', raw).strip()


def _xml_attr_encode(text: str) -> str:
    """Encode text for use as an XML attribute value (double-quote delimited)."""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace('\n', ' ')
                .replace('\r', ''))


def _xml_text_encode(text: str) -> str:
    """Encode text for XML text node content (& < > must be escaped)."""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;'))


def _find_pic_end(xml: str, start: int) -> int:
    """Return the position just after the </pic> that closes the <pic> at *start*.
    Handles nested <pic> elements (depth tracking) and self-closing <pic … />.
    Falls back to 40 000 chars if no matching close is found (malformed XML guard)."""
    depth = 0
    for m in re.finditer(r'<(/?)pic\b([^>]*)>', xml[start:], re.DOTALL):
        if m.group(0).endswith('/>'):  # self-closing <pic … /> — net zero depth
            pass
        elif m.group(1):              # closing </pic>
            depth -= 1
        else:                         # opening <pic>
            depth += 1
        if depth == 0:
            return start + m.end()
    return min(start + 40_000, len(xml))


def _update_tag_typename_at(xml: str, before_pos: int, new_text: str) -> str:
    """Update the typeName attribute of the last <pic|sp> tag before *before_pos*.

    Used for the question stem: the stem's typeName is often a stale object
    label that doesn't match its <plain> content, so value-matching fails.
    Walking backward to the nearest enclosing shape tag is reliable.
    """
    new_enc = _xml_attr_encode(new_text)
    search_region = xml[:before_pos]
    tag_matches = list(_TAG.finditer(search_region))
    if not tag_matches:
        return xml
    last = tag_matches[-1]
    old_tag = xml[last.start():last.end()]
    if 'typeName="' not in old_tag:
        return xml
    new_tag = re.sub(r'\btypeName="[^"]*"', f'typeName="{new_enc}"', old_tag, count=1)
    return xml[:last.start()] + new_tag + xml[last.end():]


def _rewrite_text_blocks(text_content: str, new_text: str) -> str:
    """
    Update escaped Block sequences within one <text>…</text> element's content.

    Storyline wraps long option/stem text by splitting it across multiple
    &lt;Block&gt; elements (one per visual line). When we replace text we want:
      - First Block: update the first Span's Text attribute in-place
        (all other Span attributes, Style, and nesting are preserved exactly)
      - Subsequent Blocks: removed (Storyline re-wraps visually at render time)

    This preserves FontFamily, FontSize, color, and all other Style attributes
    exactly as authored — only the text content changes.
    """
    new_enc = _xml_attr_encode(new_text)
    first_block = [False]

    def _sub_block(m):
        block_xml = m.group(0)
        if not first_block[0]:
            first_block[0] = True
            # Update first Span Text in-place; clear any subsequent Span Texts
            first_span = [False]
            def _sub_span_text(sm):
                if not first_span[0]:
                    first_span[0] = True
                    return sm.group(1) + new_enc + sm.group(3)
                return sm.group(1) + '' + sm.group(3)
            return re.sub(r'(\bText=")([^"]*?)(")', _sub_span_text, block_xml)
        return ''  # remove wrapped-line continuation blocks

    return _BLOCK_ESC.sub(_sub_block, text_content)


def _rewrite_option_block(block: str, new_text: str) -> str:
    """
    Replace display content within one answer-option block:
      - typeName attribute on content shapes (pattern-based; preserves opN labels)
      - <plain> tags (XML-encoded)
      - Rich-text within each <text> element (Span Text updated in-place;
        Style/formatting preserved exactly; wrapped continuation Blocks removed)

    Correct/incorrect flags in <scoringData> are intentionally preserved.
    """
    new_attr_enc  = _xml_attr_encode(new_text)
    new_plain_enc = _xml_text_encode(new_text)

    # 1. Update typeName on content shapes (all non-opN non-empty values)
    def _repl_typename(m):
        val = m.group(1)
        if not val or re.match(r'^op\d+$', val, re.IGNORECASE):
            return m.group(0)   # preserve typeName="" and typeName="opN"
        return f'typeName="{new_attr_enc}"'
    block = re.sub(r'typeName="([^"]*)"', _repl_typename, block)

    # 2. Update <plain> tags (XML-encode to keep document well-formed)
    block = _PLAIN.sub(f'<plain>{new_plain_enc}</plain>', block)

    # 3. Update rich-text: process each <text> element independently
    def _update_text_elem(m):
        return f'<text>{_rewrite_text_blocks(m.group(1), new_text)}</text>'
    block = re.sub(r'<text>(.*?)</text>', _update_text_elem, block, flags=re.DOTALL)

    return block


# ── slide-level parsing ──────────────────────────────────────────────────────

def _find_op_blocks(xml: str) -> list:
    """
    Locate each answer-option shape (typeName="opN") in the slide XML.
    Returns [(op_num, guid, block_start, block_end), …] sorted by op_num.

    Only the FIRST occurrence of each opN is used (outer/top-level shape).
    Inner duplicates inside <stateLst> share the same typeName but are children
    of the outer shape and are encompassed within its block.
    """
    found = {}   # op_num → (guid, tag_start)

    for m in _TAG.finditer(xml):
        attrs  = m.group(1)
        op_m   = _ATTR_OP.search(attrs)
        g_m    = _ATTR_G.search(attrs)
        if not op_m or not g_m:
            continue
        op_num = int(re.search(r'\d+', op_m.group(1)).group())
        guid   = g_m.group(1).lower()
        if op_num not in found:
            found[op_num] = (guid, m.start())

    if not found:
        return []

    by_pos    = sorted(found.items(), key=lambda kv: kv[1][1])
    positions = [v[1] for _, v in by_pos]

    result = []
    for i, (op_num, (guid, start)) in enumerate(by_pos):
        end = positions[i + 1] if i + 1 < len(positions) else _find_pic_end(xml, start)
        result.append((op_num, guid, start, end))

    result.sort(key=lambda r: r[0])
    return result


def _parse_choices(xml: str) -> dict:
    """Return {guid_lower: correct_bool} from the <choices> block."""
    cm = _CHOICES.search(xml)
    if not cm:
        return {}
    correct_map = {}
    for m in _CHOICE_ENT.finditer(cm.group(1)):
        correct_map[m.group(1).lower()] = (m.group(2).lower() == 'true')
    return correct_map


def _extract_stem(xml: str, option_texts: set, max_pos: int = None) -> str:
    """
    Return the question stem from the region xml[:max_pos] (or full xml if None).

    Heuristics applied in order:
      1. Last candidate that contains '?' (question mark) — most reliable signal
      2. Last candidate overall (in document order, last = closest to option blocks)

    Using the LAST candidate rather than the first avoids picking up slide titles
    or instruction text that precede the actual question.
    """
    search = xml[:max_pos] if max_pos is not None else xml
    candidates = []
    for m in _PLAIN.finditer(search):
        text = _plain_text(m.group(1))
        if not text or text in option_texts or len(text) < 8:
            continue
        candidates.append(text)
    if not candidates:
        return ''
    for t in reversed(candidates):
        if '?' in t:
            return t
    return candidates[-1]


def _parse_slide(xml: str, slide_file: str, title_map: dict) -> dict | None:
    """Parse one slide; return a question dict or None if not a quiz slide."""
    if '<choices>' not in xml:
        return None

    is_dnd = 'dragDropIntr' in xml

    correct_map = _parse_choices(xml)
    if not correct_map:
        return None

    op_blocks = _find_op_blocks(xml)
    if not op_blocks:
        return None

    options = []
    for op_num, guid, start, end in op_blocks:
        block = xml[start:end]
        text = ''
        for pm in _PLAIN.finditer(block):
            t = _plain_text(pm.group(1))
            if t:
                text = t
                break
        correct = correct_map.get(guid, False)
        options.append({'position': op_num, 'text': text, 'correct': correct, 'guid': guid})

    slide_guid_m = re.search(r'<(?:sld|slide)\b[^>]*\bg="([0-9a-f-]+)"', xml, re.IGNORECASE)
    slide_guid   = slide_guid_m.group(1).lower() if slide_guid_m else ''
    slide_title  = title_map.get(slide_guid, Path(slide_file).stem)

    option_texts = {o['text'] for o in options if o['text']}
    # Limit stem search to the region before the first option block in XML order
    # so that instruction text embedded between option blocks is never mistaken
    # for the stem.
    first_op_xml_pos = min(b[2] for b in op_blocks)
    stem = _extract_stem(xml, option_texts, max_pos=first_op_xml_pos)

    # Fallback: use slide title when no qualifying stem text found
    if not stem:
        stem = slide_title

    n_correct = sum(1 for o in options if o['correct'])
    if is_dnd:
        q_type = 'DND'
    elif len(options) == 2:
        q_type = 'TF'
    elif n_correct > 1:
        q_type = 'MRQ'
    else:
        q_type = 'MCQ'

    return {
        'slide_file':  slide_file,
        'slide_title': slide_title,
        'stem':        stem,
        'options':     options,
        'q_type':      q_type,
        'n_correct':   n_correct,
        'preview':     (stem[:120] + '…') if len(stem) > 120 else stem,
    }


# ── public API ───────────────────────────────────────────────────────────────

def get_story_questions(story_path: str) -> list:
    """
    Parse any Storyline .story file and return a list of question dicts.
    Each dict:
      slide_file   str   ZIP path e.g. "story/slides/slide3.xml"
      slide_title  str   slide title (from story.xml, or filename stem)
      stem         str   question text
      options      list  [{position, text, correct, guid}]
      q_type       str   "MCQ" | "MRQ" | "TF" | "DND"
      n_correct    int   number of correct options
      preview      str   truncated stem for UI
    """
    questions = []
    with zipfile.ZipFile(story_path, 'r', allowZip64=True) as z:
        names = z.namelist()

        title_map = {}
        if 'story/story.xml' in names:
            with z.open('story/story.xml') as f:
                s = f.read().decode('utf-8', errors='replace')
            for m in _SLIDE_TITLE.finditer(s):
                title_map[m.group(1).lower()] = m.group(2)

        slide_files = sorted(
            n for n in names
            if re.match(r'story/slides/slide[0-9a-fA-F]+\.xml$', n)
        )
        for sf in slide_files:
            with z.open(sf) as f:
                xml = f.read().decode('utf-8', errors='replace')
            q = _parse_slide(xml, sf, title_map)
            if q:
                questions.append(q)

    return questions


def apply_story_mappings(story_path: str, pptx_questions: list,
                         mappings: list, output_path: str) -> list:
    """
    Apply user-selected mappings to a .story file.
    Writes a NEW file to output_path — the original is NEVER modified.

    Streams one ZIP entry at a time so memory peak = largest single entry,
    not the whole archive.  Handles ZIP64 for large .story files.

    mappings  list of {story_slide_file: str, pptx_id: int}
    Returns   list of {slide_file, changes, error?}
    """
    pptx_by_id = {q['id']: q for q in pptx_questions}

    # Build lookup: slide filename → (pptx_q, result_dict)
    slides_to_modify = {}
    results = []
    for mapping in mappings:
        slide_file = mapping.get('story_slide_file', '')
        pptx_id    = int(mapping.get('pptx_id', 0))
        pptx_q     = pptx_by_id.get(pptx_id)
        r = {'slide_file': slide_file, 'changes': 0}
        if not pptx_q or not slide_file:
            r['error'] = 'slide or pptx question not found'
        else:
            slides_to_modify[slide_file] = (pptx_q, r)
        results.append(r)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Stream: read one entry → optionally modify → write → free memory
    with zipfile.ZipFile(story_path, 'r', allowZip64=True) as zin, \
         zipfile.ZipFile(output_path, 'w', allowZip64=True) as zout:

        for info in zin.infolist():
            raw = zin.read(info.filename)

            if info.filename in slides_to_modify:
                pptx_q, result = slides_to_modify[info.filename]
                try:
                    xml = raw.decode('utf-8', errors='replace')
                    updated_xml, n = _apply_to_slide(xml, pptx_q)
                    raw = updated_xml.encode('utf-8')
                    result['changes'] = n
                except Exception as e:
                    result['error'] = str(e)
                    # raw stays as original — slide is left unmodified on error

            # writestr with ZipInfo preserves metadata; Python recalculates
            # CRC and sizes from the actual data being written
            zout.writestr(info, raw)

    return results


def _apply_to_slide(xml: str, pptx_q: dict) -> tuple[str, int]:
    """
    Overwrite stem + option text (and their typeName object labels) in one
    slide's XML with pptx_q content.  Returns (updated_xml, n_changes).
    correct/incorrect flags in <scoringData> are intentionally preserved.
    """
    pptx_opts = pptx_q.get('options', [])
    pptx_stem = pptx_q.get('stem', '')
    changes   = 0

    op_blocks = _find_op_blocks(xml)
    if not op_blocks:
        return xml, 0

    # ── options ──────────────────────────────────────────────────────────────
    for i, (op_num, guid, start, end) in enumerate(op_blocks):
        if i >= len(pptx_opts):
            break
        new_text  = pptx_opts[i]
        old_block = xml[start:end]
        new_block = _rewrite_option_block(old_block, new_text)
        if new_block != old_block:
            xml = xml[:start] + new_block + xml[end:]
            shift = len(new_block) - len(old_block)
            op_blocks = [
                (n, g, s + shift if s > start else s, e + shift if e > start else e)
                for n, g, s, e in op_blocks
            ]
            changes += 1

    # ── stem ─────────────────────────────────────────────────────────────────
    if pptx_stem and op_blocks:
        # Use the XML-order first option as the prefix boundary so we never
        # search for the stem inside an option block.  op_blocks[0] is sorted
        # by op_num (not XML position), so take the minimum start instead.
        first_op_xml_pos = min(b[2] for b in op_blocks)
        prefix           = xml[:first_op_xml_pos]

        # Collect option texts from ALL op blocks (not just the last one in XML
        # order) so that op2/op3/op4 text is excluded from stem detection even
        # when those blocks precede op1 in document order.
        option_texts = set()
        for _, _, s, e in op_blocks:
            for m in _PLAIN.finditer(xml[s:e]):
                t = _plain_text(m.group(1))
                if t:
                    option_texts.add(t)
        option_texts.update(o for o in pptx_opts if o)  # also exclude new text
        replaced_stem  = False
        stem_match_pos = [None]   # position of stem <plain> in prefix

        pptx_stem_plain = _xml_text_encode(pptx_stem)

        def _replace_stem_plain(m):
            nonlocal replaced_stem, changes
            t = _plain_text(m.group(1))
            if not replaced_stem and t and t not in option_texts and len(t) >= 8:
                replaced_stem     = True
                stem_match_pos[0] = m.start()
                if t != pptx_stem:
                    changes += 1
                return f'<plain>{pptx_stem_plain}</plain>'
            return m.group(0)

        new_prefix = _PLAIN.sub(_replace_stem_plain, prefix, count=20)

        if replaced_stem and stem_match_pos[0] is not None:
            # Find the last non-empty <text>...</text> before the stem <plain>.
            # In a Storyline textBox, <text> (rich-text) precedes <plain> as
            # siblings, so this targets exactly the stem's rich-text element.
            last_text_m = None
            for m in re.finditer(r'<text>(.*?)</text>',
                                  new_prefix[:stem_match_pos[0]], re.DOTALL):
                if m.group(1).strip():
                    last_text_m = m

            if last_text_m:
                new_text_elem = (f'<text>'
                                 f'{_rewrite_text_blocks(last_text_m.group(1), pptx_stem)}'
                                 f'</text>')
                shift = len(new_text_elem) - (last_text_m.end() - last_text_m.start())
                new_prefix = (new_prefix[:last_text_m.start()] +
                              new_text_elem +
                              new_prefix[last_text_m.end():])
                # Adjust stem_match_pos for the size change that precedes it
                stem_match_pos[0] += shift

            # Update typeName on the shape enclosing the stem <plain>.
            # Position-based: the stem's typeName is often a stale object label
            # that differs from <plain> content.
            new_prefix = _update_tag_typename_at(
                new_prefix, stem_match_pos[0], pptx_stem
            )

        xml = new_prefix + xml[first_op_xml_pos:]

    return xml, changes
