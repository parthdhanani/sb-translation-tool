import re
import difflib
from docx import Document
from pptx import Presentation


# ── helpers ───────────────────────────────────────────────────────────────────

def _clean(text):
    """Remove Storyline line-number prefixes and collapse segment splits."""
    text = re.sub(r'\n\n\d', ' ', text)
    text = re.sub(r'^\d', '', text)
    return text.strip()


def _set_cell(cell, text, ref_run=None):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ''
    if not text:
        return
    para = cell.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
    else:
        run = para.add_run(text)
        if ref_run:
            # Copy font from source column so new runs don't revert to defaults
            if ref_run.font.name:
                run.font.name = ref_run.font.name
            if ref_run.font.size:
                run.font.size = ref_run.font.size
            if ref_run.font.bold is not None:
                run.font.bold = ref_run.font.bold
            if ref_run.font.italic is not None:
                run.font.italic = ref_run.font.italic


def _split_proportional(new_text, orig_full, orig_seg1):
    """Split new_text at a word boundary proportional to orig_seg1/orig_full."""
    if not orig_full or not new_text:
        return new_text, ''
    ratio = len(orig_seg1) / max(len(orig_full), 1)
    pos   = int(len(new_text) * ratio)
    left  = new_text.rfind(' ', 0, pos)
    right = new_text.find(' ', pos)
    if left == -1 and right == -1:
        return new_text, ''
    sp = left if left != -1 and (right == -1 or (pos - left) <= (right - pos)) else right
    return new_text[:sp].strip(), new_text[sp:].strip()


# ── question type detection ───────────────────────────────────────────────────

def _detect_type(data_rows):
    types = [r[1] for r in data_rows if len(r) > 1]
    has_drop = any('Drop Correct' in t or 'Drop Incorrect' in t for t in types)
    has_crt  = any('crt state'     in t for t in types)
    has_sel  = any('Selected state' in t for t in types)
    if has_drop:
        return 'DND'
    if has_crt and has_sel:
        sel_texts = {_clean(r[2]) for r in data_rows if 'Selected state' in r[1] and len(r) > 2}
        return 'TF' if len(sel_texts) <= 2 else 'MCQ'
    return None


def _stem_and_options(data_rows, q_type=None):
    if q_type == 'DND':
        item_texts = {
            _clean(r[2])
            for r in data_rows
            if ('Drop Correct' in r[1] or 'Drop Incorrect' in r[1]) and len(r) > 2
        }
    else:
        item_texts = {_clean(r[2]) for r in data_rows if 'Selected state' in r[1] and len(r) > 2}

    options_seen, options_set = [], set()
    stem = None

    for row in data_rows:
        if len(row) < 3 or row[1] == 'Slide name':
            continue
        if 'Normal state' not in row[1]:
            continue
        cleaned = _clean(row[2])
        if not cleaned or cleaned == 'Submit':
            continue
        if cleaned in item_texts:
            if cleaned not in options_set:
                options_seen.append(cleaned)
                options_set.add(cleaned)
        elif stem is None:
            stem = cleaned

    return stem, options_seen


# ── PPTX parser ───────────────────────────────────────────────────────────────

def get_pptx_questions(pptx_path):
    """
    Returns list of dicts:
      { id, stem, options, preview }
    Only slides with a question stem + at least one option are included.
    """
    prs = Presentation(pptx_path)
    result = []

    for idx, slide in enumerate(prs.slides):
        texts = [
            shape.text_frame.text.strip()
            for shape in slide.shapes
            if shape.has_text_frame and shape.text_frame.text.strip()
        ]
        if not texts:
            continue

        block = max(texts, key=len)
        lines = [l.strip() for l in block.split('\n') if l.strip()]

        if len(lines) < 2:
            continue

        stem    = lines[0]
        options = [re.sub(r'^[A-D]\.\s*', '', l).strip() for l in lines[1:]]

        result.append({
            'id':       idx + 1,
            'stem':     stem,
            'options':  options,
            'preview':  stem[:120] + ('…' if len(stem) > 120 else ''),
        })

    return result


# ── Word parser ───────────────────────────────────────────────────────────────

def get_word_questions(docx_path):
    """
    Returns list of dicts:
      { slide_id, slide_name, type, stem, options, preview }
    Only question slides (MCQ / MRQ / TF / DND) are included.
    """
    doc    = Document(docx_path)
    tables = doc.tables
    result = []
    i = 0

    while i < len(tables):
        row0 = [c.text.strip() for c in tables[i].rows[0].cells] if tables[i].rows else []

        if len(row0) == 2 and 'Slide ID' in row0[0]:
            slide_id = row0[1]

            if i + 1 < len(tables):
                ct_rows = [[c.text.strip() for c in r.cells] for r in tables[i + 1].rows]

                if ct_rows and 'ID 🔒' in ct_rows[0][0]:
                    data = ct_rows[1:]
                    slide_name = next(
                        (r[2] for r in data if len(r) > 2 and r[1] == 'Slide name'), ''
                    )
                    q_type = _detect_type(data)

                    if q_type:
                        stem, opts = _stem_and_options(data, q_type)
                        result.append({
                            'slide_id':   slide_id,
                            'slide_name': slide_name,
                            'type':       q_type,
                            'stem':       stem or '',
                            'options':    opts,
                            'preview':    (stem or '')[:120] + ('…' if len(stem or '') > 120 else ''),
                        })

                    i += 2
                    continue

        i += 1

    return result


# ── apply user-defined mappings ───────────────────────────────────────────────

def apply_mappings(docx_path, pptx_path, mappings):
    """
    mappings: list of { word_slide_id: str, pptx_id: int }
    Opens docx, applies each mapping by writing PPTX content into Translation column,
    returns (Document, results_list).
    results_list entries: { slide_name, filled }
    """
    # Index PPTX questions by id
    pptx_list = get_pptx_questions(pptx_path)
    pptx_by_id = {q['id']: q for q in pptx_list}

    # Parse Word doc (need table references for writing)
    doc    = Document(docx_path)
    tables = doc.tables
    word_qs = []
    i = 0

    while i < len(tables):
        row0 = [c.text.strip() for c in tables[i].rows[0].cells] if tables[i].rows else []

        if len(row0) == 2 and 'Slide ID' in row0[0]:
            slide_id = row0[1]

            if i + 1 < len(tables):
                ct      = tables[i + 1]
                ct_rows = [[c.text.strip() for c in r.cells] for r in ct.rows]

                if ct_rows and 'ID 🔒' in ct_rows[0][0]:
                    data = ct_rows[1:]
                    q_type = _detect_type(data)

                    if q_type:
                        stem, opts = _stem_and_options(data, q_type)
                        slide_name = next(
                            (r[2] for r in data if len(r) > 2 and r[1] == 'Slide name'), ''
                        )
                        word_qs.append({
                            'slide_id':   slide_id,
                            'slide_name': slide_name,
                            'type':       q_type,
                            'stem':       stem or '',
                            'options':    opts,
                            'data_rows':  data,
                            'table':      ct,
                        })

                    i += 2
                    continue

        i += 1

    word_by_id = {q['slide_id']: q for q in word_qs}

    # Apply each mapping
    results = []
    for m in mappings:
        wq     = word_by_id.get(m['word_slide_id'])
        pptx_q = pptx_by_id.get(int(m['pptx_id']))

        if not wq or not pptx_q:
            continue

        filled, mismatch = _fill(wq, pptx_q)
        results.append({'slide_name': wq['slide_name'], 'filled': filled, 'mismatch': mismatch})

    return doc, results


def _fill(word_q, pptx_entry):
    """
    Write pptx_entry stem+options into Translation column of word_q['table'].
    Always writes — no similarity check (user explicitly chose this mapping).
    Returns (filled_count, option_mismatch_count).
    """
    text_map = {}
    if word_q['stem'] and pptx_entry['stem']:
        text_map[word_q['stem']] = pptx_entry['stem']
    for w_opt, p_opt in zip(word_q['options'], pptx_entry['options']):
        text_map[w_opt] = p_opt
    mismatch = abs(len(word_q['options']) - len(pptx_entry['options']))

    table    = word_q['table']
    data     = word_q['data_rows']
    filled   = 0
    seen_key = {}

    for ri, table_row in enumerate(table.rows[1:]):
        if ri >= len(data):
            break

        row = data[ri]
        if len(row) < 4:
            continue
        row_id, row_type, src, cur_trans = row[0], row[1], row[2], row[3]

        if row_type == 'Slide name' or not src:
            continue

        # Use source-column run for font reference when translation cell is empty
        src_cell = table_row.cells[2]
        ref_run  = (src_cell.paragraphs[0].runs[0]
                    if src_cell.paragraphs and src_cell.paragraphs[0].runs
                    else None)

        is_multi = '\n\n' in src
        key      = (row_id, row_type)

        if is_multi:
            if key not in seen_key:
                cleaned  = _clean(src)
                new_text = text_map.get(cleaned)
                if new_text:
                    seg1, seg2 = _split_proportional(new_text, cleaned, cur_trans)
                    _set_cell(table_row.cells[3], seg1, ref_run)
                    seen_key[key] = seg2
                    filled += 1
                else:
                    seen_key[key] = None
            else:
                seg2 = seen_key.pop(key)
                if seg2 is not None:
                    _set_cell(table_row.cells[3], seg2, ref_run)
                    filled += 1
        else:
            cleaned  = _clean(src)
            new_text = text_map.get(cleaned)
            if new_text:
                _set_cell(table_row.cells[3], new_text, ref_run)
                filled += 1

    return filled, mismatch
