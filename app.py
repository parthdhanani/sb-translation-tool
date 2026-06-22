import os
import re
import shutil
import sys
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

sys.path.insert(0, os.path.dirname(__file__))
import parser as sb_parser
import story_parser as sp

app = Flask(__name__)

SESSIONS      = Path('/tmp/sb_tool_sessions')
MAX_AGE_S     = 1800          # 30-min TTL (files deleted on download anyway)
MAX_SESSIONS  = 15            # max concurrent sessions → ~600 MB worst case
MAX_FILE_MB   = 80
MAX_FILE_B    = MAX_FILE_MB * 1024 * 1024
app.config['MAX_CONTENT_LENGTH'] = (MAX_FILE_MB * 2 + 10) * 1024 * 1024  # ~170 MB request cap

SESSIONS.mkdir(exist_ok=True)

# Purge sessions left over from a previous run before accepting new traffic
_startup_cutoff = time.time() - MAX_AGE_S
for _d in list(SESSIONS.iterdir()):
    try:
        if _d.is_dir() and _d.stat().st_mtime < _startup_cutoff:
            shutil.rmtree(_d, ignore_errors=True)
    except Exception:
        pass


# ── helpers ────────────────────────────────────────────────────────────────────

def _check_sessions():
    """Evict expired sessions inline, then return error if cap still reached."""
    cutoff = time.time() - MAX_AGE_S
    try:
        for d in list(SESSIONS.iterdir()):
            if d.is_dir() and d.stat().st_mtime < cutoff:
                shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass
    try:
        count = sum(1 for d in SESSIONS.iterdir() if d.is_dir())
    except Exception:
        count = 0
    if count >= MAX_SESSIONS:
        return jsonify({'error': 'Server is busy — too many active sessions. Try again in a few minutes.'}), 503
    return None


def _check_file_size(path: Path, label: str):
    """Return error response if saved file exceeds limit, else None."""
    size = path.stat().st_size
    if size > MAX_FILE_B:
        return jsonify({'error': f'{label} exceeds the {MAX_FILE_MB} MB limit ({size // 1024 // 1024} MB uploaded).'}), 413
    return None


def _check_filetype(path: Path, label: str):
    """Return error response if file lacks a ZIP magic header, else None.
    PPTX, DOCX, and .story are all ZIP-based — rejecting non-ZIP catches wrong
    file uploads before the parsers attempt to open them."""
    try:
        with open(path, 'rb') as fh:
            magic = fh.read(4)
        if magic[:2] != b'PK':
            return jsonify({'error': f'{label} does not appear to be a valid file. '
                                     f'Please upload a PPTX, DOCX, or .story file.'}), 400
    except Exception:
        pass
    return None


# ── error handlers ─────────────────────────────────────────────────────────────

@app.errorhandler(413)
def too_large(_):
    return jsonify({'error': f'Upload too large. Maximum {MAX_FILE_MB} MB per file.'}), 413


# ── routes ─────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/scan', methods=['POST'])
def scan():
    cap = _check_sessions()
    if cap:
        return cap

    pptx_file = request.files.get('pptx')
    word_file = request.files.get('word')
    if not pptx_file or not word_file:
        return jsonify({'error': 'Both files are required.'}), 400

    session_id  = str(uuid.uuid4())
    session_dir = SESSIONS / session_id
    session_dir.mkdir()

    pptx_path = session_dir / 'input.pptx'
    word_path = session_dir / 'input.docx'
    pptx_file.save(str(pptx_path))
    word_file.save(str(word_path))

    err = (_check_file_size(pptx_path, 'PPTX') or _check_file_size(word_path, 'DOCX') or
           _check_filetype(pptx_path, 'Storyboard PPTX') or _check_filetype(word_path, 'Translation DOCX'))
    if err:
        shutil.rmtree(session_dir, ignore_errors=True)
        return err

    (session_dir / 'orig_name.txt').write_text(word_file.filename)

    try:
        pptx_qs = sb_parser.get_pptx_questions(str(pptx_path))
        word_qs = sb_parser.get_word_questions(str(word_path))
    except Exception as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({'error': str(e)}), 500

    return jsonify({'session_id': session_id,
                    'pptx_questions': pptx_qs,
                    'word_questions': word_qs})


@app.route('/apply', methods=['POST'])
def apply():
    data        = request.get_json()
    session_id  = data.get('session_id')
    if not _valid_sid(session_id):
        return jsonify({'error': 'Invalid session.'}), 400
    mappings    = data.get('mappings', [])
    session_dir = SESSIONS / session_id
    pptx_path   = session_dir / 'input.pptx'
    word_path   = session_dir / 'input.docx'
    out_path    = session_dir / 'output.docx'

    if not pptx_path.exists() or not word_path.exists():
        return jsonify({'error': 'Session expired or not found. Please re-upload your files.'}), 404

    try:
        doc, results = sb_parser.apply_mappings(str(word_path), str(pptx_path), mappings)
        doc.save(str(out_path))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({'ok': True, 'session_id': session_id, 'results': results})


@app.route('/scan-story', methods=['POST'])
def scan_story():
    cap = _check_sessions()
    if cap:
        return cap

    pptx_file  = request.files.get('pptx')
    story_file = request.files.get('story')
    if not pptx_file or not story_file:
        return jsonify({'error': 'Both PPTX and .story files are required.'}), 400

    session_id  = str(uuid.uuid4())
    session_dir = SESSIONS / session_id
    session_dir.mkdir()

    pptx_path  = session_dir / 'input.pptx'
    story_path = session_dir / 'input.story'
    pptx_file.save(str(pptx_path))
    story_file.save(str(story_path))

    err = (_check_file_size(pptx_path, 'PPTX') or _check_file_size(story_path, '.story file') or
           _check_filetype(pptx_path, 'Storyboard PPTX') or _check_filetype(story_path, '.story file'))
    if err:
        shutil.rmtree(session_dir, ignore_errors=True)
        return err

    (session_dir / 'orig_name.txt').write_text(story_file.filename)

    try:
        pptx_qs  = sb_parser.get_pptx_questions(str(pptx_path))
        story_qs = sp.get_story_questions(str(story_path))
    except Exception as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({'error': str(e)}), 500

    return jsonify({'session_id': session_id,
                    'pptx_questions': pptx_qs,
                    'story_questions': story_qs})


@app.route('/apply-story', methods=['POST'])
def apply_story():
    data        = request.get_json()
    session_id  = data.get('session_id')
    if not _valid_sid(session_id):
        return jsonify({'error': 'Invalid session.'}), 400
    mappings    = data.get('mappings', [])
    session_dir = SESSIONS / session_id
    pptx_path   = session_dir / 'input.pptx'
    story_path  = session_dir / 'input.story'
    out_path    = session_dir / 'output.story'

    if not pptx_path.exists() or not story_path.exists():
        return jsonify({'error': 'Session expired or not found. Please re-upload your files.'}), 404

    try:
        pptx_qs = sb_parser.get_pptx_questions(str(pptx_path))
        results = sp.apply_story_mappings(str(story_path), pptx_qs, mappings, str(out_path))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({'ok': True, 'session_id': session_id, 'results': results})


_UUID_RE = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')


def _valid_sid(session_id: str) -> bool:
    return bool(_UUID_RE.match(session_id or ''))


@app.route('/download/<session_id>')
def download(session_id):
    if not _valid_sid(session_id):
        return 'Not found', 404
    session_dir = SESSIONS / session_id
    out_path    = session_dir / 'output.docx'
    if not out_path.exists():
        return 'Not found', 404
    orig_name = (session_dir / 'orig_name.txt').read_text()
    return send_file(str(out_path), as_attachment=True,
                     download_name='updated_' + orig_name)


@app.route('/download-story/<session_id>')
def download_story(session_id):
    if not _valid_sid(session_id):
        return 'Not found', 404
    session_dir = SESSIONS / session_id
    out_path    = session_dir / 'output.story'
    if not out_path.exists():
        return 'Not found', 404
    orig_name = (session_dir / 'orig_name.txt').read_text()
    return send_file(str(out_path), as_attachment=True,
                     download_name='updated_' + orig_name)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(debug=False, host='0.0.0.0', port=port)
