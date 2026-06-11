# Contributing to SB Translation Tool

## Setup

```bash
git clone https://github.com/parthdhanani/sb-translation-tool
cd sb-translation-tool
pip install -r requirements.txt
python app.py
```

The app runs on `http://localhost:5050` by default.

## Running tests

```bash
python -m pytest test/ -v
```

There are no automated tests yet — the app is tested manually by uploading sample PPTX and .story files through the UI.

## Reporting issues

Open a GitHub issue. Include:
- The exact file types you uploaded (PPTX version, Storyline version)
- The error message or unexpected output
- A sanitised sample file if you can share one

## Pull requests

1. Fork the repo
2. Create a branch: `git checkout -b fix/your-description`
3. Make changes, test manually against a real PPTX + .story pair
4. Open a PR with a description of what changed and why

