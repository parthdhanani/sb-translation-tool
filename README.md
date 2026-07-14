[![ci](https://github.com/parthdhanani/sb-translation-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/parthdhanani/sb-translation-tool/actions/workflows/ci.yml)

# SB Translation Tool

Pushes translated question text straight into an Articulate Storyline 360 storyboard — no manual copy-paste across dozens of slides per language.

Upload the original PPTX storyboard alongside a translated Word doc (or a `.story` file directly), and the tool matches questions between the two, lets you review/adjust the mapping, then writes the translated text back into the file for re-import.

Born from doing this by hand across six languages on a production SCORM pipeline.

## Two workflows

- **PPTX + Word doc** → matches storyboard questions in the PPTX against a translated Word document, applies the mapping, outputs an updated `.docx`.
- **PPTX + `.story` file** → matches directly against a Storyline `.story` project file, outputs an updated `.story` ready to reopen in Storyline.

## Usage

```bash
pip install -r requirements.txt
python app.py            # dev server on :5050
```

Open the page, upload both files for your chosen workflow, review the question mapping, then download the updated file.

## Stack

Flask · `python-pptx` · `python-docx` · gunicorn

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) — no automated test suite yet; changes are verified manually against a real PPTX + storyboard pair.
