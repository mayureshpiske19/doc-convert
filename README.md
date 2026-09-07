# DocConvert

Convert **Word (.doc / .docx)** and **PDF** documents to **Markdown** or **HTML**, with all images extracted automatically. Runs 100% locally — your files never leave your machine.

Built for Hackathon '26.

## Features

- **Multiple inputs:** `.doc` (legacy binary), `.docx` (OOXML), and `.pdf`
- **Two output formats:**
  - **Markdown** — with a choice of layout:
    - **Single merged `.md`** — one combined file for the whole document
    - **Section-wise split** — one `.md` per section + a linked, nested `index.md`
  - **HTML** — a single-file replica of the document (images embedded as base64)
- **Recursive size-based splitting** — in split mode, any section larger than
  **500 lines or 40 KB** is automatically split further by its own sub-headings,
  recursively, so no single file is unwieldy. Files are named hierarchically.
- **LLM-friendly output** (great for RAG / feeding to an LLM):
  - **YAML front-matter + breadcrumb** on every file (`source`, `document`,
    `section`, `section_path`) so each chunk is self-locating
  - **`index.json` manifest** — machine-readable list of files, titles,
    breadcrumbs, sizes and the images each references
  - **Figure captions folded into image alt-text** so diagrams/waveforms are
    self-describing to a text-only LLM
- **Image extraction** — every figure / diagram / waveform is pulled into a
  `media/` folder and linked inline
- **Clean, LLM-ready Markdown** — converts block-by-block via BeautifulSoup/lxml so
  headings, paragraphs, numbered/bulleted lists and tables stay properly separated;
  bakes tracked changes, drops comments, and fixes Word's UTF-8/Windows-1252 encoding
  quirks (no mojibake)
- **Downloads as a ZIP** containing the output file(s) + `media/` (+ `index.md` /
  `index.json` in split mode)
- Modern drag-and-drop web UI

## How it works

| Input | Engine |
|-------|--------|
| `.docx` (true OOXML) | [mammoth](https://github.com/mwilliamson/python-mammoth) |
| `.doc` / mis-saved OLE2 `.docx` | Microsoft Word (COM, in a subprocess) → filtered HTML → BeautifulSoup/lxml normalize → block-by-block Markdown |
| `.pdf` | [PyMuPDF](https://pymupdf.readthedocs.io/) (text + image extraction) |

> The Word COM path requires Microsoft Word installed (Windows) and runs the export
> in a separate process (heavy Word automation deadlocks on Flask worker threads).
> `.docx` and `.pdf` paths work without Word.

## Setup

```bash
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

## Usage

1. Drag a `.doc` / `.docx` / `.pdf` onto the drop zone (or click to choose).
2. Pick **Markdown** or **HTML**.
3. For Markdown, choose **Single merged** or **Section-wise split**.
4. Click **Convert** and download the ZIP.

## Project layout

```
doc-convert/
├─ app.py               # Flask app + conversion logic
├─ templates/
│  └─ index.html        # UI
├─ requirements.txt
└─ README.md
```

## Notes

- Legacy `.doc` conversion uses Microsoft Word via COM automation and is Windows-only.
- Complex merged-cell tables may flatten in Markdown; use HTML output to preserve them exactly.
- PDF text extraction quality depends on whether the PDF has a real text layer (scanned PDFs need OCR, not included).

## License

MIT
