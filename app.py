"""
DocConvert - convert .doc / .docx / .pdf to Markdown or HTML, with images.

Markdown output supports two modes:
  - single : one merged .md for the whole document
  - split  : one .md per top-level section + an index.md (linked TOC)
HTML output is always a single-file replica of the document.

Author: Mayuresh Piske  |  Hackathon '26
"""
import os, re, shutil, zipfile, tempfile, traceback, codecs, json
from flask import Flask, request, send_file, render_template, jsonify

# Word's filtered HTML is mostly UTF-8 but sprinkles a few raw Windows-1252 bytes
# (e.g. a stray 0xA0). Decode as UTF-8 and fall back to cp1252 for ONLY those bad
# bytes, so multibyte UTF-8 (×, ·, →, —) stays correct and nothing turns into �.
def _u8_1252(err):
    return (err.object[err.start:err.end].decode("cp1252", "replace"), err.end)
codecs.register_error("u8_1252", _u8_1252)

def _decode_word_html(data):
    return data.decode("utf-8", "u8_1252")

app = Flask(__name__)
WORK = os.path.join(tempfile.gettempdir(), "docconvert_v2")
os.makedirs(WORK, exist_ok=True)

# ----------------------------- helpers -----------------------------

def _is_ole2(path):
    with open(path, "rb") as f:
        return f.read(4) == b"\xd0\xcf\x11\xe0"

def _clean_md(t):
    t = t.replace("\u00c2\u00a0", " ").replace("\u00c2", "").replace("\u00a0", " ")
    for bad, good in [("\u00e2\u20ac\u201c", "-"), ("\u00e2\u20ac\u2122", "'"),
                      ("\u00e2\u20ac\u009c", '"'), ("\u00e2\u20ac\u009d", '"'),
                      ("\u00e2\u20ac\u02dc", "'"), ("\u00e2\u20ac\u201d", "--")]:
        t = t.replace(bad, good)
    # SAFE strikethrough removal: keep the text, drop only the markers, same-line, bounded
    # (never span lines - avoids deleting large content blocks)
    t = re.sub(r"~~([^\n]{0,200}?)~~", r"\1", t)
    t = re.sub(r"[ \t]{2,}", " ", t)                      # collapse runs of spaces/tabs
    t = re.sub(r" +\n", "\n", t)                          # trailing spaces
    # blank line BEFORE a heading that is glued to preceding text
    # (group 1 excludes '#' so multi-# heading runs are never split)
    t = re.sub(r"([^\n#])(#{1,6} )", r"\1\n\n\2", t)
    # blank line BEFORE a table that is glued to preceding text
    t = re.sub(r"([^\n|])(\n)(\|)", r"\1\n\n\3", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip() + "\n"

def _slug(s):
    s = re.sub(r"[^\w\s-]", "", s).strip().lower()
    return re.sub(r"[\s]+", "-", s)[:60] or "section"

# nice, readable stylesheet injected into exported HTML
_HTML_CSS = """
<style id="docconvert-style">
  body{max-width:940px;margin:28px auto;padding:0 26px;
       font-family:'Segoe UI',system-ui,-apple-system,Arial,sans-serif;
       line-height:1.65;color:#1f2933;background:#ffffff;}
  h1,h2,h3,h4,h5,h6{color:#12263f;line-height:1.28;margin:1.5em 0 .5em;font-weight:600;}
  h1{font-size:1.9em;border-bottom:2px solid #e6eaf0;padding-bottom:.25em;}
  h2{font-size:1.5em;border-bottom:1px solid #eef1f6;padding-bottom:.2em;}
  h3{font-size:1.25em;} h4{font-size:1.1em;}
  p{margin:.55em 0;}
  a{color:#2563eb;text-decoration:none;} a:hover{text-decoration:underline;}
  img{max-width:100%;height:auto;display:block;margin:16px 0;
      border:1px solid #eef1f6;border-radius:8px;box-shadow:0 2px 10px rgba(15,40,70,.08);}
  /* tables: size to content, don't stretch the whole window */
  table{width:auto !important;max-width:100% !important;border-collapse:collapse !important;
        margin:18px 0;font-size:.94em;box-shadow:0 1px 8px rgba(15,40,70,.06);border-radius:8px;overflow:hidden;display:block;overflow-x:auto;}
  th,td{border:1px solid #d7deea !important;padding:8px 13px !important;
        text-align:left;vertical-align:top;width:auto !important;}
  tr:nth-child(even) td{background:#f7f9fc;}
  code,kbd{background:#f2f4f8;border-radius:5px;padding:1px 5px;font-size:.92em;}
  pre{background:#f2f4f8;border-radius:8px;padding:14px;overflow:auto;}
  hr{border:none;border-top:1px solid #e6eaf0;margin:26px 0;}
  ul,ol{margin:.5em 0 .5em 1.4em;}
</style>
"""

def _prettify_html(raw):
    # 1) fix charset so UTF-8 bytes render correctly (kills the "Â"/"A A A" mojibake)
    raw = re.sub(r'charset\s*=\s*["\']?\s*windows-1252', 'charset=utf-8', raw, flags=re.I)
    raw = re.sub(r'(<meta[^>]*charset=)[^">]*', r'\1utf-8', raw, flags=re.I)
    if not re.search(r'<meta[^>]*charset', raw, re.I):
        raw = re.sub(r'(<head[^>]*>)', r'\1<meta charset="utf-8">', raw, flags=re.I, count=1)
    # 2) strip Word's forced table widths so tables size to their content
    #    (Word writes them unquoted: width=624, or quoted: width="99%")
    raw = re.sub(r'(<table\b[^>]*?)\s+width=(?:"[^"]*"|\'[^\']*\'|[^\s>]+)', r'\1', raw, flags=re.I)
    raw = re.sub(r'(<t[dh]\b[^>]*?)\s+width=(?:"[^"]*"|\'[^\']*\'|[^\s>]+)', r'\1', raw, flags=re.I)
    raw = re.sub(r'(<col\b[^>]*?)\s+width=(?:"[^"]*"|\'[^\']*\'|[^\s>]+)', r'\1', raw, flags=re.I)
    raw = re.sub(r'width:\s*\d+(?:\.\d+)?%', 'width:auto', raw, flags=re.I)
    # 3) inject our readable stylesheet last so it wins
    if '</head>' in raw:
        raw = raw.replace('</head>', _HTML_CSS + '</head>', 1)
    else:
        raw = _HTML_CSS + raw
    return raw

def _unwrap_md(t):
    """Join hard-wrapped lines back into flowing paragraphs. Word's filtered HTML
    wraps source lines mid-sentence (and even mid-**bold**); markdownify keeps those
    newlines, which makes the .md look broken. We merge a wrapped continuation line
    onto the previous line, while leaving blank lines, headings, tables, list markers,
    block quotes, horizontal rules, standalone images and code blocks untouched."""
    def is_block_start(l):
        s = l.lstrip()
        return (s.startswith("#") or s.startswith("|") or s.startswith(">")
                or s.startswith("```") or s.startswith("![")
                or bool(re.match(r"[-*+]\s", s)) or bool(re.match(r"\d+\.\s", s))
                or bool(re.match(r"([-=*_])\1{2,}\s*$", s)))   # hr / setext underline
    out, in_code = [], False
    for l in t.split("\n"):
        if l.lstrip().startswith("```"):
            in_code = not in_code
            out.append(l); continue
        if in_code:
            out.append(l); continue
        s = l.strip()
        if s == "" or is_block_start(l):
            out.append(l if s else ""); continue
        prev = out[-1] if out else ""
        pstrip = prev.strip()
        # append to previous line only if it is an ordinary paragraph/list continuation
        if (pstrip and not is_block_start(prev)
                and not pstrip.startswith(("|", "#", "```"))):
            out[-1] = prev.rstrip() + " " + s
        else:
            out.append(s)
    return "\n".join(out)

def _normalize_word_html(raw):
    """Clean up Word's filtered HTML so it converts to tidy Markdown:
    drop style/script, unwrap bookmark anchors + internal TOC links, unwrap the
    styling <div> wrappers Word puts around headings, and collapse the source
    line-wrapping (Word hard-wraps its HTML text at ~80 cols) inside every text
    node. Returns a BeautifulSoup object."""
    from bs4 import BeautifulSoup, NavigableString
    # Word declares charset=windows-1252 but the bytes are UTF-8; if we don't fix
    # this, lxml re-decodes the text as cp1252 (mojibake) and can truncate the doc.
    raw = re.sub(r'charset\s*=\s*["\']?\s*windows-1252', 'charset=utf-8', raw, flags=re.I)
    soup = BeautifulSoup(raw, "lxml")
    for tag in soup(["style", "script"]):
        tag.decompose()
    for a in soup.find_all("a"):
        if a.get("name") or a.get("href", "").startswith("#"):
            a.unwrap()                                   # keep text, drop the anchor/link
    for d in soup.find_all("div"):
        d.unwrap()                                       # flatten heading wrappers
    for t in list(soup.find_all(string=True)):
        t.replace_with(NavigableString(re.sub(r"\s+", " ", t)))
    return soup

def _html_to_markdown(raw):
    """Convert Word HTML to Markdown BLOCK-BY-BLOCK. Converting the whole document
    in one markdownify() call fails on Word's malformed markup (a single broken
    element earlier in the doc makes every later block render glued with no blank
    lines). Converting each top-level block on its own guarantees clean separation:
    headings, paragraphs, numbered/bulleted lists and tables all stay distinct."""
    from bs4 import NavigableString
    from markdownify import markdownify as md
    soup = _normalize_word_html(raw)
    body = soup.body or soup
    parts = []
    for el in body.children:
        if isinstance(el, NavigableString):
            s = el.strip()
            if s:
                parts.append(s)
            continue
        frag = md(str(el), heading_style="ATX", strip=['o:p', 'span'], bullets="-").strip()
        if frag:
            parts.append(frag)
    return "\n\n".join(parts)

# ----------------------------- converters -----------------------------

def convert_docx_mammoth(path, media_dir, out_fmt):
    import mammoth
    counter = {"n": 0}
    def conv_img(image):
        counter["n"] += 1
        ext = (image.content_type.split("/")[-1] or "png").replace("jpeg", "jpg")
        name = f"fig{counter['n']:03d}.{ext}"
        with image.open() as fh:
            data = fh.read()
        with open(os.path.join(media_dir, name), "wb") as o:
            o.write(data)
        return {"src": f"media/{name}"}
    with open(path, "rb") as f:
        if out_fmt == "html":
            res = mammoth.convert_to_html(f, convert_image=mammoth.images.img_element(conv_img))
            doc = ('<!doctype html><html><head><meta charset="utf-8">'
                   + _HTML_CSS + '</head><body>' + res.value + '</body></html>')
            return doc, counter["n"]
        res = mammoth.convert_to_markdown(f, convert_image=mammoth.images.img_element(conv_img))
        return _clean_md(res.value), counter["n"]

def _word_export_html(src, tmp_html):
    """COM-only step: open src in Word, bake tracked changes, drop comments, then
    SaveAs filtered HTML (+ companion .files image folder).
    Heavy Word COM automation deadlocks on Flask/werkzeug worker threads (STA with
    no message pump), so convert_via_word runs THIS via a subprocess where it executes
    on a fresh process's main thread."""
    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    word = None
    try:
        files_dir = os.path.splitext(tmp_html)[0] + ".files"
        if os.path.exists(files_dir):
            shutil.rmtree(files_dir, ignore_errors=True)
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(src, ReadOnly=False)
        # 1) produce a CLEAN final document: bake tracked changes, drop markup
        try:
            doc.TrackRevisions = False
            if doc.Revisions.Count > 0:
                doc.Revisions.AcceptAll()
        except Exception:
            pass
        try:
            doc.ShowRevisions = False           # render Final, no markup
        except Exception:
            pass
        # 2) remove all comments so they are not dumped into the output
        try:
            doc.DeleteAllComments()
        except Exception:
            try:
                while doc.Comments.Count > 0:
                    doc.Comments(1).Delete()
            except Exception:
                pass
        # 3) export filtered HTML (+ images to a companion folder)
        doc.SaveAs2(tmp_html, FileFormat=10)   # wdFormatFilteredHTML
        doc.Close(False)
    finally:
        if word is not None:
            try:
                word.Quit()                     # always quit so no Word process is orphaned
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_via_word(path, media_dir, out_fmt):
    import base64, mimetypes, sys, subprocess, uuid
    from markdownify import markdownify as md
    stem = "w_" + uuid.uuid4().hex[:12]                  # unique per run: never collides
    tmp_html = os.path.join(WORK, stem + ".htm")
    files_dir = os.path.join(WORK, stem + ".files")
    for p in (tmp_html, files_dir):
        try:
            (shutil.rmtree if os.path.isdir(p) else os.remove)(p)
        except Exception:
            pass
    # Run the Word COM export in a SEPARATE process (see _word_export_html docstring).
    try:
        proc = subprocess.run(
            [sys.executable, os.path.abspath(__file__), "--word-export", path, tmp_html],
            capture_output=True, text=True, timeout=600)
        if proc.returncode != 0 or not os.path.exists(tmp_html):
            raise RuntimeError("Word export failed: " +
                               (proc.stderr or proc.stdout or "unknown error").strip())

        with open(tmp_html, "rb") as f:
            data = f.read()
        raw = _decode_word_html(data)

        # ---- scrub any residual Word comment/markup from the exported HTML ----
        # Word always dumps the comment list as the final block before </body>.
        m = re.search(r"<div[^>]*mso-element:\s*comment-list[^>]*>", raw, re.I)
        if m:
            raw = raw[:m.start()] + "</body></html>"
        # remove inline comment reference anchors + comment ranges
        raw = re.sub(r"<a[^>]*MsoCommentReference[^>]*>.*?</a>", "", raw, flags=re.I | re.S)
        raw = re.sub(r"<span[^>]*mso-comment[^>]*>.*?</span>", "", raw, flags=re.I | re.S)
        raw = re.sub(r"<!--\s*\[if[^>]*\]>.*?<!\[endif\]\s*-->", "", raw, flags=re.I | re.S)
        imgs = {}
        n = 0
        if os.path.isdir(files_dir):
            for fn in sorted(os.listdir(files_dir)):
                if os.path.splitext(fn)[1].lower() in (".png", ".jpg", ".jpeg", ".gif"):
                    n += 1
                    imgs[fn] = os.path.join(files_dir, fn)

        if out_fmt == "html":
            # embed every image as a base64 data URI -> self-contained single .html
            for fn, full in imgs.items():
                mime = mimetypes.guess_type(fn)[0] or "image/png"
                with open(full, "rb") as fh:
                    b64 = base64.b64encode(fh.read()).decode("ascii")
                data_uri = f"data:{mime};base64,{b64}"
                raw = raw.replace(f"{stem}.files/{fn}", data_uri)
                raw = raw.replace(f"./{stem}.files/{fn}", data_uri)
            return _prettify_html(raw), n

        # markdown path: copy images to media/ with clean names + relink
        rename = {}
        i = 0
        for fn, full in imgs.items():
            i += 1
            ext = os.path.splitext(fn)[1].lower()
            nn = f"fig{i:03d}{ext}"
            shutil.copy(full, os.path.join(media_dir, nn))
            rename[fn] = f"media/{nn}"
        markdown = _html_to_markdown(raw)
        def fix_img(m):
            base = os.path.basename(m.group(1))
            return f"]({rename.get(base, 'media/' + base)})"
        markdown = re.sub(r"\]\(([^)]*\.(?:png|jpg|jpeg|gif))\)", fix_img, markdown, flags=re.I)
        # make figures self-describing for LLMs: fold a "Figure N …"/"Table N …" caption
        # that follows an image into the image's alt text (the caption line is kept too)
        markdown = re.sub(
            r"!\[\]\((media/[^)]+)\)\n\n((?:Figure|Table|Waveform|Diagram)\b[^\n]*)",
            lambda m: f"![{m.group(2).strip()}]({m.group(1)})\n\n{m.group(2).strip()}",
            markdown, flags=re.I)
        return _clean_md(markdown), n
    finally:
        for p in (tmp_html, files_dir):                  # tidy up this run's temp files
            try:
                (shutil.rmtree if os.path.isdir(p) else os.remove)(p)
            except Exception:
                pass

def convert_pdf(path, media_dir, out_fmt):
    import pymupdf
    doc = pymupdf.open(path)
    parts, n = [], 0
    for pno, page in enumerate(doc, 1):
        parts.append(f"\n\n<!-- Page {pno} -->\n\n")
        parts.append(page.get_text("text"))
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                pix = pymupdf.Pixmap(doc, xref)
                if pix.n - pix.alpha >= 4:
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                n += 1
                nm = f"fig{n:03d}.png"
                pix.save(os.path.join(media_dir, nm))
                parts.append(f"\n\n![figure](media/{nm})\n\n")
            except Exception:
                pass
    text = "".join(parts)
    if out_fmt == "html":
        body = text.replace("&", "&amp;").replace("<", "&lt;")
        body = re.sub(r"!\[figure\]\(([^)]+)\)", r'<img src="\1">', body)
        return f"<html><body><pre>{body}</pre></body></html>", n
    return re.sub(r"\n{3,}", "\n\n", text).strip() + "\n", n

def base_convert(path, out_fmt):
    media_dir = os.path.join(WORK, "media")
    if os.path.exists(media_dir):
        shutil.rmtree(media_dir, ignore_errors=True)
    os.makedirs(media_dir, exist_ok=True)
    low = path.lower()
    if low.endswith(".pdf"):
        content, n = convert_pdf(path, media_dir, out_fmt)
    elif low.endswith((".doc", ".docx")):
        if _is_ole2(path):
            content, n = convert_via_word(path, media_dir, out_fmt)
        else:
            content, n = convert_docx_mammoth(path, media_dir, out_fmt)
    else:
        raise ValueError("Unsupported file type. Please use .doc, .docx, or .pdf")
    return content, n, media_dir

SPLIT_MAX_BYTES = 40 * 1024      # a section bigger than this is split by sub-headings
SPLIT_MAX_LINES = 500            # ...or bigger than this many lines

def _yaml_str(s):
    s = str(s).replace('"', "'")
    return f'"{s}"'

def _front_matter(source_name, doc_title, title, breadcrumb, fname):
    """A small YAML front-matter block + a visible context line, so an LLM knows
    exactly where this chunk sits in the document even when read in isolation."""
    lines = [
        "---",
        f"source: {_yaml_str(source_name)}",
        f"document: {_yaml_str(doc_title)}",
        f"section: {_yaml_str(title)}",
        f"section_path: {_yaml_str(breadcrumb)}",
        f"file: {fname}",
        "---",
        "",
        f"**Context:** {doc_title} / {breadcrumb}",
        "",
        "",
    ]
    return "\n".join(lines)

def _too_big(text):
    return len(text.encode("utf-8")) > SPLIT_MAX_BYTES or text.count("\n") + 1 > SPLIT_MAX_LINES

def _cut_at_level(text, level):
    """Split text at the given ATX heading level.
    Returns (preamble_before_first_heading, [(title, body_incl_heading), ...])."""
    pat = re.compile(rf"(?m)^#{{{level}}} (.+)$")
    ms = list(pat.finditer(text))
    if not ms:
        return text, []
    pre = text[:ms[0].start()]
    secs = []
    for i, m in enumerate(ms):
        s = m.start()
        e = ms[i + 1].start() if i + 1 < len(ms) else len(text)
        secs.append((m.group(1).strip(), text[s:e].strip() + "\n"))
    return pre, secs

def _shallowest_below(text, level):
    """Shallowest heading level greater than `level` that appears in text (or None)."""
    for lvl in range(level + 1, 7):
        if re.search(rf"(?m)^#{{{lvl}}} ", text):
            return lvl
    return None

def _doc_title_from(md_text, fallback):
    """Document title for front-matter: the clean source name (reliable) — the doc's
    own cover title is often duplicated/boilerplate, so we prefer the file name."""
    return fallback

def split_markdown(md_text, source_name="document"):
    """Split a Markdown document into per-section files. Any section that exceeds
    SPLIT_MAX_LINES lines or SPLIT_MAX_BYTES bytes is split further by its OWN
    sub-headings, recursively, so no single file is unwieldy. Each file gets YAML
    front-matter + a breadcrumb (LLM-friendly), files are named hierarchically, and
    a nested index.md plus a machine-readable manifest are produced.
    Returns (files, index_text, manifest) where files=[(fname, text)]."""
    files, index, manifest = [], ["# Table of Contents", ""], []
    doc_title = os.path.splitext(source_name)[0]

    levels = sorted(set(len(m.group(1)) for m in re.finditer(r"(?m)^(#{1,6}) ", md_text)))
    if not levels:
        body = md_text.strip() + "\n"
        return [("document.md", body)], None, [{"file": "document.md", "title": doc_title,
                "section_path": doc_title, "bytes": len(body.encode("utf-8")),
                "lines": body.count("\n") + 1, "images": []}]
    top = levels[0]

    def record(fname, title, breadcrumb, body, depth):
        text = _front_matter(source_name, doc_title, title, breadcrumb, fname) + body
        files.append((fname, text))
        index.append(f"{'  ' * (depth - 1)}- [{title}]({fname})")
        manifest.append({
            "file": fname, "title": title, "section_path": breadcrumb,
            "bytes": len(text.encode("utf-8")), "lines": text.count("\n") + 1,
            "images": re.findall(r"!\[[^\]]*\]\((media/[^)]+)\)", body),
        })

    def emit(title, body, level, prefix, depth, trail):
        crumb = " / ".join(trail + [title])
        if not _too_big(body):
            record(f"{prefix}-{_slug(title)}.md", title, crumb, body, depth)
            return
        nl = _shallowest_below(body, level)
        if nl is None:
            record(f"{prefix}-{_slug(title)}.md", title, crumb, body, depth)
            return
        pre, subs = _cut_at_level(body, nl)
        record(f"{prefix}-00-{_slug(title)}.md", title, crumb, pre.strip() + "\n", depth)
        for j, (stitle, sbody) in enumerate(subs, 1):
            emit(stitle, sbody, nl, f"{prefix}-{j:02d}", depth + 1, trail + [title])

    pre, secs = _cut_at_level(md_text, top)
    if pre.strip():
        record("00-preamble.md", "Preamble", "Preamble", pre.strip() + "\n", 1)
    for i, (title, body) in enumerate(secs, 1):
        emit(title, body, top, f"{i:02d}", 1, [])

    return files, "\n".join(index) + "\n", manifest

# ----------------------------- routes -----------------------------

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def do_convert():
    try:
        if "file" not in request.files or not request.files["file"].filename:
            return jsonify(error="No file selected"), 400
        up = request.files["file"]
        out_fmt = request.form.get("format", "md")          # md | html
        md_mode = request.form.get("md_mode", "single")     # single | split
        src = os.path.join(WORK, "input_" + re.sub(r"[^\w.\-]", "_", up.filename))
        up.save(src)
        content, n_img, media_dir = base_convert(src, out_fmt)
        base = os.path.splitext(os.path.basename(up.filename))[0]

        zip_path = os.path.join(WORK, f"{base}_{out_fmt}.zip")
        files_written = []
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            if out_fmt == "html":
                z.writestr(f"{base}.html", content)
                files_written.append(f"{base}.html")
            elif md_mode == "split":
                sections, idx, manifest = split_markdown(content, up.filename)
                if idx:
                    z.writestr("index.md", idx)
                    files_written.append("index.md")
                for fname, text in sections:
                    z.writestr(fname, text)
                    files_written.append(fname)
                # machine-readable manifest for RAG / programmatic ingestion
                mani = {
                    "source": up.filename,
                    "mode": "split",
                    "media_dir": "media",
                    "section_count": len(sections),
                    "image_count": n_img,
                    "files": manifest,
                }
                z.writestr("index.json", json.dumps(mani, indent=2, ensure_ascii=False))
                files_written.append("index.json")
            else:
                doc_title = base
                fm = _front_matter(up.filename, doc_title, doc_title, doc_title, f"{base}.md")
                z.writestr(f"{base}.md", fm + content)
                files_written.append(f"{base}.md")
            if os.path.isdir(media_dir):
                for fn in os.listdir(media_dir):
                    z.write(os.path.join(media_dir, fn), f"media/{fn}")

        preview = content[:4000]
        return jsonify(ok=True, images=n_img, chars=len(content),
                       files=files_written, preview=preview,
                       download=f"/download/{os.path.basename(zip_path)}",
                       outname=os.path.basename(zip_path))
    except Exception as e:
        return jsonify(error=str(e), trace=traceback.format_exc()), 500

@app.route("/download/<name>")
def download(name):
    p = os.path.join(WORK, re.sub(r"[^\w.\-]", "_", name))
    if not os.path.exists(p):
        return "Not found", 404
    return send_file(p, as_attachment=True)

if __name__ == "__main__":
    import sys
    # subprocess entry point: run the Word COM export on this process's main thread
    if len(sys.argv) >= 4 and sys.argv[1] == "--word-export":
        _word_export_html(sys.argv[2], sys.argv[3])
        sys.exit(0)
    print("DocConvert v2 -> http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
