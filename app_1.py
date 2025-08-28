"""
Flask app to turn a 5‑column Excel into multiple Outlook Web compose windows.

✅ Multi‑user friendly: runs as a stateless web app (no desktop Outlook/COM)
✅ Uploads: one Excel (.xlsx/.xls/.csv) + one image (optional)
✅ Flexible headers: auto‑detects (to, subject, body, cc, bcc) in any order/case
✅ Result: Generates a web page with a button that opens N Outlook Web compose windows

IMPORTANT LIMITATIONS (by design of Outlook Web deeplinks):
- Attachments cannot be pre‑attached via URL. The uploaded image is hosted and
  linked in the body; users can drag/drop or paste it into each draft manually.
- Body is plain text in deeplink. Outlook Web does not honor HTML in the body
  parameter. URLs are auto‑linked.

Run locally:
  pip install flask pandas openpyxl python-dotenv
  set FLASK_APP=app.py (Windows) / export FLASK_APP=app.py (macOS/Linux)
  flask run --host=0.0.0.0 --port=8000

Deploy (examples):
- Any WSGI host (gunicorn/uwsgi + nginx). App is thread/process safe.
"""
from __future__ import annotations

import os
import io
import uuid
import urllib.parse as urlparse
from datetime import datetime
from typing import List, Dict

from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    url_for,
    send_from_directory,
    abort,
)
import pandas as pd
from werkzeug.utils import secure_filename

# ----------------------------
# Config
# ----------------------------
APP_TITLE = "Outlook Web Composer"
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", os.path.join(os.path.dirname(__file__), "uploads"))
MAX_CONTENT_LENGTH = 30 * 1024 * 1024  # 30 MB total payload
ALLOWED_SHEET_EXTS = {".xlsx", ".xls", ".csv"}
ALLOWED_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
ALLOWED_TEXT_EXTS = {".txt"}

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config.update(
    MAX_CONTENT_LENGTH=MAX_CONTENT_LENGTH,
    SEND_FILE_MAX_AGE_DEFAULT=0,
)

# ----------------------------
# Utilities
# ----------------------------

def _ext(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()


def _is_allowed(filename: str, allowed: set[str]) -> bool:
    return "." in filename and _ext(filename) in allowed


def _save_upload(file_storage, subdir: str) -> str:
    os.makedirs(os.path.join(UPLOAD_DIR, subdir), exist_ok=True)
    fname = secure_filename(file_storage.filename)
    ext = _ext(fname)
    uid = uuid.uuid4().hex
    new_name = f"{uid}{ext}"
    save_path = os.path.join(UPLOAD_DIR, subdir, new_name)
    file_storage.save(save_path)
    return new_name


def _read_table(file_storage) -> pd.DataFrame:
    """Read Excel/CSV into a DataFrame. Returns at least 5 columns.
    We keep all columns; mapping to expected fields happens later.
    """
    name = file_storage.filename
    ext = _ext(name)
    raw = file_storage.read()

    if ext == ".csv":
        df = pd.read_csv(io.BytesIO(raw))
    elif ext in {".xlsx", ".xls"}:
        df = pd.read_excel(io.BytesIO(raw))
    else:
        raise ValueError("Unsupported sheet format. Use .xlsx, .xls, or .csv")

    # Drop entirely empty columns/rows
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
    if df.shape[1] < 5:
        raise ValueError("Your file must contain at least 5 columns (to, subject, body, cc, bcc).")
    return df


def _normalize_headers(cols: List[str]) -> Dict[str, str]:
    """Try to map any 5+ columns to canonical fields.
    Priority mapping by name; otherwise fall back to first 5 columns.
    """
    lower = [c.strip().lower() for c in cols]
    mapping = {c: c for c in cols}  # default identity

    def find(*aliases: str) -> str | None:
        for a in aliases:
            if a in lower:
                return cols[lower.index(a)]
        return None

    to_col = find("to", "email", "recipient", "mailto")
    sub_col = find("subject", "subj")
    body_col = find("body", "message", "msg", "content")
    cc_col = find("cc")
    bcc_col = find("bcc")

    chosen = [c for c in [to_col, sub_col, body_col, cc_col, bcc_col] if c]
    # If some are missing, backfill from leftmost unused columns
    unused = [c for c in cols if c not in chosen]
    while len(chosen) < 5 and unused:
        chosen.append(unused.pop(0))

    labels = ["to", "subject", "body", "cc", "bcc"]
    return dict(zip(labels, chosen))


def _encode_for_query(value: str | float | int | None) -> str:
    if pd.isna(value):
        return ""
    if not isinstance(value, str):
        value = str(value)
    # Outlook Web expects URL-encoding; spaces must be %20, not + (to avoid edge cases after login).
    return urlparse.quote(value, safe="@._-:/?=#&\n ")



def _parse_template(file_storage):
    raw = file_storage.read().decode("utf-8", errors="ignore")
    subject, body = None, None
    for line in raw.splitlines():
        if line.lower().startswith("subject_line="):
            subject = line.split("=", 1)[1].strip()
        elif line.lower().startswith("text_email="):
            body = line.split("=", 1)[1].strip()
    return subject, body


def _compose_deeplink(to: str, subject: str, body: str, cc: str = "", bcc: str = "") -> str:
    """Build Outlook Web compose deeplink.
    Officially supported params (2021 docs snapshot): to, subject, body.
    cc/bcc may or may not be honored; we include them best‑effort.
    """
    base = "https://outlook.office.com/mail/deeplink/compose"
    q = [
        ("to", to or ""),
        ("subject", subject or ""),
        ("body", body or ""),
    ]
    if cc:
        q.append(("cc", cc))
    if bcc:
        q.append(("bcc", bcc))
    # encode
    params = "&".join(f"{k}={_encode_for_query(v)}" for k, v in q if v is not None)
    return f"{base}?{params}"


# ----------------------------
# Routes
# ----------------------------
@app.get("/")
def index():
    return render_template_string(INDEX_HTML, app_title=APP_TITLE)


@app.post("/process")
def process():
    # Validate files
    if "sheet" not in request.files or request.files["sheet"].filename == "":
        return _error("Please upload an Excel/CSV file.")

    sheet = request.files["sheet"]
    if not _is_allowed(sheet.filename, ALLOWED_SHEET_EXTS):
        return _error("Sheet must be .xlsx, .xls, or .csv")

    image_url = None
    if "image" in request.files and request.files["image"].filename:
        img = request.files["image"]
        if not _is_allowed(img.filename, ALLOWED_IMAGE_EXTS):
            return _error("Image must be PNG/JPG/GIF/WEBP")
        saved = _save_upload(img, "images")
        image_url = url_for("uploaded_file", subdir="images", filename=saved, _external=True)

    # Read table
    try:
        df = _read_table(sheet)
    except Exception as e:
        return _error(f"Failed to read table: {e}")

    # Map columns
    mapping = _normalize_headers(list(df.columns))
    to_c, sub_c, body_c, cc_c, bcc_c = mapping["to"], mapping["subject"], mapping["body"], mapping["cc"], mapping["bcc"]

    # Handle template if uploaded
    subject_template, body_template = None, None
    if "template" in request.files and request.files["template"].filename:
        tmpl = request.files["template"]
        if _is_allowed(tmpl.filename, ALLOWED_TEXT_EXTS):
            subject_template, body_template = _parse_template(tmpl)

    # Create rows
    rows = []
    for _, r in df.iterrows():
        to = str(r.get(to_c, "") or "").strip()
        subject = (subject_template or str(r.get(sub_c, "") or "")).strip()
        body = (body_template or str(r.get(body_c, "") or "")).strip()
        cc = str(r.get(cc_c, "") or "").strip()
        bcc = str(r.get(bcc_c, "") or "").strip()

        # Append image link (cannot auto‑attach in Outlook Web deeplink)
        if image_url:
            if body:
                body = f"{body}\n\nImage: {image_url}"
            else:
                body = f"Image: {image_url}"

        deeplink = _compose_deeplink(to, subject, body, cc=cc, bcc=bcc)
        rows.append({
            "to": to,
            "subject": subject,
            "cc": cc,
            "bcc": bcc,
            "body": body,
            "deeplink": deeplink,
        })

    if not rows:
        return _error("No rows found in your file.")

    created = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    return render_template_string(RESULTS_HTML,
                                  app_title=APP_TITLE,
                                  rows=rows,
                                  n=len(rows),
                                  image_url=image_url,
                                  created=created)


@app.get("/uploads/<path:subdir>/<path:filename>")
def uploaded_file(subdir: str, filename: str):
    # Serve uploaded assets (images). In production, put behind auth/CDN as needed.
    safe_subdir = secure_filename(subdir)
    directory = os.path.join(UPLOAD_DIR, safe_subdir)
    if not os.path.isfile(os.path.join(directory, filename)):
        abort(404)
    return send_from_directory(directory, filename)


# ----------------------------
# Templates (inline for single‑file deploy)
# ----------------------------
INDEX_HTML = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{{ app_title }}</title>
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 2rem; }
    .card { max-width: 880px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1.25rem; box-shadow: 0 10px 30px rgba(0,0,0,0.06); }
    h1 { margin: 0 0 .5rem; font-size: 1.6rem; }
    p.hint { color: #4b5563; margin-top: 0; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; }
    .row { margin: .75rem 0; }
    label { display: block; font-weight: 600; margin-bottom: .25rem; }
    input[type=file] { width: 100%; padding: .75rem; border: 1px solid #e5e7eb; border-radius: 10px; }
    button { appearance: none; background: black; color: white; border: 0; border-radius: 999px; padding: .75rem 1.25rem; font-weight: 700; }
    .muted { color: #6b7280; font-size: .925rem; }
    ul { margin-top: .25rem; }
    code { background: #f3f4f6; padding: .15rem .4rem; border-radius: 6px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>📬 {{ app_title }}</h1>
    <p class="hint">Upload a 5‑column Excel/CSV (any order): <code>to</code>, <code>subject</code>, <code>body</code>, <code>cc</code>, <code>bcc</code>. Headers are auto‑detected (case‑insensitive). Then optionally upload one image to reference in the email body.</p>

    <form method="post" action="/process" enctype="multipart/form-data">
      <div class="row">
        <label>Excel/CSV file</label>
        <input type="file" name="sheet" accept=".xlsx,.xls,.csv" required />
      </div>
      <div class="row">
        <label>Template (optional .txt)</label>
        <input type="file" name="template" accept=".txt" />
        <div class="muted">Upload a .txt file with lines starting with <code>subject_line=</code> and <code>text_email=</code>.</div>
      </div>
      <div class="row">
        <label>Image (optional)</label>
        <input type="file" name="image" accept="image/*" />
        <div class="muted">Due to Outlook Web limitations, attachments cannot be auto‑added. The image URL will be included in the body for easy copy/paste.</div>
      </div>
      <div class="row">
        <button type="submit">Create compose links</button>
      </div>
    </form>

    <hr/>
    <div class="muted">
      <strong>Notes & Tips</strong>
      <ul>
        <li>Make sure recipients are comma‑separated in <code>to</code>/<code>cc</code>/<code>bcc</code>.</li>
        <li>Body is inserted as plain text; Outlook Web does not support HTML in compose deeplinks.</li>
        <li>Pop‑up blockers may prevent opening multiple windows at once. Use the provided
            <em>Open all</em> button on the results page — it opens windows with a short stagger to reduce blocking.</li>
      </ul>
    </div>
  </div>
</body>
</html>
"""

RESULTS_HTML = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{{ app_title }} – Results</title>
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 2rem; }
    .card { max-width: 1100px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1.25rem; box-shadow: 0 10px 30px rgba(0,0,0,0.06); }
    h1 { margin: 0 0 .5rem; font-size: 1.45rem; }
    .muted { color: #6b7280; font-size: .925rem; }
    table { width: 100%; border-collapse: collapse; margin-top: 1rem; }
    th, td { text-align: left; border-bottom: 1px solid #f3f4f6; padding: .5rem .4rem; vertical-align: top; }
    code { background: #f3f4f6; padding: .15rem .4rem; border-radius: 6px; }
    .btns { display: flex; gap: .5rem; margin-top: 1rem; flex-wrap: wrap; }
    button { appearance: none; background: black; color: white; border: 0; border-radius: 999px; padding: .6rem 1rem; font-weight: 700; cursor: pointer; }
    a.btn { display: inline-block; text-decoration: none; background: #111827; color: white; border-radius: 999px; padding: .6rem 1rem; font-weight: 700; }
  </style>
</head>
<body>
  <div class="card">
    <h1>✅ Generated {{ n }} compose links</h1>
    <div class="muted">Created {{ created }}. {% if image_url %} Image hosted at: <code>{{ image_url }}</code>{% endif %}</div>

    <div class="btns">
      <button id="openAll">Open {{ n }} compose windows</button>
      <a class="btn" href="/">Start over</a>
    </div>

    <table>
      <thead>
        <tr><th>#</th><th>To</th><th>Subject</th><th>CC</th><th>BCC</th><th>Body (preview)</th><th>Action</th></tr>
      </thead>
      <tbody>
        {% for row in rows %}
          <tr>
            <td>{{ loop.index }}</td>
            <td>{{ row.to }}</td>
            <td>{{ row.subject }}</td>
            <td>{{ row.cc }}</td>
            <td>{{ row.bcc }}</td>
            <td style="max-width: 360px; white-space: pre-wrap">{{ row.body[:400] }}{% if row.body|length > 400 %}…{% endif %}</td>
            <td><a class="btn" target="_blank" rel="noopener" href="{{ row.deeplink }}">Open</a></td>
          </tr>
        {% endfor %}
      </tbody>
    </table>

    <p class="muted" style="margin-top: 1rem">
      <strong>Heads‑up:</strong> Outlook Web compose deeplink supports <code>to</code>, <code>subject</code>, <code>body</code> reliably.
      <code>cc</code>/<code>bcc</code> parameters may be ignored in some tenants/browsers.
      Attachments cannot be pre‑attached via URL. Paste/drag the image if needed.
    </p>
  </div>

  <script>
    (function(){
      const links = [
        {% for row in rows %}"{{ row.deeplink }}"{% if not loop.last %},{% endif %}{% endfor %}
      ];
      const openAll = document.getElementById('openAll');
      openAll.addEventListener('click', function(){
        const delay = 350; // ms; stagger to reduce popup blocking
        links.forEach((href, idx) => setTimeout(() => {
          window.open(href, '_blank', 'noopener');
        }, idx * delay));
      });
    })();
  </script>
</body>
</html>
"""


# ----------------------------
# Error helper
# ----------------------------
def _error(msg: str):
    return render_template_string(
        """
        <html><head><meta charset='utf-8'><title>Error</title></head>
        <body style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; padding: 2rem;">
          <div style="max-width: 760px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1rem;">
            <h2>⚠️ Error</h2>
            <p>{{ msg }}</p>
            <a href="/">Back</a>
          </div>
        </body></html>
        """,
        msg=msg,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
