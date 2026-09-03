# Transfer Matrix Reshaper

A small web app around `build_modified_matrix.py`: upload a TRANSFER MATRIX
workbook, pick a few options, and download the reshaped "modified matrix"
as an `.xlsx` (or a `.zip` of every block, across every sheet).

## Run it

```bash
pip install -r requirements.txt
python3 app.py
```

Then open **http://localhost:5000**.

## How it works

- **Frontend** — plain HTML/CSS/JS in `static/`. No build step, no framework.
- **Backend** — a small Flask app (`app.py`) that calls straight into your
  existing `build_modified_matrix.py` (copied in unchanged).
- **Flow**:
  1. You drop in a workbook. The frontend immediately posts it to
     `/api/inspect`, which reads the sheet names, every `TRANSFER MATRIX`
     block on each sheet, the detected STA range, and the outbound codes —
     purely from the upload's in-memory stream, without writing anything
     to disk.
  2. You choose a sheet/block, sort order, an optional STA window, and an
     optional outbound-destination filter — or check "process every
     block" to run the whole workbook at once.
  3. Clicking **Generate** re-sends the same file plus your choices to
     `/api/process`, which builds the output and streams it back as a
     download.

## No storage, by design

Every request that touches a file runs inside a
`tempfile.TemporaryDirectory()`, which is deleted the instant the request
finishes — whether it succeeds or fails. The uploaded workbook and the
generated output both live only for the lifetime of that one request; nothing
is written to any permanent location on the server, and nothing accumulates
across uploads. The frontend never stores your file anywhere either — it
just re-sends the same in-browser `File` object each time you submit.

## Notes / things you may want to tweak

- `MAX_CONTENT_LENGTH` in `app.py` caps uploads at 25 MB — raise it if your
  workbooks are bigger.
- The dev server (`python3 app.py`) is fine for local/internal use. For
  anything beyond that, run it behind a real WSGI server (gunicorn, etc.)
  and turn `debug=False` (already the default here).
- `build_modified_matrix.py` is untouched from what you gave me — if you
  change its logic (new columns, different layout, etc.), the app picks it
  up automatically since it just imports from that file.
