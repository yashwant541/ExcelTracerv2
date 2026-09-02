# Excel Formula Tracer

A **standalone** Dataiku Standard webapp that traces the formulas in **any** `.xlsx` /
`.xlsm` workbook. Point it at a managed folder + filename, or upload a file, then walk any
cell's calculation node by node — which `IF` branch is live, what each sub-expression
evaluates to, what a `VLOOKUP` matched, the precedents back to the raw inputs, plus a
workbook-wide lint. No config, no auth, read-only.

Formula-only workbooks (no cached results — what openpyxl / most Python writers produce)
are fine: every formula is **evaluated** by the engine.

```
excel-tracer-webapp/
  frontend/  body.html · style.css · app.js
  backend/   backend.py
  python-lib/  webapp_core/ · excel_deep_trace.py    ← add to the project Python libraries
  dev/serve.py                                        ← local preview (:5179)
```

Self-contained — move or rename the folder freely. `python-lib/` is a copy of the shared
formula engine, same as the RWA project's.

## Deploy

1. **+ New → Webapp → Standard → "Empty (code your own)"**. Enable the Python backend.
2. **Project → Libraries → Python** — add this folder's `python-lib/` (so `import
   webapp_core` and `import excel_deep_trace` resolve).
3. Paste `frontend/body.html` → HTML, `frontend/style.css` → CSS, `frontend/app.js` → JS,
   `backend/backend.py` → Python.
4. **Save**, **Restart backend**, **View web app**.
5. **Settings → Security / Connections** — grant read access to any managed folders whose
   workbooks you want to open (upload works without this).

## Using it

* **From a managed folder** — pick a folder (or paste its id) → *List files* → pick a
  workbook → *Open & trace*.
* **Upload** — drop an `.xlsx` / `.xlsm` (parsed in memory, nothing stored).

Then: pick a **sheet + row** for a row journey (every formula cell, each with its
plain-English narrative), or type a **cell** (`Sheet!A1`) to open its full trace drawer.
The left rail also has a **mismatch scan** (rows where two columns disagree), the
**findings** list, and a searchable **formulas** browser. Cell links inside the trace tree
re-root the trace on that cell (breadcrumb back-tracking).

### Steps page (formula transformation + debugger)

From a row-journey card or the trace drawer, **Steps →** opens a full-page
transformation: the formula reduced one operation at a time — resolve the cell
references, then evaluate each sub-expression innermost-first, showing the running
formula after every step. A **debugger bar** (⏮ ◀ ▶ ▶▶ ⏭ / Run) walks the steps; each
`VLOOKUP` step shows the matched table row and each `SUMIF/SUMIFS/COUNTIF(S)/AVERAGEIF(S)`
step shows the **matching-records table**. The side panel has a copy-pasteable *"Why is
this number here?"* lineage, and **Download JSON** exports the whole transformation.

### Flowchart page

From a row-journey card or the trace drawer, **Flowchart →** opens a full-page
backtracking flowchart for that cell. It is generated from the live expression tree — a
tidy-tree layout, inputs at the top flowing down to the result, edges labelled with the
value passed, boxes coloured by role (formula / lookup / source / calculation / condition
/ result / error / not-taken). Drag to pan, scroll to zoom, click a box for detail
(including the matched `VLOOKUP` row), click a blue formula-cell box to re-root the chart
on it. Toggles: *hide constants*, *only the taken path*. No external chart library.

## Local preview

```bash
cd excel-tracer-webapp
pip install flask openpyxl pandas numpy
python dev/serve.py     # http://127.0.0.1:5179  — "Open & trace" loads a demo workbook
```

## API (all thin wrappers over `ExcelDeepTraceEngine`)

`GET /api/health` · `GET /api/folders` · `GET /api/folder-files?folder_id=` ·
`POST /api/open {folder_id, path}` · `POST /api/upload` (multipart) ·
`POST /api/columns {session_id, sheet}` · `POST /api/journey {session_id, sheet, row}` ·
`POST /api/tree {session_id, cell}` · `POST /api/explain {session_id, cell}` ·
`POST /api/transformation {session_id, cell}` · `POST /api/explain-plain {session_id, cell}` ·
`POST /api/mismatch {session_id, sheet, left, right}` ·
`GET /api/findings?session_id=` · `GET /api/formulas?session_id=[&sheet=&function=]`
