# QBG Doc Formatter

Turns test-paper Word files into one Word doc per question in the exact QBG
ingestion layout, assigns serial numbers (`QBGDOC000001`, …), and stores the
originals, the question docs and an Excel record per file in Google Drive.
Google Drive is the only external dependency.

Creator: Aditya Raj · aditya.raj3@pw.live

---

## What it does

1. Upload any number of `.docx` / `.doc` files and give each a **QBG Category** and **File Name**.
2. **Check parsing** (instant, offline): see how many questions were found, how many are
   ready, and exactly which ones need attention and why, before touching Drive.
3. **Format and upload** — Drive layout:
   ```
   Root/Input/<Category>/<File Name>.docx                 original, as uploaded
   Root/Output/<Category>/<File Name>.xlsx                question-level record (cols A–G)
   Root/Output/<Category>/<File Name>/QBGDOC000001.docx   one doc per question
   Root/Output/_serial_counter.json                       last serial issued
   ```
   Excel columns A–G: Category, File Name, Serial, Original doc link, Question doc
   link, Status, Remarks (plus a Summary sheet with the counts).
4. Live progress with a **Stop** button, then **Excel / CSV** download of the whole batch.

Question docs are uploaded as plain `.docx` (no Google Docs conversion), which is
the fastest way to put a file in Drive. If an Excel record with the same name
already exists (re-run of the same File Name), the new one gets the run ID
appended, so earlier records are never overwritten.

Question docs are built from `assets/question_template.docx` (your reference sample), so
fonts, colours, borders and row layout match it exactly. Bold, italic, underline,
tables (match-the-following) and images are carried over from the source.

## How questions are recognised

| Part | Recognised forms |
|---|---|
| Question start | `1.` `1)` `1:` `Q1.` `Q.1` `Question 1:` `13 'Admission'…` — typed or Word auto-numbered |
| Options | `(a)` `a)` `a.` `(A)` `A)` `A.` — also `(1)` `1)` when no letter options exist; options inside a table |
| Answer | `Ans. (d)` `Ans.b` `Ans.- c` `Answer: B)` `Answer:(b)` `Correct Answer: Option B` … |
| Solution | everything after the answer up to the next question; `Explanation:` / `Solution:` labels removed |

- **Type 3 (statements labelled like options):** options are found by scanning
  *up* from the answer line for the last complete run `(d) (c) (b) (a)`. Anything
  above option (a), including statements labelled (a)/(b)/(c), stays in the question.
- **Headers / instructions / title pages** are ignored automatically, even when the
  instructions themselves are a numbered list.
- **Numbered lists inside explanations** are not mistaken for the next question.
- **Section headings** between questions (e.g. `COMPETITION ACT`) are dropped.
- **A question with no number** is still detected; it is flagged in Remarks for review.

A question is marked **Failure** (no doc created, reason in Remarks) when its
options, answer key, question text or explanation is missing, or the answer letter
does not match any option. Every Remarks entry starts with the source question
number, e.g. `Source Q.23: answer key not found`.

---

## Setup (one time, ~10 minutes)

### 1. Google Cloud project
1. <https://console.cloud.google.com> → create a project (e.g. `qbg-formatter`).
2. **APIs & Services → Library**: enable **Google Drive API**.

### 2. Credentials — choose one

**Option A — Google account login (recommended).** Files are owned by a real pw.live
account and go straight into the existing My Drive folder.

1. **OAuth consent screen** → User type **Internal** (keeps the login from expiring
   every 7 days).
2. **Credentials → Create credentials → OAuth client ID → Desktop app** → download
   the JSON as `client_secret.json`.
3. On your laptop:
   ```bash
   pip install google-auth-oauthlib
   python scripts/get_refresh_token.py client_secret.json
   ```
   Sign in with the pw.live account that owns (or can edit) the Drive folder.
   Copy the printed `[google_oauth]` block.

**Option B — service account.** Only works if the root folder is inside a
**Shared Drive** with the service account added as a *Content manager*. Service
accounts have no My Drive storage, so uploading into a normal folder fails with
`storageQuotaExceeded`.

### 3. Secrets
Copy `.streamlit/secrets.toml.example` → fill in the credentials block. The root
folder ID and parallelism are already set.

---

## Deploy as a public link (Streamlit Community Cloud)

1. Push this folder to a GitHub repo (the real `secrets.toml` is git-ignored).
   ```bash
   git init && git add . && git commit -m "QBG Doc Formatter"
   git branch -M main
   git remote add origin https://github.com/<you>/qbg-doc-formatter.git
   git push -u origin main
   ```
2. <https://share.streamlit.io> → **Create app** → pick the repo, branch `main`,
   main file `app.py`.
3. **Advanced settings → Secrets** → paste your secrets.
4. Deploy. In app settings set sharing to **Public** so anyone with the link can use it.

The Google credentials stay on the server; visitors never see them.

### Run locally
```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then fill it in
streamlit run app.py
```

---

## Speed and limits

- Parsing and doc building are local and near-instant (121 questions ≈ 0.3 s to
  parse, ≈ 1 s to build all docs).
- Drive calls per file: 1 original + 1 per question + 1 Excel, plus a few folder
  lookups. Question uploads run in parallel (**Parallel uploads**, default 12).
  Google throttling is retried automatically with backoff.
- **Serial numbers** live in `Output/_serial_counter.json`. The range for a run is
  written there *before* any upload, so a crash or stop never reuses a serial.
  Allocation is also locked inside the app, so two people running batches at once
  never get the same serials. If the counter file is deleted, the app recovers by
  finding the highest `QBGDOC` file name in Drive. Don't edit it by hand unless you
  mean to move the numbering.
- **Stop**: uploads already in flight finish; every remaining question is recorded
  in that file's Excel as `Failure – Not processed: run stopped by user`, so every
  serial is accounted for. Re-upload the file to process them again (new serials).
- If the page is refreshed during a run, the run continues and the page re-attaches
  to it (the run ID is kept in the URL).
- Question doc links open in Drive's viewer for anyone with access to the folder.
  A system that fetches them programmatically needs Drive access too; the file ID
  is the part of the link after `/d/`.

## Checking a new document style offline
```bash
python scripts/parse_local.py "new_paper.docx" out_folder --category Judiciary
```
Writes one `.docx` per question locally and prints any question that needs attention.

## Tests
```bash
python tests/test_parser.py     # sample paper + 7 synthetic layouts
python tests/test_pipeline.py   # Drive layout, serial continuity, stop mid-run (fake Drive)
```

## Project layout
```
app.py                      Streamlit UI
qbg/docx_reader.py          Word → lines/tables (auto-numbering, line breaks, images)
qbg/parser.py               lines → questions, options, answer, solution
qbg/builder.py              question → QBG-format .docx from the template
qbg/google_services.py      Drive (auth, folders, uploads, serial counter, retries)
qbg/pipeline.py             background job, serials, stop, per-file Excel, CSV export
assets/question_template.docx   reference layout
assets/creator.png
scripts/get_refresh_token.py    one-time OAuth helper
scripts/parse_local.py          offline CLI
tests/
```

## Troubleshooting

| Message | Fix |
|---|---|
| `storageQuotaExceeded` | You're using a service account with a My Drive folder. Use Option A, or move the folder to a Shared Drive. |
| `File not found` / 404 on the root folder | The signed-in account (or service account) has no access to the folder ID in secrets. |
| `invalid_grant` | The refresh token expired or was revoked. Set the consent screen to Internal and rerun `get_refresh_token.py`. |
| `insufficient scopes` after upgrading | Tokens made for the old version still work (Drive scope is a subset). If not, rerun `get_refresh_token.py`. |
| A question shows `options … could not be identified` | Its options don't start with a recognisable label. Fix the source line and re-upload. |
| `answer keys in file = X, questions detected = Y` | A question boundary is unusual; open the listed remarks to find it. |
