# QBG Doc Formatter

Converts test-paper Word files into one Word doc per question in the QBG ingestion
format, numbers each question (`QBGDOC000001`, `QBGDOC000002`, …), saves everything
in Google Drive and gives an Excel record with a Drive link for every question.

Creator: Aditya Raj (aditya.raj3@pw.live)

---

## Contents

1. [What the app does](#1-what-the-app-does)
2. [Before you start](#2-before-you-start)
3. [Deployment guide](#3-deployment-guide)
   - [Step 1: Set up Google Cloud](#step-1-set-up-google-cloud-10-min)
   - [Step 2: Create the Drive login token](#step-2-create-the-drive-login-token-5-min)
   - [Step 3: Put the code on GitHub](#step-3-put-the-code-on-github-5-min)
   - [Step 4: Deploy on Streamlit Community Cloud](#step-4-deploy-on-streamlit-community-cloud-5-min)
   - [Step 5: Make the link public](#step-5-make-the-link-public-1-min)
   - [Step 6: Test it](#step-6-test-it-2-min)
4. [How to use the app (for the team)](#4-how-to-use-the-app-for-the-team)
5. [Updating the app later](#5-updating-the-app-later)
6. [Troubleshooting](#6-troubleshooting)
7. [How questions are recognised](#7-how-questions-are-recognised)
8. [Technical reference](#8-technical-reference)

---

## 1. What the app does

1. The user uploads one or more Word files (`.docx` or `.doc`).
2. For each file they enter a **QBG Category** (for example `Judiciary`) and a **File Name**.
3. The app finds every question, its options, the correct answer and the explanation.
4. Everything is saved in Google Drive like this:

```
Your Drive folder (root)
├── Input
│   └── Judiciary
│       └── Full Length Test 5.docx          ← the original file, as uploaded
└── Output
    ├── _serial_counter.json                 ← remembers the last serial number used
    └── Judiciary
        ├── Full Length Test 5.xlsx          ← one row per question, with links
        └── Full Length Test 5
            ├── QBGDOC000001.docx            ← one doc per question, QBG format
            ├── QBGDOC000002.docx
            └── …
```

5. English + Hindi questions are written into the bilingual template (English and Hindi
   columns side by side). English-only and Hindi-only questions use the single-column
   template, with **Input Language** set to English or Hindi.
6. The user can also download the Excel (or CSV) for the whole batch. Columns:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| QBG Category | File Name | Question Serial No. | Original Doc Link | Question Doc Link | Status | Remarks |

---

## 2. Before you start

You need three things. Have them ready before Step 1.

| What | Why | Where to get it |
|---|---|---|
| The **pw.live Google account** that owns the Drive folder `1gzMrWz-I7yrqW6S_viabr5QI4G9QY8QL` | The app saves files into Drive as this account | You already have it |
| A **GitHub account** | Streamlit hosts apps from GitHub | <https://github.com/signup> (free) |
| A **Streamlit Community Cloud account** | This hosts the app and gives you the public link | <https://share.streamlit.io>: sign in with GitHub (free) |

You do **not** need Python or any software installed. Everything below is done in a web browser.

---

## 3. Deployment guide

Total time: about 30 minutes, once. After that the app runs on its own.

### Step 1: Set up Google Cloud (10 min)

This tells Google that your app is allowed to use Google Drive.

**1.1 Create a project**

1. Go to <https://console.cloud.google.com> and sign in with the **pw.live account** that owns the Drive folder.
2. At the top of the page, click the project dropdown (next to the Google Cloud logo) → **New Project**.
3. Project name: `qbg-doc-formatter`. Leave Organization as `pw.live`. Click **Create**.
4. Wait a few seconds, then make sure the new project is selected in that same dropdown.

**1.2 Turn on the Google Drive API**

1. In the search bar at the top, type **Google Drive API** and open it.
2. Click **Enable**.

**1.3 Set up the sign-in screen**

Google renamed this area recently. It may be called **Google Auth Platform** or **OAuth consent screen**; both are the same thing.

1. In the search bar, type **OAuth consent screen** and open it. If you see **Get started**, click it.
2. **App name:** `QBG Doc Formatter`. **User support email:** your pw.live email. Click **Next**.
3. **Audience / User type:** choose **Internal**. Click **Next**.
   > Choose **Internal**, not External. With External, the login stops working after 7 days.
4. **Contact email:** your pw.live email. Click **Next**, agree to the policy, click **Create**.

**1.4 Create the OAuth client**

1. Go to **Clients** (or **Credentials → Create credentials → OAuth client ID**).
2. Click **Create client**.
3. **Application type:** **Web application**. **Name:** `qbg-formatter`.
4. Under **Authorized redirect URIs**, click **Add URI** and paste exactly:
   ```
   https://developers.google.com/oauthplayground
   ```
   (no slash at the end)
5. Click **Create**.
6. A box shows your **Client ID** and **Client secret**. Copy both into a Notepad / TextEdit
   file now, or click **Download JSON**.
   > Google may show the client secret only once. If you lose it, open the client,
   > click **Add secret**, and use the new one.

Keep the client secret private. Do not upload it to GitHub or share it in chats.

---

### Step 2: Create the Drive login token (5 min)

This produces the **refresh token**, which lets the app sign in to Drive by itself.
It's done in your browser with Google's **OAuth 2.0 Playground**.

**2.1 Enter your own client details**

1. Open <https://developers.google.com/oauthplayground>.
2. Click the **gear icon** (top right, "OAuth 2.0 configuration").
3. Leave **OAuth flow** as *Server-side* and **Access type** as *Offline*.
4. Tick **Use your own OAuth credentials**.
5. Paste your **OAuth Client ID** and **OAuth Client secret** from Step 1.4.
6. Click **Close**.
   > This step matters. Without your own credentials, Google cancels the token after 24 hours.

**2.2 Choose the Drive permission and sign in**

1. In the left panel, **Step 1 – Select & authorize APIs**, find the box
   **"Input your own scopes"** at the bottom and paste:
   ```
   https://www.googleapis.com/auth/drive
   ```
2. Click **Authorize APIs**.
3. Choose the **pw.live account that owns the Drive folder**, then click **Allow** / **Continue**.

**2.3 Get the refresh token**

1. You're sent back to the Playground with **Step 2 – Exchange authorization code for tokens** open.
2. Click **Exchange authorization code for tokens**.
3. A **Refresh token** box appears (starts with `1//`). Copy the whole value.
   - Ignore the **Access token**: it expires in an hour, and the app makes new ones itself.
   - If no refresh token appears, click the **gear icon**, confirm **Access type** is *Offline*,
     and repeat 2.2.

**2.4 Put the three values together**

In your Notepad file, arrange them like this. You'll paste it in Step 4:

```toml
[google_oauth]
client_id = "123456789-abc123.apps.googleusercontent.com"
client_secret = "GOCSPX-AbCdEf123456"
refresh_token = "1//0gXyZaBcDeF..."
```

The `client_id` and `client_secret` must be the same ones you entered in the Playground.
A refresh token only works together with the client that created it.

> The `refresh_token` works like a password to that account's Drive. Keep it private.
> You can close the Playground now; nothing is stored there.

<details>
<summary><b>Alternative: create the token with the Python script instead</b></summary>

Use this only if you'd rather not use the Playground. It needs Python 3.10+ on your laptop.

1. In Step 1.4, choose **Desktop app** instead of Web application (no redirect URI needed),
   download the JSON and rename it `client_secret.json`.
2. Unzip the project, put `client_secret.json` inside the `qbg-doc-formatter` folder, and open a
   terminal there (Windows: type `cmd` in the folder's address bar; Mac: `cd ` then drag the folder in).
3. Run:
   ```bash
   pip install google-auth-oauthlib
   python scripts/get_refresh_token.py client_secret.json
   ```
4. Sign in with the pw.live account in the browser window that opens. The terminal prints the
   ready-made `[google_oauth]` block.

</details>

---

### Step 3: Put the code on GitHub (5 min)

**3.1 Create an empty repository**

1. Go to <https://github.com/new>.
2. **Repository name:** `qbg-doc-formatter`.
3. Choose **Private** (recommended; the app link can still be public, see Step 5).
4. Leave everything else unticked. Click **Create repository**.

**3.2 Upload the files** (choose one way)

**Option A: in the browser (no software needed)**

1. Unzip `qbg-doc-formatter.zip` on your laptop (Windows: right-click → **Extract All**; Mac: double-click).
2. On the new repository page, click **uploading an existing file**.
   **Drag** the items onto the page. Don't use "choose your files": that file picker can't select
   folders, so the `qbg` and `assets` folders would be left out and the app won't start.
3. Open the `qbg-doc-formatter` folder on your laptop and select **everything inside it**
   (`app.py`, `requirements.txt`, `README.md`, and the folders `qbg`, `assets`, `scripts`, `tests`, `.streamlit`).
   Drag them all into the GitHub page.
   - **Do not** upload `client_secret.json` (if you made one) or your Notepad file with the token.
   - **Mac:** `.streamlit` is hidden. Press **Cmd + Shift + .** in Finder to show it.
   - If the folders don't upload and every file lands at the top level, that's fine: the app
     works in that layout too. Then add the theme file by hand: **Add file → Create new file**,
     name it `.streamlit/config.toml` (typing `/` creates the folder), paste its contents, **Commit**.
4. Wait until every file is listed, then click **Commit changes**.
5. Check the repository page shows `app.py` and `requirements.txt` **at the top level**,
   not inside an extra `qbg-doc-formatter` folder. If they are inside a folder, see Troubleshooting.
   (The `qbg`, `assets`, `scripts` and `tests` folders are optional as folders: their files may also
   sit directly next to `app.py`.)

**Option B: with Git (if you have it installed)**

In the terminal, inside the project folder:

```bash
git init
git add .
git commit -m "QBG Doc Formatter"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/qbg-doc-formatter.git
git push -u origin main
```

Replace `YOUR-USERNAME` with your GitHub username. `client_secret.json` and
`secrets.toml` are listed in `.gitignore`, so Git skips them automatically.

---

### Step 4: Deploy on Streamlit Community Cloud (5 min)

1. Go to <https://share.streamlit.io> and sign in with **GitHub**.
   The first time, allow Streamlit to access your repositories (including private ones).
2. Click **Create app** (top right) → **Deploy a public app from GitHub**.
3. Fill in:

   | Field | Value |
   |---|---|
   | Repository | `YOUR-USERNAME/qbg-doc-formatter` |
   | Branch | `main` |
   | Main file path | `app.py` |
   | App URL | pick a name, e.g. `qbg-doc-formatter` → your link becomes `https://qbg-doc-formatter.streamlit.app` |

4. Click **Advanced settings**.
   - **Python version:** `3.12` (3.11 also works).
   - **Secrets:** paste the text below, replacing the bottom three lines with the
     block you copied in Step 2.4:

   ```toml
   ROOT_FOLDER_ID = "1gzMrWz-I7yrqW6S_viabr5QI4G9QY8QL"
   MAX_WORKERS = 12

   [google_oauth]
   client_id = "PASTE-YOUR-client_id-HERE"
   client_secret = "PASTE-YOUR-client_secret-HERE"
   refresh_token = "PASTE-YOUR-refresh_token-HERE"
   ```

   What the two top lines mean (no need to change them):
   - `ROOT_FOLDER_ID`: the ID of your Drive folder, taken from its link.
   - `MAX_WORKERS`: how many question docs upload at the same time.

5. Click **Save**, then **Deploy**.
6. The first deploy takes 2–4 minutes while it installs packages. When it finishes, the app opens.
   There should be **no** yellow "Google credentials are not configured" banner.

---

### Step 5: Make the link public (1 min)

1. On the app page, click **Share** (top right). Or from <https://share.streamlit.io>, click **⋮** next to the app → **Settings → Sharing**.
2. Under **Who can view this app**, choose **This app is public and searchable**.
3. Copy the app link and share it with the team. Anyone with the link can use it without logging in.
   They never see your Google login; it stays inside Streamlit's secrets.

---

### Step 6: Test it (2 min)

1. Open the app link.
2. Upload `tests/sample_test_paper.docx` from the project folder.
3. Click **Check parsing**. You should see **121 found, 120 ready, 1 need attention**.
   The one flagged question is the last one in that paper, which has no explanation in the source.
4. Click **Format and upload** and wait for it to finish.
5. Open your Drive folder and check:
   - `Input/Judiciary/` has the original file.
   - `Output/Judiciary/` has the Excel and a folder of `QBGDOC…docx` files.
6. If this was only a test, you can delete the test files from Drive. **Do not** delete
   `Output/_serial_counter.json` unless you also want numbering to restart.

Deployment is done.

---

## 4. How to use the app (for the team)

1. **Add Word files:** drag in one or many `.docx` / `.doc` files.
2. **Name each file:**
   - **QBG Category:** decides the folder in Drive (e.g. `Judiciary`). Type it once at the top and click **Apply to every row**, or edit rows individually.
   - **File Name:** becomes the Excel name and the folder name for that file's questions. It starts as the uploaded file's name; change it if needed.
3. **Check parsing** (optional but recommended): instant, uploads nothing. Lists every question that needs attention with its question number and the reason.
4. **Format and upload:** shows live progress. Click **Stop processing** at any time.
5. When it finishes, click **Download Excel** or **Download CSV**. The same Excel for each file is also saved in Drive.

What the Status column means:

- **Success:** the question doc was created; its link is in column E.
- **Success - Solution Written Automatically:** the source had no explanation, so the solution
  was written as *"The correct answer is &lt;correct option text&gt;."* (Hindi: *"सही उत्तर है:
  &lt;option text&gt;"*). For English + Hindi questions with only one explanation, the missing
  side is written this way and Remarks says which.
- **Failure:** no doc was created. Column G says why, starting with the source question number, for example `Source Q.23: answer key ('Ans.' / 'Answer:') not found`. Fix that question in the Word file and upload it again.
- **Failure – Not processed: run stopped by user:** you pressed Stop before this question was reached.

Counts shown when you check and run:

| Count | Meaning |
|---|---|
| Bilingual · English / Bilingual · Hindi | Questions with both versions. The two numbers should match; if they don't, the tiles turn red and the questions missing a translation are marked Failure. |
| English-only / Hindi-only | Single-language questions, e.g. the English and Hindi sections of a language paper. |
| Solutions auto-written | Questions whose solution was written from the correct option. |

Rules that keep results accurate:

- Number every question (`1.`, `2.`, … or `Q1.`, `Q2.`, …).
- Label options `(a) (b) (c) (d)`, `a) b)`, `A. B.` or `(A) (B)`.
- Give every question an answer line, for example `Ans. (b)` or `Answer: B`.
- Put the explanation after the answer line.

Section headings, instructions and title pages are ignored automatically.

---

## 5. Updating the app later

- **Changing the code:** edit or re-upload files in the GitHub repository. Streamlit picks up the change and redeploys automatically within a minute or two.
- **Changing secrets** (for example, a new token): <https://share.streamlit.io> → **⋮** next to the app → **Settings → Secrets** → edit → **Save**. The app restarts by itself.
- **Making a new token:** repeat Step 2, then replace the three lines in Secrets.
- **Changing upload speed:** change `MAX_WORKERS` in Secrets. Higher is faster; if Google slows you down, the app waits and retries automatically. 8–16 is a sensible range.

---

## 6. Troubleshooting

| What you see | What to do |
|---|---|
| Yellow banner: **Google credentials are not configured yet** | Secrets are missing or misspelled. Check that `[google_oauth]` is on its own line and that each value is in "double quotes". |
| Playground shows **Error 400: redirect_uri_mismatch** | In Step 1.4 the redirect URI must be exactly `https://developers.google.com/oauthplayground`, with no slash at the end, on a **Web application** client. Saving can take a few minutes to apply. |
| Playground shows **invalid_client** or **unauthorized_client** | The Client ID or secret in the gear settings is wrong or has extra spaces. Paste them again. |
| No refresh token in the Playground | Gear icon → **Access type: Offline**, then repeat Step 2.2. |
| Token stopped working after about 24 hours | **Use your own OAuth credentials** wasn't ticked in the Playground. Redo Step 2 with it ticked. |
| `unauthorized_client` in the app logs | The `client_id` / `client_secret` in Secrets don't belong to the client that made the token. Use the same pair as in the Playground. |
| Browser shows **Access blocked: … can only be used within its organization** | You signed in with a non-pw.live account. Sign in with the pw.live account. |
| `invalid_grant` when the app runs | The token was revoked or expired. Repeat Step 2 and update Secrets. Make sure the consent screen is **Internal**. |
| `File not found` / 404 for the root folder | The account you signed in with in Step 2 can't open the Drive folder. Use the account that owns it, or give that account Editor access to the folder. |
| `storageQuotaExceeded` | You used a service account instead of Step 2's login. Use Step 2, or see the service-account note in section 8. |
| **ModuleNotFoundError** on `from qbg.google_services import …` | You're running an older `app.py`. The current version works whether the files are inside `qbg/` and `assets/` or all directly next to `app.py`. Upload the latest `app.py`, `builder.py`, `parser.py` and `pipeline.py`. |
| Red box: **Some project files are missing** | Upload the files it lists next to `app.py`. |
| App looks dark or colours look wrong | `.streamlit/config.toml` is missing. In GitHub: **Add file → Create new file**, type `.streamlit/config.toml` as the name (the `/` creates the folder), paste the contents of that file from the zip, **Commit**. |
| Streamlit error: **No module named …** or the app won't start | `requirements.txt` must be at the top level of the repository, next to `app.py`. |
| Files ended up inside an extra `qbg-doc-formatter/` folder on GitHub | Either move them up a level, or in Streamlit set **Main file path** to `qbg-doc-formatter/app.py`. |
| App shows **"This app has gone to sleep"** | Free Streamlit apps sleep after a period with no visitors. Click **Yes, get this app back up**; it wakes in under a minute. |
| A question says **options … could not be identified** | Its options don't start with a recognisable label. Fix the labels in the Word file and re-upload. |
| Remarks say **answer keys in file = X, questions detected = Y** | One question boundary is unusual. The other remarks point to which question. |

---

## 7. How questions are recognised

| Part | Forms the app understands |
|---|---|
| Question start | `1.` `1)` `1:` `Q1.` `Q.1` `Question 1:` `13 'Admission'…`, typed or Word auto-numbered |
| Options | `(a)` `a)` `a.` `(A)` `A)` `A.`; also `(1)` `1)` when there are no letter options; options laid out in a table |
| Answer | `Ans. (d)` `Ans.b` `Ans.- c` `ANs. (d)` `Answer: B)` `Answer: C) Article 17` `Correct Answer: b) …` `Answer: 2` `उत्तर: (ख)` `Ans. ©` |
| Solution | Everything after the answer line up to the next question, including an explanation written on the answer line itself; labels `Explanation:` `स्पष्टीकरण:` `व्याख्या:` are removed |
| Hindi options | `(a)` … as above, or `(क) (ख) (ग) (घ)`; Devanagari digits `(१)` are understood |

Supported paper layouts:

| Layout | Example | Output |
|---|---|---|
| English only | Original test papers, DPPs | Single-column, English |
| English + Hindi, Hindi copy numbered | BIHAR APO: Q, options, `1. हिंदी प्रश्न`, options, answer, English explanation, `स्पष्टीकरण:` Hindi explanation | Bilingual template |
| English + Hindi, Hindi copy not numbered | TEST-13 Full Length: same, Hindi question without a number; options `1. 2. 3. 4.` + `(5) Question not attempted`, `Answer: 2` | Bilingual template (all 5 options kept) |
| Language paper | English questions 1–50, then Hindi questions numbered again from `Q1.`, `उत्तर:` / `व्याख्या:` | English questions single-column English; Hindi questions single-column Hindi |
| No explanations | BNSS consolidated DPPs | Solution auto-written, status *Success - Solution Written Automatically* |

Language is decided by script only: Devanagari is **Hindi**, everything else **English**.
A Hindi line full of English legal terms (e.g. *"बीएसए के अंतर्गत- Falsus in uno … है?"*) is still read as Hindi.

Special cases handled:

- **Statements labelled like options** (for example statements (a), (b), (c) followed by options (a)–(d)): the app finds the options by reading *upwards* from the answer line, so the statements stay in the question.
- **Numbered statements** (`1.`, `2.`, `I.`, `II.`) inside a question are not mistaken for new questions.
- **Numbered lists inside an explanation** are not mistaken for the next question.
- **Match-the-following tables** are kept as tables in the question doc. Bold, italic, underline and images are kept too.
- **Headings between questions** (for example `COMPETITION ACT`) are dropped.
- **A question with no number** is still detected and flagged in Remarks for checking.
- **Numbering that restarts** (a new section starting again at 1) is followed.
- **Chapter titles in Word Heading style** between questions are dropped (a question that is itself
  styled as a heading is kept).
- **Word quirks fixed automatically:** `©` / `(©)` typed for (c); an option inside a bullet list;
  Word lists that keep counting (options shown as e–h) are read in order and flagged *please verify*.
- **Flagged for a person to fix:** options labelled (a), (b), (a), (b); options with identical text;
  a question missing its Hindi or English version inside a bilingual paper.

A question is marked **Failure** when its options, answer line, question text or explanation is missing, or when the answer letter doesn't match any option.

---

## 8. Technical reference

### Speed

Reading and formatting are done in memory and take about a second per 100 questions.
The rest of the time is uploading to Drive: one upload per question doc (plain `.docx`,
no conversion), running `MAX_WORKERS` at a time, with automatic retry if Google throttles.

### Serial numbers

- Stored in `Output/_serial_counter.json`.
- Each run reserves its whole range there **before** uploading anything, so a stop or crash never reuses a number.
- Two people running batches at the same time never get the same numbers.
- If the counter file is deleted, the app continues from the highest `QBGDOC` file name it finds in Drive.
- Re-running the same File Name creates new serials. The new Excel gets the run ID added to its name, so the earlier record is kept.

### Stop button and page refresh

- **Stop:** uploads already in progress finish. Every remaining question is written to the Excel as "Not processed", so every serial is accounted for.
- **Refreshing the page** during a run doesn't stop it; the page reconnects (the run ID is kept in the URL).
- If Streamlit itself restarts the app mid-run (rare), that run ends. Files already uploaded stay in Drive, and the serial counter stays correct.

### Using a service account instead of Step 2

Only works if the root folder is inside a **Shared Drive**. Service accounts have no
storage of their own, so uploads into a normal My Drive folder fail with
`storageQuotaExceeded`. Add the service-account email to the Shared Drive as
**Content manager**, then use this Secrets block instead of `[google_oauth]`:

```toml
ROOT_FOLDER_ID = "your-shared-drive-folder-id"
MAX_WORKERS = 12

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```

### Running on your own laptop

```bash
pip install -r requirements.txt
```

Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill it in as in Step 4. Then:

```bash
streamlit run app.py
```

The app opens at <http://localhost:8501>.

### Checking a new paper style without Google

```bash
python scripts/parse_local.py "new_paper.docx" out_folder --category Judiciary
```

Writes one `.docx` per question to `out_folder` and prints any question that needs attention.

### Tests

```bash
python tests/test_parser.py     # 4 real papers + synthetic layouts (bilingual, Hindi, auto-solution, ...)
python tests/test_pipeline.py   # Drive layout, serial numbers, stop mid-run (simulated Drive)
```

### Project structure

```
app.py                          the web app (Streamlit)
requirements.txt                Python packages Streamlit installs
qbg/docx_reader.py              reads Word files (auto-numbering, line breaks, tables, images)
qbg/parser.py                   finds questions, options, answer, explanation
qbg/builder.py                  writes each question into the QBG format
qbg/google_services.py          Google Drive: login, folders, uploads, serial counter
qbg/pipeline.py                 runs a batch: progress, stop, Excel/CSV
assets/question_template.docx   single-language QBG layout (English or Hindi)
assets/question_template_bilingual.docx   English + Hindi QBG layout
assets/creator.png              creator badge photo
scripts/get_refresh_token.py    optional Drive login helper (alternative to the Playground)
scripts/parse_local.py          offline checker
tests/                          automated tests + sample paper
.streamlit/config.toml          colour theme
.streamlit/secrets.toml.example secrets template (the real secrets.toml is never uploaded)
```
