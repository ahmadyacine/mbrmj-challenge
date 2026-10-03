# تحدي مبرمج — "First Day Challenge"

Files: `index.html` (the whole quiz, one file) · `backend.gs` (Google Apps Script) · `logo-black.png` (source logo; a small copy is embedded in the page) · `test_code_facts.py` (verifies the code facts used in Q5–Q8).

## 1. Deploy the backend
1. Create a Google Sheet → **Extensions → Apps Script** → paste `backend.gs`. (If you created the script from script.google.com instead, set `SHEET_ID` at the top of the file to the Sheet's ID.)
2. Run `setup()` once and authorize. It creates `Events`, `Sessions`, `Answers`, `Leads`, `Dashboard`.
3. **Deploy → New deployment → Web app** → Execute as **Me** → Access **Anyone**. Copy the `/exec` URL.
4. After any later edit of the script: Deploy → Manage deployments → Edit → **New version**.

## 2. Connect the page
Open `index.html`, find `const CONFIG` near the top of the `<script>`, and fill:
`APPS_SCRIPT_URL` (from step 1), `WHATSAPP_NUMBER` (e.g. `9715XXXXXXXX`, digits only), `WHATSAPP_CHANNEL_URL`, `WAITLIST_URL`, `CHECKLIST_PDF_URL`, `INTRO_VIDEO_URL` (optional), `SITE_URL`, and `RESOURCES` (one learning link per stage, shown in the report).
Buttons whose URL is empty are hidden. Without `APPS_SCRIPT_URL` the quiz works fully, events stay queued in the browser.

## 3. Host on mbrmj.ae
It's static: upload `index.html` to `mbrmj.ae/challenge/` (any static host, e.g. the site's hosting, Netlify, GitHub Pages). No build step.

## 4. Links
- A/B title: `?v=a` (تحدي أول يوم دوام, default) · `?v=b` · `?v=c`
- Source/club: `?ref=aus-cs-club` (saved in Sessions and Leads)
- UTM: `?utm_source=instagram&utm_medium=story&utm_campaign=oct`
- Combined: `https://mbrmj.ae/challenge/?v=b&ref=aus-cs-club&utm_source=whatsapp`

## 5. Editing content
Everything editable is at the top of the script (`CONFIG`, `TITLES`, `QUESTIONS`, `GLOSSARY`, `REPORT`, `STORY`). Inline markup: `**bold**`, `` `code` ``, `{{Glossary Term}}` (tap-to-explain; the term must exist in `GLOSSARY`). If you change questions or options, bump `STORAGE_KEY` to `mbrmj_challenge_v2` so old saved progress is ignored, and update `QUESTION_META` in `backend.gs` (used for the trap-rate table), then re-run `setup()`.

## 6. Reading the data
- **Sessions**: one row per visitor session (status, last step, furthest question, stage statuses, contact left).
- **Answers**: one row per answer, with original option ids, confidence and time.
- **Leads**: contacts (and the report email is sent if an email was given).
- **Dashboard** (formulas, re-run `setup()` to rebuild): funnel, drop-off per question, % correct, % "sure but wrong" per question, average time, stage statuses, and trap rate per option — each split by variant (all / a / b / c). To split by `utm_source` or `ref`, copy a Dashboard block and add a `Sessions!F:F` / `Sessions!I:I` criterion to the COUNTIFS.
- **Events**: raw log (including `abandon`, `tooltip_open`, `explain_expand`, share/PDF clicks).

## Notes
- Report email: the page builds the report HTML and the script wraps it; mail is sent from the Google account that deployed the script (Apps Script daily mail quota applies).
- Share card: 1080×1920 PNG made on a `<canvas>`; native share on phones, download on desktop.
- Privacy: no IPs, no third-party trackers; contact only after the consent checkbox.
- Tests: `python test_code_facts.py` checks the bug/fix behaviour claimed in Q5–Q8.
