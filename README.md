# K-Map Quiz: From Zero to Expert

Interactive Karnaugh-map quiz for Digital Logic Design (EED, UET Lahore).
Students draw groups on the K-map, write the minimal equation, and get marked automatically.
All marks go into one Google Sheet.

- 5 levels × 3 questions, 10 marks each (150 total)
- Marking per question: covers every 1 (2) · groups as large as possible (2) · no extra groups (1) · equation correct (3) · equation minimal (2)
- Random maps for every student, one submission per question, progress kept if the page is reloaded

## Setup (one time)

### 1. Google Sheet for marks
1. Create a new Google Sheet (e.g. "K-Map Marks").
2. **Extensions → Apps Script**. Delete the sample code and paste everything from `apps-script/Code.gs`. Save.
3. **Deploy → New deployment** → type **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Press **Deploy**, allow the permissions, and copy the **Web app URL** (ends in `/exec`).

### 2. Connect the quiz to the sheet
In `index.html`, find this line and paste your URL:
```js
const SHEET_URL='PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
```
Commit the change.

### 3. Turn on GitHub Pages
Repo **Settings → Pages** → Source: **Deploy from a branch** → Branch: **main**, folder **/ (root)** → Save.
The quiz link will be `https://msm-official.github.io/kmap-quiz/`.

## In class
- Share the link with students. No accounts needed.
- Marks appear in the **Marks** tab of your Google Sheet, one row per student, updating after every question.
- If a student had no internet at the end, they get a **result code**. Open
  `https://msm-official.github.io/kmap-quiz/#teacher`, paste the codes, and press **Send to sheet**.

## Notes
- Progress is stored in the student's browser. If they clear browser data or switch devices, they start a new attempt, which shows up as a new row.
- Submissions carry a checksum to stop casual tampering. This is not strong security, so the quiz suits practice and class activity rather than high-stakes exams.
