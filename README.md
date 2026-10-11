# Verify Desk

Insurance Benefits & Eligibility Verification tool for interns working live accounts. It is both the daily production tool and a learning tool:

- **Interns**: personal daily queue, an 8-step guided verification with how-to prompts and skill checklists (skippable per step, or all at once with Express mode), Decision-Matrix status suggestions, auto-generated TEBRA note, and a My Progress page.
- **Managers**: upload the day's Excel/CSV (Provider Name, Patient Name, DOB, Type of Service) or paste from Sheets, split it between interns, track the Team Board, reassign work.
- **QA**: pass/fail with common findings; failed accounts go back to the intern.
- **Knowledge Base**: searchable definitions, Decision Matrix, statuses, payer call script, plan types, service rules, documentation standards. Open it anywhere with the button or `/`.

The Google Sheet is the single source of truth. The app runs as a Google Apps Script web app bound to that Sheet, so patient data stays in your Google Workspace.

## Repository layout

| Path | What it is |
|---|---|
| `src/Code.gs` | Apps Script backend (reads/writes the Sheet, roles, upload, QA) |
| `src/Index.html` | The web app UI (single file). Knowledge Base content is the `KB` array near the top of its script. |
| `src/appsscript.json` | Apps Script manifest (V8, web app: execute as owner, domain access) |
| `template/Verify-Desk-Template.xlsx` | Clean starter workbook: Pt Verif, Decision Matrix, Roster. No patient data. |

> Opened outside Google Apps Script, `src/Index.html` runs in **demo mode** with fictional sample data, so you can preview it in any browser.

## Setup (about 10 minutes)

1. **Create the Sheet.** Upload `template/Verify-Desk-Template.xlsx` to Google Drive and open it with Google Sheets (File > Save as Google Sheets). You can also use your existing verification workbook instead.
2. **Add the code.** In the Sheet: Extensions > Apps Script. Replace `Code.gs` with `src/Code.gs`. Add an HTML file named exactly `Index` and paste `src/Index.html`. (Optional: Project Settings > show manifest, paste `src/appsscript.json`.)
3. **Run setup.** Select `setup` in the function dropdown, click Run, approve permissions. It adds the tool columns, dropdowns, Roster and Activity Log tabs.
4. **Fill the Roster.** NAME, EMAIL (company Google account), ROLE (`Intern`, `Manager`, or `QA`), ACTIVE (`Yes`).
5. **Deploy.** Deploy > New deployment > Web app. Execute as: **Me**. Who has access: **Anyone within your organization**. Share the URL with interns. Interns do not need access to the Sheet.
6. **Updating later.** Paste new code, then Deploy > Manage deployments > Edit > Version: New version. The URL stays the same.

### Optional: deploy from this repo with clasp

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json   # paste the Script ID from Apps Script > Project Settings
clasp push
```

## Things to confirm before go-live

- "Policy Not Active" vs "No self pay agreement" share the same conditions in the Decision Matrix. The tool uses *No self pay agreement* when an agreement was requested but not signed, and *Policy Not Active* when self-pay was not discussed. Adjust `suggest()` in `src/Index.html` if your rule differs.
- Replace the generated TEBRA note (`noteText()` in `src/Index.html`) with your team's exact VAC note format if it differs.
- Interns must sign in with company Google Workspace accounts (personal Gmail accounts don't expose the email to the script).
- Knowledge Base boxes marked "Team rule" are team-specific and should be reviewed by the program manager.

## Data & privacy

Never commit real patient data to this repository. Keep it private.
