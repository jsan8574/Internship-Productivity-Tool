# Verify Desk

Insurance Benefits & Eligibility Verification tool for interns working live accounts. It is both the daily production tool and a learning tool:

- **Interns**: personal daily queue, an 8-step guided verification with how-to prompts and skill checklists (skippable per step, or all at once with Express mode), Decision-Matrix status suggestions, auto-generated TEBRA note, and a My Progress page.
- **Managers**: upload the day's Excel/CSV (Provider Name, Patient Name, DOB, Type of Service) or paste from Sheets, split it between interns, track the Team Board, reassign work.
- **QA**: pass/fail with common findings; failed accounts go back to the intern.
- **Status pre-selection**: the approval status fills in from the Decision Matrix as the intern answers; they confirm or change it. Each record stores STATUS SOURCE (Decision Matrix / Override).
- **Time study (managers)**: active minutes per section are recorded on every account (pauses after 3 idle minutes). The Time Study page shows a heatmap of interns × sections and flags sections taking 1.5× the rest of the team.
- **Team & PINs (managers)**: upload or paste the intern list (Name, Email, Role). New people get a PIN shown once to copy; active interns feed the Upload & Assign dropdowns. Reset PINs and deactivate people from the same page.
- **Export (managers)**: filter by date range, intern, work status, approval status or QA result, then download CSV / Excel or create a Google Sheet copy in the owner's Drive. Every export is written to the Activity Log.
- **Knowledge Base**: searchable definitions, Decision Matrix, statuses, payer call script, plan types, service rules, documentation standards. Open it from the sidebar or press `/`.

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
4. **Add your team.** Easiest: open the web app after deploying and use **Team & PINs** to upload your intern list. Or fill the Roster tab by hand: `setup` adds you as Manager. Add each person: NAME, EMAIL, ROLE (`Intern`, `Manager`, or `QA`), ACTIVE (`Yes`). Leave PIN blank and run `setup` again: it gives everyone without a PIN a random 6-digit PIN. Send each person their PIN privately.
5. **Deploy.** Deploy > New deployment > Web app. Execute as: **Me**. Who has access: **Anyone**. Share the URL with interns. They sign in with their email and PIN, and never need access to the Sheet.
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
- Sign-in uses email + PIN, so it works with personal Gmail accounts or no Google account at all. Sessions last 8 hours; 5 wrong PINs lock that email for 15 minutes. To reset someone's PIN, clear it in the Roster and run `setup`. To remove access, set ACTIVE to `No`.
- On a Google Workspace domain you can deploy with "Anyone within your organization" instead; Roster members signed in to Google then skip the PIN.
- Knowledge Base boxes marked "Team rule" are team-specific and should be reviewed by the program manager.

## Data & privacy

- Never commit real patient data to this repository. Keep it private.
- **HIPAA:** a personal Google account is not covered by a Business Associate Agreement (only Google Workspace accounts can sign one). Get approval from your compliance team before putting real patient names or DOBs in a Sheet owned by a personal account. Test with made-up patients until then.
