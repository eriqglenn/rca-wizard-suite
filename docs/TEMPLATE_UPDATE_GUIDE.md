# Updating a City template

When the City revises one of its Word templates (new director, new wording,
new letterhead), you no longer need to keep the exact old file name or edit
any code. Every wizard reads its template list from
`Wizard_Files\templates\manifest.json`, and the **Manage Templates** screen
edits that file for you.

## The 3-minute way (Manage Templates)

1. Start the suite and click **Manage Templates** on the home screen.
2. Pick the wizard tab (MLS, MBL, MLS Renewal, MBL Renewal).
3. Select the document row (for example *Director approval*) and click
   **Replace selected template...**, then choose the new `.docx`.
   - The old file is moved to `templates\<wizard>\_previous\` with a date
     stamp. Nothing is deleted.
   - The new file keeps its own name; the manifest now points at it.
4. Say **Yes** to *Run check*. The suite generates every document for that
   wizard from the bundled sample workbook and checks for unfilled
   placeholders, broken footers, protected output, wrong dates, and so on.
   - **PASS**: the template is live. Click **Generate sample documents...** and
     open the folder it shows to eyeball the wording in Word.
   - **FAIL**: read the `[FAIL]` lines in the results panel. You can put the
     old template back with one click (**Restore previous version**).

The wizard picks up the new template the next time it is opened. Planners
who had a template path saved in their session memory get the new default
automatically, because the old file no longer exists at that path.

## Two kinds of templates

The *Filled by* column tells you what a revised template must preserve.

### `tokens` (MBL ordinance)

The wizard replaces `{{Token_Name}}` placeholders. A revised template must
keep the **same token names**. Adding a token the wizard does not know makes
it print literally (`{{New_Thing}}`), and the check will flag it as
"unfilled placeholder". Removing a token is harmless.

Type each token in one motion in Word. If Word splits a token across runs
(a spellcheck squiggle in the middle often does this) it will not fill; the
Manage Templates screen warns *"a token is split across Word runs"*. Retype
the token and replace again.

### `wording` (MLS 2026 templates, MBL letters and forms, both renewal wizards)

The 2026 MLS Renewal templates and the exact phrases their code matches are
listed in `docs\mls_renewal\TEMPLATES_2026.md`; the 2026 MLS templates in
`docs\mls\TEMPLATES_2026.md`.

These wizards find the template's own placeholder wording and replace it,
for example `XX Block of XX, X side,`, `Planner Name@houstontx.gov`,
`Month XX, 2026`, `ADDRESS BLOCK`, `Insert Map`. A revised template must keep
those exact phrases. Change the surrounding wording freely; do not rewrite
the placeholder phrases themselves.

Because the check cannot know what the City *meant* to write, always finish
a wording-template update with **Generate sample documents...** and read the
output in Word. Any placeholder phrase still visible in the output means the
new template changed that phrase and the code needs a matching edit (this is
the one case that still requires a developer; see the developer manual,
section 3).

## Manual way (no Manage Templates)

1. Copy the new `.docx` into `Wizard_Files\templates\<wizard>\`.
2. Open `Wizard_Files\templates\manifest.json` in Notepad and change the
   `"file"` value of that role to the new file name.
3. From `Wizard_Files`, run `python tools\validate.py --wizard <key>` (or
   `"RCA Wizard Tools.exe" validate --wizard <key>` in the packaged build).

If `manifest.json` is deleted or damaged, the suite rebuilds it from the
built-in defaults on the next start, so the manual way can never brick a
wizard. The historical file names remain a fallback: if the manifest names a
file that is missing, the wizard uses the original file name if it exists.

## Wizard keys and roles

| Wizard | Key | Roles (manifest `key`) |
|---|---|---|
| MLS | `mls` | template, protest_template, hearing_template, applicant_template, director_template, pc_template, routing_template, staff_report_template, ordinance_template |
| MBL | `mbl` | template, hearing_template, protest_template, reschedule_template, applicant_template, director_template, pc_template, routing_template, staff_report_template, staff_report_less51_template, staff_report_protest_template, ordinance_template |
| MLS Renewal | `mls_renewal` | sfr_template, no_sfr_template, protest_template, applicant_template, director_template, pc_template, routing_template, staff_report_director_template, staff_report_pc_template, ordinance_template |
| MBL Renewal | `mbl_renewal` | owner_template, protest_template, reschedule_template, acceptance_template, applicant_template, director_template, pc_template, routing_template, staff_report_template, ordinance_template |

## Per-case overrides are unchanged

Every wizard still has Browse buttons on its Select Files page. Those choose
a template for one case only and never change the manifest.
