# Game Accessibility Testing Template — Project Context

Context for Claude Code sessions in this folder. Last updated 2026-08-23.

## What this is

A screen-reader-friendly Excel template for logging accessibility bugs in game
accessibility mods. Built for blind/low-vision testers using NVDA/JAWS, and intended to
be published open-source on GitHub for other mod creators. Author: Christopher Tavarez.

## Layout

The publishable repo is a **subfolder**; working files stay outside it so they don't ship.

```
Accessibility Templits\              <- working dir, NOT the repo
├── CLAUDE.md                        <- this file; deliberately outside the repo
└── Game Accessibility Testing Templit\   <- THE REPO
    ├── Game Accessibility Testing templit.xlsx
    ├── rebuild_template.py
    ├── README.md
    └── LICENSE
```

The user moved `CLAUDE.md` out of the repo folder on purpose — Claude's project memory is
not meant to be published. Keep it out, and don't write new working files into the repo
subfolder without asking.

| File | Role |
| --- | --- |
| `rebuild_template.py` | **Source of truth.** Generates the entire workbook. Now inside the repo, so it ships publicly — that's what lets contributors propose changes as readable diffs instead of binary spreadsheets. |
| `Game Accessibility Testing templit.xlsx` | Generated output. Do not hand-edit — regenerate instead. |
| `...\README.md` | Finished and published. Fully written (collaboratively), no placeholder brackets or guidance comments remain. |
| `...\LICENSE` | CC BY 4.0, downloaded verbatim from Creative Commons. Do not regenerate from memory. |

The original pre-rework draft (`... - BACKUP.xlsx`) is no longer present; the user removed
it after the rework was verified.

**The `.xlsx` is a build artifact.** Any change to the workbook means editing
`rebuild_template.py` and re-running it, never editing cells directly — a manual edit is
silently destroyed by the next rebuild.

```
python rebuild_template.py     # idempotent; safe to re-run against its own output
```

Run it from inside the repo folder — `SRC` is a bare filename resolved against the current
directory, and the script now sits next to the workbook. Verified working from there.
Running it from the parent fails fast with `FileNotFoundError` and writes nothing, so a
misplaced run is noisy rather than destructive.

Fails with `PermissionError` if the workbook is open in Excel. Close it first.

The file on disk is a clean rebuild from the current script (last rebuilt 2026-08-23,
after fixing the WCAG/GAG on-sheet instructional text — see Design decisions below). No
outstanding drift between script and shipped file as of that rebuild. Still confirm with
the user before rebuilding again, since they may have opened and re-saved it from Excel
since, which would make the file Excel's normalized version rather than byte-identical
script output again.

## Workbook structure

8 sheets: `Setup`, `Game Sections`, `Severity Scale`, `Visual`, `Screen Reader`,
`Keyboard`, `Controller`, `Summary`.

The four log tabs share one 14-column layout and 2,000 pre-built rows. Dropdowns are
driven by 10 workbook-scoped named ranges (`GameSections`, `SeverityLevels`,
`StatusList`, `PlatformList`, `InputMethodList`, `ScreenReaderList`, and four
`*ErrorTypes`).

## Design decisions and why

These were deliberate; don't "fix" them without checking with the user first.

- **No merged cells, no text boxes, no split headers.** Merged cells break screen-reader
  cell navigation. This is the core constraint of the whole project.
- **Every data grid is a real Excel Table.** That's what makes NVDA/JAWS announce the
  column header when moving down a row. Losing the Table to gain formatting flexibility
  is a bad trade here.
- **Short column headers** (`Game Section`, not `Which specific part of the game are you
  testing?`). A long header is re-read on every cell. Original question wording is kept
  as a cell comment on each header so context isn't lost.
- **Status and severity are explicit text, never color-only.**
- **Issue ID is `=IF(C{row}="","","V-"&TEXT(ROW()-1,"000"))`** so unused rows render
  completely blank. Triggered by column C (Game Section) because that's the first field
  a tester fills.
- **Severity values are self-describing** (`1 - Cosmetic` … `5 - Blocking`) so a screen
  reader announces meaning, not a bare digit.
- **`Severity Scale` ships fully unprotected, on purpose.** Teams rename or restructure
  the severity levels themselves — this is about impact wording, not an outside
  framework. Every other sheet is protected with headers/formulas locked, entry cells
  unlocked, no password.
- **WCAG / Game Accessibility Guidelines terminology belongs on the Setup tab's dropdown
  lists (mainly the Error Type lists), not the Severity Scale.** Those frameworks
  categorize the *kind* of issue, not its severity. Those guidelines are copyrighted, so
  their wording is deliberately NOT shipped — users paste their own terms into the
  unlocked list cells below each dropdown header on Setup.
- **Summary counts per-tab first, then sums.** Each tab×severity cell is an isolated
  single-sheet formula, so one mistyped tab name breaks only its own row. An earlier
  design used one shared cross-sheet formula where a single typo silently zeroed every
  total — don't go back to that.

## Excel/openpyxl traps found the hard way

Each of these was a real shipped bug, caught by testing in actual Excel. They are not
theoretical.

- **`COUNTA` counts a formula returning `""` as non-blank.** Counting the Issue ID column
  reported 2,000 issues on an empty sheet. Count column C instead.
- **`COUNTA`/`COUNTIFS` over `INDIRECT` to a nonexistent sheet does NOT raise a catchable
  error** — it treats the `#REF!` as one item, so `IFERROR` never fires. Validate the
  sheet name with `ISREF(INDIRECT(...))` instead.
- **Apostrophes in sheet names must be doubled** inside a quoted `INDIRECT` reference, or
  a legitimately named tab (`Ava's Tab`) reads as invalid. Wrap with
  `SUBSTITUTE(name,"'","''")`.
- **openpyxl defaults `showErrorMessage=False`.** Data validation then looks configured
  but silently accepts anything typed. Must pass `showErrorMessage=True` explicitly.
- **Clearing `cell.value` does not clear styling.** Rebuilding in place left ~103
  orphaned styled cells and a bloated used range. `reset_sheet()` now deletes and
  recreates each sheet.
- **Row insertion inside an Excel Table is blocked while the sheet is protected.** The
  working procedure (verified) is: unprotect → type in the row below the table (it
  auto-expands, dropdowns carry) → copy the Issue ID cell down (it does not auto-fill).
- **F2 breaks dropdowns for screen readers.** F2 enters text-edit mode where arrow keys
  stop responding to the list. Correct keystroke is Alt+Down Arrow. This is standard
  Excel behavior, not a template defect; it's documented on the Setup tab.

## Verification approach

openpyxl only shows formula *text*. Formula *results* must be checked by driving real
Excel via PowerShell COM. Several bugs above passed inspection and failed live.

```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("<absolute path>")
# write test values, then $excel.CalculateFullRebuild(), then read .Text
$wb.Close($false)   # ALWAYS close without saving
$excel.Quit()
```

Always `Close($false)` so test data never lands in the real file, and always `Quit()` in
a `finally` block to avoid orphan Excel processes. PowerShell here is 5.1 — no `&&`, no
ternary.

## Status

**Published.** Live at https://github.com/ChrisTavar2022/Game-Accessibility-Testing-Templit
as of 2026-08-23. Three subagent audit passes (2026-08-22 through 2026-08-23) found and
fixed bugs across the workbook, script, and README, the last two with live Excel/COM
verification, not just static reads. Final pass came back clean: no blocking issues, no
open "should fix" items. Confirmed clean on: merged cells (none), table refs, named
ranges, validation wiring across all 2,000 rows, protection behavior (live-tested, not
just XML-flag reads), idempotency (byte-identical rebuilds except the modified
timestamp), and README/workbook consistency (the on-sheet instructional text and the
README were reconciled after they were found to contradict each other on where
WCAG/GAG terminology belongs — see Design decisions above).

## Where we left off (2026-08-23)

Shipped. README is fully written, uploaded via WSL/git, and the repo URL is filled into
the License section's attribution string. Nothing outstanding on this project unless the
user brings something new: a bug report, a contribution to review, a new feature, or
LibreOffice/Google Sheets compatibility testing (still unverified — don't claim
compatibility without testing it).

He uses this file as the between-session channel: read it first, treat anything he's
added as current instructions, and keep it accurate as things change.

## Open items for the user

None outstanding. Everything from the original list is done: README written, repo
created and pushed, `rebuild_template.py` shipped in the public repo, contribution
policy decided (PRs against the script only, not the `.xlsx`, since binary files don't
diff/merge in git; issues for everything else). LibreOffice/Google Sheets compatibility
remains untested and unclaimed, which is a deliberate honesty choice, not a gap to fill
before shipping.

## Working preferences

- The user wants improvements to their existing design, not replacements. Ask before
  restructuring something they built.
- They catch and care about correctness issues; verify claims in real Excel before
  stating them, and say plainly when something failed.
- Don't add their email address to any published file. Name-only attribution is intentional.
