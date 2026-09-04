# Game Accessibility Testing Template — Project Context

Context for Claude Code sessions in this folder. Last updated 2026-09-04.

## What this is

A screen-reader-friendly Excel template for logging accessibility bugs in game
accessibility mods. Built for blind/low-vision testers using NVDA/JAWS, and intended to
be published open-source on GitHub for other mod creators. Author: Christopher Tavarez.

## Layout

`CLAUDE.md` lives **inside** the repo folder and ships publicly, on purpose (decided
2026-09-04, reversing an earlier "keep it out" stance from 2026-08-23). Rationale: it
lets other contributors' own Claude sessions pick up full project context to fix issues
themselves, and lets the user keep the repo in sync across multiple machines without a
separate out-of-repo file to carry along.

```
Accessibility Templits\              <- working dir (parent, not the repo)
└── Game-Accessibility-Testing-Templit\   <- THE REPO
    ├── CLAUDE.md                        <- tracked, committed, published
    ├── Game Accessibility Testing templit.xlsx
    ├── rebuild_template.py
    ├── README.md
    └── LICENSE
```

Don't write new working/scratch files into the repo folder without asking — that
guidance still stands even though `CLAUDE.md` itself is now an intentional exception.

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
- **A dropdown source range with trailing blank rows opens scrolled to the bottom on an
  empty cell.** Every dropdown source list here reserves extra blank rows below its seed
  values for growth. Opening a dropdown on a blank cell makes Excel match the first blank
  entry in the source list — which, since the blanks sit at the end, opens the list at the
  bottom instead of the top. This is real Excel behavior (documented independently by
  several Excel sources), not an NVDA bug, and was mistaken for one at first — see
  `dynamic_list_ref()`. Fixed 2026-09-04 by trimming every dropdown named range
  (`GameSections`, `SeverityLevels`, `StatusList`, `PlatformList`, `InputMethodList`,
  `ScreenReaderList`, and the four `*ErrorTypes`) to `INDEX(range,COUNTA(range))` so the
  range contains zero blank cells. Live-Excel-verified (COM): all 10 ranges now stop at
  their last real entry with no trailing blanks. Tradeoff: this assumes entries stay one
  unbroken run from the first data row — a blank row *in the middle* now truncates the
  dropdown right there, so that warning is on the Setup tab, Game Sections tab, and
  Severity Scale tab instructions. **User-confirmed working 2026-09-05** (tested live
  with NVDA on a real copy of the rebuilt file).
- **A hand-set NVDA+Shift+C/R header designation does not survive closing and reopening
  the file.** Tried as a way to pre-configure column-header announcement on the Setup
  tab's Dropdown Lists grid (which wasn't a Table, unlike the log tabs). Reverse-
  engineered the on-disk format from a live NVDA session — pressing NVDA+Shift+C then
  NVDA+Shift+R on one cell and saving writes a single sheet-scoped defined name,
  `Title_<random hex id>`, `RefersTo` pointing at just that one intersection cell (no
  explicit range). Baked the same shape into the script with a fixed id for
  idempotency — looked structurally correct (byte-for-byte matched what NVDA itself
  had written) but **failed live-testing 2026-09-05**: reopening the file did not
  announce headers. Root cause not chased further (not worth it once the real fix was
  found) — abandoned in favor of the option below. If revisited, the open question is
  whether NVDA's own persistence needs something beyond the bare defined name (e.g. a
  fresh in-session keypress each time), not just the file contents.
- **Fix that actually worked: make the grid a real Excel Table, same as the log tabs.**
  Converting `Setup`'s Dropdown Lists grid (all 8 lists bumped to a uniform
  `row_count=15` so the range is a clean rectangle, then wrapped in
  `tbl_DropdownLists`) reuses the exact mechanism already proven elsewhere in this
  workbook, instead of a fragile reverse-engineered defined name. Not yet
  live-tested with NVDA as of this note — the defined-name attempt was verified to
  fail, but the Table replacement itself is still pending an actual screen-reader
  confirmation (only checked structurally via COM so far).
- **Freezing a Table's own header row splits it from its data body into two
  Excel panes, which breaks NVDA's table-region tracking across that
  boundary.** Live NVDA repro 2026-09-04: cursor in the header row, arrow
  Down into the first data row -- NVDA announces "out of table" even though
  it's the same `ListObject`. Root cause: freeze panes literally splits the
  worksheet window into separate panes; wherever the split fell exactly on
  a Table's header/data line, the header row and its data body ended up in
  different panes, and NVDA's table tracking is scoped per pane, so crossing
  that split looks like leaving the table. Checked every sheet for this
  exact pattern (freeze row == Table's own first data row), not just the
  reported one: the four log tabs (`freeze_panes = "A2"`, header row 1) and
  `Game Sections` / `Severity Scale` (each froze at their own Table's first
  data row) all had it; fixed by removing `freeze_panes` entirely on all six
  -- Excel Tables already show a floating header when scrolled past, so
  nothing is lost. `Setup` and `Summary` freeze only their top instructions
  row (row 1), well above where their own Tables start, so they were never
  affected and were left as-is. Verified via COM on all eight sheets:
  FreezePanes is False exactly on the six that had the bug, True on `Setup`/
  `Summary` at row 1 only, and every Table range/header is unaffected.
  **Not yet confirmed live with NVDA** -- next thing to ask the user to
  check, on all six affected sheets.

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

## Where we left off (2026-09-05)

Not just shipped-and-idle anymore — real user feedback has started coming in. The user
keeps a running notes file at `..\Spreadsheet Templit feedback.txt` (one level up, in
the parent working dir, NOT in the repo) and pastes tester feedback there between
sessions; read it fresh each time rather than trusting a stale summary here. The user
also has NVDA actually installed and running on this machine (not a sandbox) — real
screen-reader testing is possible in-session, but driving NVDA via simulated
keystrokes was deliberately avoided since it's their live, active session; the user
does the keypresses themselves and reports back.

Feedback batch (from a blind NVDA/JAWS tester, relayed 2026-09-04) and its resolution:
1. **FIXED, user-confirmed 2026-09-05.** Dropdown focus started at the bottom of the
   list after Alt+Down Arrow, forcing a scroll up through every item. Root cause: a
   real Excel behavior (not NVDA-specific) where an empty cell's dropdown matches the
   first blank entry in its source list, and every dropdown source here reserves
   trailing blank rows for growth. Fixed by trimming all 10 dropdown named ranges to
   their populated rows only — see `dynamic_list_ref()` and the matching entry under
   Excel/openpyxl traps above. The user tested this live with NVDA and confirmed it
   works.
2. **Still open, NVDA-side.** Choosing a dropdown option doesn't always return focus to
   the cell and can produce a keyboard trap (only escape is tab out of and back into
   Excel). This is a known upstream NVDA/Excel bug (nvaccess/nvda#13850, #13426), not
   fixable from the workbook. Documented on the Setup tab and in the README with the
   "toggle UI Automation in NVDA's Advanced settings" workaround.
3. Tester is on on-device Excel, Windows 10, exact build unknown. No workbook change
   needed — offered to add a "how to check your Excel version" pointer if wanted;
   user hasn't asked for it yet.
4. **Attempted two approaches; second one is current but not yet screen-reader-tested.**
   First tried pre-baking a hand-set NVDA+Shift+C/R header designation as a defined
   name (see the failed attempt logged under Excel/openpyxl traps above) — the user
   tested it live and it did NOT work after reopening the file. Replaced with making
   the Setup tab's Dropdown Lists grid a real Excel Table (`tbl_DropdownLists`), the
   same mechanism the log tabs already use successfully. Verified structurally via
   COM (proper rectangular Table, item-1 fix and Summary counts undisturbed) but
   **not yet confirmed live with NVDA** — that's the next thing to ask the user to
   check.
5. **NEW, fixed 2026-09-04, not yet NVDA-confirmed.** On the log tabs, NVDA announced
   "out of table" when arrowing from the header row (row 1) down into the first data
   row. Root cause: freeze panes split a Table's own header row from its data body
   into two Excel panes, and NVDA's table-region tracking is scoped per pane --
   crossing that split looked like leaving the table even though it's the same
   `ListObject`. Checked every sheet for the same exact pattern (freeze row equal to
   a Table's own first data row), not just the reported tab: found it on all four log
   tabs (`freeze_panes = "A2"`) and also on `Game Sections` and `Severity Scale`
   (each froze at their own Table's first data row) -- fixed by removing
   `freeze_panes` on all six; Excel Tables already show a floating header when
   scrolled past, so nothing is lost. `Setup` and `Summary` freeze only their top
   instructions row, well above where their own Tables start, so they were never
   affected and were left as-is. See the matching entry under Excel/openpyxl traps
   above. Verified structurally via COM on all eight sheets (FreezePanes False on
   the six fixed sheets, True at row 1 only on `Setup`/`Summary`, Table ranges/
   headers unaffected everywhere) but **not yet confirmed live with NVDA** -- ask
   the user to arrow from the header row into the data row on each of the six fixed
   sheets and confirm the announcement is gone.

The user mentioned they'd already copied the template into a separate file for a
specific game before this session's fix landed — that copy still has the old
trailing-blank named ranges and won't have picked up any of this. It's outside the repo
so rebuild_template.py can't touch it; flag to the user if it comes up again.

## Prior status (2026-08-23)

README fully written, uploaded via WSL/git, repo URL filled into the License section's
attribution string. LibreOffice/Google Sheets compatibility still unverified — don't
claim compatibility without testing it.

He uses this file as the between-session channel: read it first, treat anything he's
added as current instructions, and keep it accurate as things change.

## Open items for the user

The original launch checklist is done: README written, repo created and pushed,
`rebuild_template.py` shipped in the public repo, contribution policy decided (PRs
against the script only, not the `.xlsx`, since binary files don't diff/merge in git;
issues for everything else). LibreOffice/Google Sheets compatibility remains untested
and unclaimed, which is a deliberate honesty choice, not a gap to fill before shipping.

New as of 2026-09-04: see "Where we left off" above for the open tester-feedback items
(dropdown focus-at-bottom, possible keyboard trap, NVDA+Shift+C/R documentation tip).

## Working preferences

- The user wants improvements to their existing design, not replacements. Ask before
  restructuring something they built.
- They catch and care about correctness issues; verify claims in real Excel before
  stating them, and say plainly when something failed.
- Don't add their email address to any published file. Name-only attribution is intentional.
