# Game Accessibility Testing Templit

A screen-reader-friendly Excel workbook for logging accessibility bugs while building or testing game accessibility mods. It's built for testers using NVDA, JAWS, or other screen readers, and free for anyone to copy and use on their own project.

## Overview

This is a manual accessibility issue/audit log, built for people creating accessibility mods, whether solo, with the help of AI, or as part of a team. It's flexible enough to send feedback to mod developers or game developers, or to use on your own games to collect feedback from others.

It's for anyone who wants a single, organized log instead of scattered notes. It gives you, your team, and any AI assistant you're using a clear record of: what issues exist, how players are affected, what environment each issue happens in (keyboard, controller, screen reader, visual, etc.), how severe it is, how to reproduce it, what the actual vs. expected behavior is, and its current status (open, resolved, and so on).

## What's Included

| Tab | What it's for |
| --- | --- |
| **Setup** | Game/mod info, plus every dropdown list the rest of the workbook uses (Status, Platform, Input Method, Screen Reader Used, and the five Error Type lists). |
| **Game Sections** | Every part of your game or mod being tested, one per row. Feeds the Game Section dropdown on all five log tabs. |
| **Severity Scale** | Ships with a basic 1 through 5 scale (Cosmetic through Blocking) that's fully unlocked. Rename it, extend it, or replace it outright. |
| **Visual** | Log visual accessibility issues: color, contrast, text scaling, missing visual feedback, and so on. |
| **Screen Reader** | Log issues found while testing with a screen reader: missing labels, bad reading order, silent interactions, and the like. |
| **Keyboard** | Log keyboard-only accessibility issues: focus, tab order, remapping, shortcut conflicts. |
| **Controller** | Log controller-specific issues: button mapping, remapping, prompt/glyph mismatches, haptics. |
| **Gameplay** | Log flow and pacing issues: confusing task order, unclear objectives, softlocks, missing save/pause points, difficulty spikes, tutorial gaps, and regressions where something that used to work no longer does. |
| **Summary** | Auto-updating totals for everything you log, by severity, by status, and by tab. |

## Accessibility Design

This is the part I care about most, so I want to actually explain the choices instead of just listing them.

- **No merged cells, no text boxes, no split headers.** Merged cells break screen-reader cell navigation. You land on a cell and it doesn't announce what you'd expect. Every sheet here is a plain grid on purpose.
- **Every log tab is a real Excel Table**, not just formatted cells. That's what makes NVDA/JAWS announce the column header every time you move down a row, so you always know what field you're in without having to remember it.
- **Status and severity are always plain text**, never color-only. A screen reader can't see a colored cell, so the words have to carry the meaning on their own: things like "5 - Blocking" instead of just a number, and "Open" instead of a colored dot.
- **Column headers are kept short.** A screen reader re-reads the header every time you land on a cell in that column, so a short header like "Game Section" beats a long one like "Which specific part of the game are you testing?" The original wording from the source questions is still there. It's just moved into a comment on the header, so you don't lose the context, you just don't hear it read out every single row.
- **Unused rows are genuinely blank.** The Issue ID column only fills in once you've picked a Game Section for that row, so an empty log doesn't read as 2,000 issues to a screen reader scanning down the column.

I tested this directly with NVDA and Narrator on Windows. The dropdowns, Tables, and formulas here all lean on standard Excel accessibility features, so the experience should be largely the same on other screen readers, but I haven't personally tried it with JAWS or VoiceOver. If you have and want to share how it went, see Contributing below.

## Getting Started

1. Download `Game Accessibility Testing templit.xlsx` from this repo.
2. Open it in Excel and fill in your game/mod details on the **Setup** tab.
3. List your game's components (levels, menus, systems, whatever fits your project) on the **Game Sections** tab, one per row.
4. Check the **Severity Scale** tab. It ships with a basic 1 through 5 scale, but it's fully unlocked, so rename or extend the levels any time.
5. Start logging issues on whichever tab matches the issue type: **Visual**, **Screen Reader**, **Keyboard**, **Controller**, or **Gameplay**.
6. Check the **Summary** tab any time for a running count of issues by tab and severity.

Tested in Microsoft Excel 365. Not yet tested in LibreOffice Calc or Google Sheets. Dropdown validation and Table behavior may not carry over exactly. If you try it there, I'd genuinely like to hear how it goes (see Contributing below).

If you're using a screen reader, see Using This With a Screen Reader below for the exact keystrokes and a walkthrough of logging your first issue.

## Checking for Updates

Once you've been using this for a while, your copy has your own logged issues in it, so you shouldn't just grab whatever's newest off GitHub and overwrite your file. `check_for_updates.py` handles this safely: it checks GitHub for a newer release, and if one exists, downloads it as a **separate, new file** next to yours. Your existing file is never touched, edited, or overwritten.

1. Make sure you have Python 3 installed (the same requirement as `rebuild_template.py` - this project's other script).
2. Download `check_for_updates.py` into the same folder as your workbook.
3. Run it:
   ```
   python check_for_updates.py
   ```
4. If a newer version exists, it downloads to a file like `Game Accessibility Testing templit (v1.1.0).xlsx` in that same folder. Open both files and copy your logged issues, Game/Mod Information, Game Sections, and Severity Scale entries over into the new one, then delete or archive the old one once you've confirmed everything moved over.

If your workbook lives somewhere else, point the script at it: `python check_for_updates.py --template-path "C:\path\to\your\file.xlsx"`.

No pip installs needed - it only uses Python's standard library, and works the same way on Windows, macOS, and Linux.

## Using This With a Screen Reader

If you're picking this up for the first time with a screen reader, here's how it's laid out and how to move around it.

### Layout at a glance

The workbook has 9 tabs across the bottom: Setup, Game Sections, Severity Scale, Visual, Screen Reader, Keyboard, Controller, Gameplay, and Summary. Move between them with **Ctrl+Page Down** and **Ctrl+Page Up**.

Every tab is a plain grid: no merged cells, no floating boxes, so arrow keys move exactly one cell at a time and always land where you expect. Every data grid in this workbook - the five log tabs (Visual, Screen Reader, Keyboard, Controller, Gameplay), the Dropdown Lists and Game Info areas on Setup, Game Sections, Severity Scale, and Summary - is built as a real Excel Table, so your screen reader announces the column header every time you move to a new cell in that column. You don't need to count columns or remember what you're filling in as you go.

On the five log tabs (Visual, Screen Reader, Keyboard, Controller, Gameplay), row 1 holds the column headers and your data starts on row 2. Game Sections and Severity Scale work the same way, just with a short instructions block above their own header row. Setup and Summary lead with a row or two of instructions before their actual headers, since there's more to explain up front on those - and those two are the only tabs that freeze that top instructions row in place while you scroll.

### Keyboard shortcuts you'll actually use here

These are standard Excel shortcuts, not anything specific to this template, but they're the ones you'll reach for constantly while working in this workbook.

| Shortcut | What it does |
| --- | --- |
| Ctrl+Page Down / Ctrl+Page Up | Move to the next / previous tab |
| Ctrl+Home | Jump to the top of the sheet. On Setup and Summary this lands you just below their frozen top instructions row. Everywhere else (the five log tabs, Game Sections, Severity Scale) nothing is frozen, so it lands on literal row 1, the column header row - press Down Arrow once to reach the first data row |
| Arrow keys | Move one cell at a time |
| Tab / Shift+Tab | Move one cell right / left |
| Enter / Shift+Enter | Move one cell down / up |
| Alt+Down Arrow | Open the dropdown on the current cell |
| Up/Down Arrow (dropdown open) | Move through the dropdown's options |
| Enter (dropdown open) | Select the highlighted option |
| Esc | Close a dropdown without choosing anything |
| Ctrl+; | Insert today's date (handy for the Date Reported column) |

Don't press F2 to open a dropdown. It switches the cell into text-edit mode, where the arrow keys stop moving through the list. Alt+Down Arrow is the one that works. This is standard Excel behavior, not something specific to this template.

**A known NVDA/Excel bug can leave focus stuck after picking a dropdown option.** After you press Enter to choose an item, focus can occasionally get stuck, with tabbing out of the Excel window and back in as the only way to recover. This is a longstanding issue in how NVDA and Excel talk to each other ([nvaccess/nvda#13850](https://github.com/nvaccess/nvda/issues/13850), [#13426](https://github.com/nvaccess/nvda/issues/13426)), not something in this workbook. If you hit it, try toggling "Use UI Automation to access Microsoft Excel spreadsheet controls" in NVDA's Preferences > Settings > Advanced menu - whichever state you flip it to (on if it's off, off if it's on) clears it for most people, depending on your Excel/NVDA version.

The Dropdown Lists grid on the Setup tab (Status, Platform, and so on) is a real Excel Table too, same as the five log tabs, so it announces its column headers the same automatic way as you move down it - no extra setup needed there either.

A heads-up on two shortcuts that don't behave the way you'd expect here: **Ctrl+End** and **Ctrl+Arrow key** normally jump to the last used cell or the edge of your data. But since each log tab is a pre-built 2,000-row table, both of them jump to the edge of the whole table instead of to your last logged row. The reliable way to find your last entry is to move down column C (Game Section) from the top until you hit a blank cell.

### Logging your first issue

1. From wherever you are, Ctrl+Page Down until you reach the tab that matches what you're testing: Visual, Screen Reader, Keyboard, Controller, or Gameplay.
2. Ctrl+Home to jump to row 1, the column header row, then Down Arrow once to reach row 2, the first data row. If you've already logged issues on this tab, keep arrowing down until you reach a blank row.
3. Move to the Game Section column and press Alt+Down Arrow to pick from the list.
4. Tab across the row, filling in Error Type, Actual Behavior, Expected Behavior, and Steps to Reproduce as you go. The dropdown columns (Error Type, Severity, Platform, Input Method, Screen Reader Used, Status) all open the same way: Alt+Down Arrow, arrow keys, Enter.
5. As soon as you pick a Game Section, the Issue ID cell at the start of the row fills itself in (V-001 for your first entry on Visual, V-002 for your second, and so on), so you'll hear it change from blank to a real ID.
6. Move down to the next row for your next issue.

Unused rows stay genuinely blank, so scanning past your last logged issue reads as empty, not as thousands of phantom entries. And status and severity are always read out as plain words, like "Open" or "5 - Blocking", never just a color, so nothing here depends on being able to see the sheet.

## Customizing the Template

You're more than welcome to download this and add or remove whatever you want. It's built to be flexible, easy to use, and fully open-source for the community. (Full disclosure: I'm not an Excel expert. I used Claude to help get everything working the way I wanted it to, and I reviewed and edited everything along the way.)

The main things designed to be changed:

- **Game Sections**: swap in your own game's components.
- **Severity Scale**: the whole tab ships fully unlocked, so rename the levels, add more, or restructure the scale entirely, any time (see Getting Started above).
- **Dropdown lists on the Setup tab** (Status, Platform, Input Method, Screen Reader Used, and the Error Type lists): the list values themselves are left unlocked even though the Setup tab stays protected, so you can edit or add entries directly in the blank rows already reserved below each list, no unprotecting required. This is also where you'd swap in terminology from a framework like WCAG or the Game Accessibility Guidelines, most naturally in the Error Type lists, since those categorize the kind of issue rather than how severe it is. Adding more entries than that, or extending a list past what's reserved, does require unprotecting the sheet first (see Sheet protection below) and updating that list's named range in Excel's Name Manager.
- **Custom log tabs**: copy an existing tab, rename it, and add the new tab name to the Summary tab so it gets counted.

**Reusing this for a new project:** make a fresh copy of the spreadsheet for each mod or game you're working on. That gives each project its own clearly organized issue log, instead of whatever tool you'd otherwise be using to track things.

### Sheet protection

Every tab ships protected except Severity Scale, which is left fully unlocked on purpose. On every protected tab, headers and formulas are locked so you can't accidentally overwrite them, but every cell you're meant to type into, including the dropdown list values on the Setup tab, is left unlocked. There's no password on any of it.

If you ever need to restructure a tab, for example adding a column or inserting extra rows, go to **Review > Unprotect Sheet** first. Nothing is locked down to stop you, it's just locked by default so a stray keystroke doesn't wipe out a header or a formula.

The Severity Scale tab is the one exception: the whole tab ships fully unlocked, because it's meant to be rewritten with your own severity wording. The dropdown lists on the Setup tab work a little differently: the tab itself stays protected, but each list's values are individually left unlocked so you can edit them without unprotecting anything (see Customizing the Template above).

## Known Limitations

- **LibreOffice and Google Sheets are untested.** Everything here has only been verified in Microsoft Excel. Dropdown validation and Table behavior can work differently in other spreadsheet apps, so I can't say yet whether this holds up there. If you try it, I'd like to know (see Contributing below).
- **The pre-built rows have a limit.** Each log tab ships with 2,000 rows and each dropdown list has some room to grow, but if you ever fill all of them, you'll need to unprotect the sheet first before inserting more (see Sheet protection above).
- **Custom tab names have to match exactly.** If you add a custom log tab and list it on the Summary tab, the name you type there has to match the actual tab name exactly, apostrophes and all, or it'll show "Check tab name" instead of a count.
- **A dropdown could open scrolled to the bottom instead of the top - fixed.** Every dropdown source list here reserves blank rows below its seed values so you can add your own entries later. Those trailing blanks used to make Excel's dropdown open already scrolled to the bottom on an empty cell (Excel matches a blank cell to the first blank entry in its source list, which sat at the end) - a real Excel behavior, not NVDA-specific. The dropdown source ranges now trim off their own unused rows, so this no longer happens. This does mean an entry deleted from the *middle* of a dropdown list (rather than typed over) can cut off whatever comes after it in the dropdown - keep entries as a single unbroken run starting at the first data row, as the on-sheet instructions already say.
- **Some NVDA users still hit a post-selection focus issue.** A tester reported NVDA occasionally getting stuck after choosing a dropdown option. This traces back to a known bug in NVDA's Excel support, not this workbook - see the NVDA/Excel bug note under Using This With a Screen Reader above for the workaround.

## Contributing

You're more than welcome to fork this, adapt it, and use it however works for your project.

For contributions back to the main project: open a pull request against `rebuild_template.py`, not the `.xlsx` file itself. The spreadsheet is a generated build artifact, and binary files like `.xlsx` don't diff or merge cleanly in git, so I can't meaningfully review a change made directly inside the file. Change the script, and once it's merged I'll regenerate and commit the workbook.

If you'd rather not touch Python, or just want to suggest something, open an issue instead. I'd also like to hear about screen reader testing results on setups other than mine (see Known Limitations above); open an issue if you try it somewhere new.

Want to talk something through first? Reach me on Discord (`nion_light0972`).

## Credits

Built by Christopher Tavarez, with Claude's help designing the formulas, dropdowns, and structure, and testing all of it in real Excel.

[github.com/ChrisTavar2022](https://github.com/ChrisTavar2022)

Additional thanks to [Ross Minor](https://github.com/RossMinor) for testing and suggestions.

Thanks to anyone who sends in fixes, testing results, or feedback. I'll add names here as that happens.

## License

This project is licensed under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). See the [LICENSE](LICENSE) file for the full terms.

That means you're free to use, adapt, and redistribute this template, including commercially, as long as you give appropriate credit.

Suggested attribution:

> "Game Accessibility Testing Templit" by Christopher Tavarez, licensed under CC BY 4.0. Source: https://github.com/ChrisTavar2022/Game-Accessibility-Testing-Templit
