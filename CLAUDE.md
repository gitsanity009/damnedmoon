# CLAUDE.md

## Project Overview

Blood Moon Protocol is an Excel-based choose-your-own-adventure RPG about lycanthropy, set in Detroit. The entire game lives in a single `.xlsx` file (`BloodMoonProtocol_RPG.xlsx`) with 18 sheets representing branching narrative scenes.

## Architecture

- **Format**: Single `.xlsx` spreadsheet file — no code, no macros, no dependencies
- **Navigation**: Hyperlinked buttons (styled cells) connect sheets to form a branching story graph
- **State tracking**: HP and Humanity stats are displayed per-scene but updated manually by the player
- **Layout convention**: Each sheet follows a consistent structure:
  - Rows 2-6: Header (title, scene name, HP/Humanity bars, moon phase)
  - Row 7: Divider
  - Rows 8-20: Narrative text
  - Row 22: Choice prompt
  - Rows 23-24: Option A (column A) and Option B (column H)
  - Row 27: Navigation buttons ("GO" hyperlinks)
  - Row 29: Optional warnings for dangerous choices

## Sheet Names and Flow

Sheets are named by narrative node: `TITLE`, `AWAKENING`, `SEARCH_AREA`, `FLASHBACK`, `CALL_MAYA`, `TREELINE`, `GO_APARTMENT`, `TELL_TRUTH`, `MAYA_KNOWS`, `LISTEN_ALPHA`, `ATTACK_ALPHA_EARLY`, `SUPPRESSANT`, `REJECT_HOLT`, `ACCEPT_HOLT_INFILTRATE`, `FINAL_HUNT`, `ENDING_VICTORY`, `ENDING_PYRRHIC`, `ENDING_BEAST`.

## Key Design Rules

- Every scene offers exactly 2 choices (A/B), except transition scenes and endings which have a single "continue" button
- Humanity score determines ending quality — high Humanity paths preserve the protagonist's agency
- The ATTACK_ALPHA_EARLY and ACCEPT_HOLT paths include explicit warnings when choices lead to bad endings
- Silver items are referenced narratively but tracked manually by the player
- Visual elements (progress bars, moon phase emoji, dividers) are rendered with Unicode characters in cells

## Working with the Excel File

- The `.xlsx` is a binary format — diffs are not meaningful in git
- To inspect content programmatically, use Python with `openpyxl`: `openpyxl.load_workbook('BloodMoonProtocol_RPG.xlsx', data_only=True)`
- Alternatively, `.xlsx` is a ZIP archive of XML files — sheets live under `xl/worksheets/sheet{N}.xml`
- When adding new scenes: create a new sheet, follow the existing layout convention, and add hyperlinks from the preceding choice sheet
