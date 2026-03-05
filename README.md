# Blood Moon Protocol

**A Werewolf Choose-Your-Own-Adventure RPG — built entirely in a single Excel spreadsheet.**

## Overview

Blood Moon Protocol is an interactive fiction RPG where you play as **Marcus Daye** — an ex-Army medic and night-shift security guard in Detroit who wakes up in Elgin Park three days after being attacked by something inhuman. The entire game runs inside an `.xlsx` file using hyperlinked buttons to navigate between scenes.

## How to Play

1. Open `BloodMoonProtocol_RPG.xlsx` in Microsoft Excel, Google Sheets, or any compatible spreadsheet application
2. Start on the **TITLE** sheet and read the character overview
3. Click the **GO** buttons to make choices and advance through the story
4. Track your **HP** and **Humanity** stats as the narrative instructs — update the character sheet manually as you play
5. Your choices determine which of **3 endings** you reach

## Game Mechanics

| Stat | Description |
|------|-------------|
| **HP** | Your health. Starts at 100. Combat and reckless choices reduce it. |
| **Humanity** | Starts at 100. Embracing the beast lowers it. Dropping to 0 means losing yourself permanently. |
| **Moon Phase** | Tracks the current lunar cycle. The full moon intensifies the wolf's pull. |
| **Silver Items** | Collectible items that matter in the final confrontation. |
| **Special Trait** | *Iron Will* — military training grants +10 resistance against the first transformation urge. |

## Story Structure

The game is organized across **18 Excel sheets**, each representing a scene or story beat:

```
TITLE ─── AWAKENING
              ├── SEARCH_AREA
              │     ├── CALL_MAYA ──┬── TELL_TRUTH ──┬── FINAL_HUNT ──┬── ENDING_VICTORY (★★★★)
              │     │               │                │                ├── ENDING_PYRRHIC (★★★)
              │     │               │                │                └── ENDING_BEAST (★)
              │     │               └── MAYA_KNOWS ──┤
              │     └── TREELINE                     │
              │           ├── LISTEN_ALPHA ──┬── REJECT_HOLT ────────┤
              │           │                  └── ACCEPT_HOLT ────────┤
              │           └── ATTACK_ALPHA_EARLY ── SUPPRESSANT ─────┘
              └── FLASHBACK
                    ├── GO_APARTMENT ──► (merges into CALL_MAYA / TREELINE paths)
                    └── (merges into CALL_MAYA path)
```

### The Three Endings

- **The Sentinel** (★★★★) — You defeat Holt and keep your humanity. Marcus becomes Detroit's protector.
- **The Price** (★★★) — You win, but the wolf is closer than ever. Victory at a cost.
- **Gone** (★) — The wolf takes over completely. Marcus Daye's choices end here.

## Characters

- **Marcus Daye** — The protagonist. Ex-Army medic, night-shift security guard, newly turned werewolf.
- **Maya Okafor** — Marcus's best friend from Howard University. Forensic biologist who has been independently researching anomalous cases for months.
- **Raymond Holt** — The Alpha. A former CDC researcher who isolated a retroviral lycanthropy sequence and deliberately infected Marcus.

## Requirements

Any spreadsheet application that supports `.xlsx` files with hyperlinks:
- Microsoft Excel (recommended)
- Google Sheets
- LibreOffice Calc

## License

All rights reserved.
