# Blood Moon Protocol — Implementation TODO

## Phase 1: Core Engine
- [x] **Scene Database System** — `SceneDB` sheet + `modData` VBA module with scene lookup (SceneID, story text, choices, next scene links, conditions, on-enter effects)
- [ ] **Game State Manager** — `modState` VBA module with stat/flag get/set, save/load to hidden sheet, auto-logging of actions for replay/debug
- [ ] **Choice Resolver** — `modEngine` VBA module with `LoadScene`, `ApplyEffects`, `ResolveChoice`; button-based choices, condition validation, effect application

## Phase 2: UI/UX
- [ ] **Game Screen UI** — Dedicated sheet with story window (auto-scroll/paging), 3–6 dynamic choice buttons, Continue/Back buttons
- [ ] **modUI Rendering** — `modUI` VBA module for `RenderScene`, `RenderChoices`, HUD bar updates, status effect icons

## Phase 3: RPG Systems
- [ ] **Stat System** — Base stats + derived stats (e.g. Control = Humanity − Rage), clamping, death/failure states, difficulty scaling
- [ ] **Inventory System** — Items sheet with IDs/names/types/stack rules, equip slots (weapon, charm, coat), item effects (passive buffs, consumables)
- [ ] **Flags & Branching** — Boolean flags (MetHunter, StoleSilverKnife, BittenConfirmed), numeric reputation (VillageTrust, HunterSuspicion)
- [ ] **Skill Checks / Dice** — D20/percentile/custom system, advantage/disadvantage mechanics, adaptive success/failure text
- [ ] **Combat Encounters** — Turn-based quick combat loop (player choice → enemy response), randomized damage ranges and crits, escape/flee logic
- [ ] **Moon/Time System** — Time advances per scene or per action, moon phase changes over time (affects transformation chance), forced events at thresholds (full moon, near-dawn)

## Phase 4: Werewolf-Specific Features
- [ ] **Transformation System** — Multi-stage (itch → crack → blackout → aftermath), control checks based on Rage/Humanity/Hunger + moon phase, blackout scenes with semi-random outcomes
- [ ] **Scent & Tracking** — Scent level increases with blood/fear/running, hunters track you better at high scent
- [ ] **Morality Tension** — Humanity influences story tone and endings, feeding grants power but costs Humanity / raises Suspicion
- [ ] **Hunter AI / Suspicion Meter** — Village suspicion rises from choices, hunters appear more often, choices become restricted

## Phase 5: Persistence & Safety
- [ ] **Save Slots & Persistence** — Multiple save slots with timestamps, auto-save on scene change, undo/rewind via stack-based snapshots, reset game function, defensive error handling with friendly messages

## Phase 6: Content Authoring Tools
- [ ] **Scene Editor Helpers** — "Add Scene" UserForm that writes rows into SceneDB, validation checks (missing SceneIDs, dead links, orphan scenes), branch map generator (SceneID → NextIDs graph with unreachable/loop detection)
- [ ] **Localization-ready Text** — Separate text strings table so you can rewrite narrative without breaking logic

## Phase 7: Polish
- [ ] **Achievements & Endings Tracker** — Ending gallery (Exile, Contained, Cured?) with unlocks
- [ ] **Random Encounters** — Weighted event tables (forest events, village events, hunter events)
- [ ] **Secrets & Dynamic Text** — Hidden flags + special scenes unlocked by unusual combinations, pronoun/name insertion, stat-aware narration (e.g. if Rage > 80, narration becomes more predatory)
- [ ] **Modal Popups & Sound Hooks** — Dice roll results, alerts (Transformation imminent), item pickups, optional sound hooks (clicks, howls, heartbeat, tension stingers)
- [ ] **Styling System** — Dark theme toggle, readable typography, consistent layout, optional typewriter text effect for story output
- [ ] **Anti-corruption & Error Handling** — Reset game function that wipes state clean, defensive error handling with friendly messages

---

## VBA Module Architecture
```
modEngine  — LoadScene, ApplyEffects, ResolveChoice
modState   — Get/Set stats, flags, inventory
modUI      — RenderScene, RenderChoices, HUD updates
modData    — Story lookup, validation, random tables  ✅
frmSaveLoad — Optional save/load UI
```
