# CferNpcMaker — Feature Backlog

Work through this list one item per cycle using the instructions in `CYCLE.md`.
Check off each item when merged.

## User Experience
- [x] 1. **Retry button on failed history entries** — failed entries show a Retry button that re-submits the same form data without re-typing
- [x] 2. **Generation progress steps** — replace the spinner with labelled steps ("Sending request… Building spell mapping… Creating actor…")
- [x] 3. **History entry deletion** — delete individual history entries or clear all with a "Clear History" button
- [x] 4. **History search/filter** — search box above the history panel that filters entries by name
- [x] 5. **Duplicate from history** — button on each history entry to pre-fill the form with its data
- [x] 6. **Undo last generation** — one-click delete of the most recently created actor/item from inside the builder
- [x] 7. **Name suggestions** — "Suggest a name" button that generates a random thematic name based on the description without a full generation
- [x] 8. **Description templates** — dropdown of starter templates ("Town guard", "Merchant", "Cult leader") that pre-fill the description field
- [x] 9. **Compact history panel** — toggle to collapse the history panel to a narrow strip to give the form more space
- [x] 10. **Keyboard shortcuts** — `Ctrl+Enter` to generate, `Escape` to clear the form

## Reliability & Error Handling
- [x] 11. **Offline detection** — detect no network and show a clear banner instead of a confusing auth error
- [x] 12. **Session pre-validation** — on app open, silently test the stored key and proactively prompt re-auth if invalid before the user clicks Generate
- [x] 13. **Partial failure recovery** — if actor creation fails after all retries, offer to download the raw JSON so the user doesn't lose generated data
- [x] 14. **Rate limit countdown** — parse the reset timestamp from 429 responses and show "Limit resets in X days"
- [x] 15. **Concurrent generation guard** — disable the Generate button (not just via CSS class) while a generation is in-flight

## Generation Quality
- [x] 16. **Post-generation quick-edit dialog** — after creation, show a lightweight overlay (name, level, key stats) so users can fix small mistakes before the full sheet opens
- [x] 17. **Variation slider** — a "Creativity" slider sent as a parameter to the endpoint to control how closely the AI follows the description
- [x] 18. **Bulk generation** — generate multiple NPCs in one request and create them all as separate history entries
- [x] 19. **Level Up for D&D 5e** — add the same CR-change sheet button that PF2e has, for 5e creatures
- [x] 20. **Image generation for item modules** — optional portrait/icon generation for PF2e and 5e items

## Developer Experience
- [x] 21. **Adapter validation on construction** — check all required getters return non-empty values at init time and throw descriptive errors immediately
- [x] 22. **`Storage.clear()`** — method that wipes all stored keys for a module (useful for testing and troubleshooting)
- [x] 23. **Build-time type checking** — add `jsconfig.json` with `checkJs: true` so VS Code surfaces type errors in the shared core
- [x] 24. **Shared test harness** — minimal Vitest test file covering `isNewerVersion`, `escapeHtml`, `devUrl`, and `Storage`

## Performance
- [x] 25. **Spell mapping cache** — cache the PF2e spell mapping in `sessionStorage` keyed by a hash of installed packs, skipping the 5–10 s build on repeat generations
- [x] 26. **Template preloading at build time** — inline the Handlebars template into the entry bundle so `loadTemplates` is a no-op at runtime

## Polish & Accessibility
- [x] 27. **Accessible focus management** — move focus to the new history entry after generation so keyboard users don't lose their place
- [x] 28. **CSS custom properties for colours** — replace inline hex colours in dialogs with CSS variables that respect Foundry's theme
- [x] 29. **Localisation scaffold** — wrap all user-facing strings in `game.i18n.localize()` and add a starter `en.json` lang file
- [x] 30. **Settings page audit** — surface "Reset welcome message" and "Clear session" as buttons in the Foundry module settings UI

---

# Pf2e Settlements + Calendar Backlog (v2.5.5 onwards)

Long-running roadmap for **`Pf2eNationsAndCitiesMaker`** and
**`Pf2eCalendarTimeline`** until they form a fully-fledged nation-management
system (settlements → economy → military → diplomacy → war → calendar). Each
item is sized to be **one cycle / one PR**. Pick the lowest-numbered unchecked
item that has all its dependencies satisfied.

> Run one a day. Tick the box and add the PR number when each lands.

## A. AI backend & generation quality (31–38)
- [ ] 31. **`/webhook/city-builder` endpoint live** — design the n8n workflow, the system prompt, and the JSON-schema-constrained output that maps 1:1 to the `sanitizer.sanitizeSettlement` shape. Cap stores at 12/city, 6/town, 2/village.
- [ ] 32. **`/webhook/city-builder-dev` mirror** — same endpoint with relaxed rate limits, routed via the existing `devMode` setting.
- [ ] 33. **Prompt-tuning pass #1** — run 50 generations across biomes & sizes, log failures (names too modern, inventory anachronisms, missing fields). Add few-shot examples until 8/10 generations are usable as-is.
- [ ] 34. **Variation / creativity slider** — float 0–1 sent to the endpoint; low values follow the description tightly, high values riff freely. Default 0.5.
- [ ] 35. **Bulk store regeneration** — header button on the City sheet that re-rolls just `settlement.stores` (preserving everything else).
- [ ] 36. **Per-store re-roll** — card-level "Re-roll this store" action that hits the AI for a single store of the same `type`.
- [ ] 37. **`/webhook/nation-builder`** — top-level endpoint that produces a Nation plus 3–6 child city stubs in one call, with auto-linked `childCityIds`.
- [ ] 38. **Cost-aware generation** — show the user the estimated Patreon-use cost before generating (1 use per ~3 stores).

## B. Settlement sheet UX & polish (39–50)
- [ ] 39. **Inline NPC chip** — when `staff.actorId` (or `owner.actorId`) is set, render the actor's portrait + level + a Level-Up button inline instead of a bare icon.
- [ ] 40. **Drag-drop actor → staff row** — drop a PF2e NPC onto a staff name field to link `actorId`.
- [ ] 41. **Drag-drop item → inventory** — drop a PF2e item onto the inventory table to add a row with price/stock pulled from `item.system.price`.
- [ ] 42. **Per-store wealth tier** — `priceTier: low|standard|high|luxury` select on each card; multiplies displayed prices and biases item-generator prefill level.
- [ ] 43. **Weekly schedule grid** — replace the single `open/close` pair with a 7-day grid using the calendar's `weekdays` so a shop can be closed on rest-day.
- [ ] 44. **Shift roster as enum** — `shift` becomes `morning|day|evening|night|graveyard`, rendered as colour-coded chips.
- [ ] 45. **District / quarter subdivision** — add `districts: [{ id, name, storeIds[], descriptor, leaderActorId }]`; Stores tab gains a district filter alongside the type filter.
- [ ] 46. **Sheet permissions per section** — `gmOnly` flags on Treasury / Income / Military panels; respect them when a player owns the journal.
- [ ] 47. **Partial re-render** — switch the sheet to ApplicationV2 PARTS-based rerender so editing one field no longer loses focus.
- [ ] 48. **Debounced auto-save** — replace `change`-on-blur with debounced auto-save while typing (no extra Save button).
- [ ] 49. **PDF / Markdown handout export** — header button that renders the settlement as a printable handout (inputs collapsed, statblock-style typography).
- [ ] 50. **Compact / expanded sheet toggle** — collapse store cards to one-line summaries (name, owner, balance) for quick scanning.

## C. Economy depth (51–62)
- [x] 51. **Income jitter knob** — world setting for the variance percentage used in `economy.applyDailyTick` (currently hard-coded 15%).
- [x] 52. **Production → treasury credit** — `applyDailyTick` also credits the treasury based on `production.length × population / 1000` gp per day.
- [ ] 53. **Guard wage upkeep** — each rank has a `dailyWage`; total drains treasury on every `dayAdvanced`.
- [ ] 54. **Unrest dynamics** — `unrest += 1` per tax fired, `-1` per festival event fired, clamped 0–100.
- [ ] 55. **Riots at high unrest** — when unrest > 80, store `dailyAvg` is halved and a chat card warns the GM.
- [ ] 56. **Bankruptcy & closure** — stores with `balance < 0` for 30+ days get `closed: true`, hidden from the active stores tab (with a "Show closed" toggle).
- [ ] 57. **Trade-goods catalogue** — bundle a small JSON table of typical PF2e trade goods + prices; production tags map to goods automatically.
- [ ] 58. **Trade routes between settlements** — link two journals with `tradeRoutes: [{ partnerId, goods[], gpPerWeek }]`; calendar tick moves gp + adds goods to both treasuries.
- [ ] 59. **Market days** — each store can opt into a `marketWeekday`; on that day its `dailyAvg` doubles.
- [ ] 60. **Black-market toggle** — `isBlackMarket: true` on a store: hidden from players unless they have permission, higher prices, attracts unrest.
- [ ] 61. **Treasury history graph** — last 30 days of gp totals as an inline sparkline on the Overview tab (no chart library, plain SVG).
- [ ] 62. **Inflation modifier** — settlement-wide `priceMultiplier` (1.0 default) applied to all displayed prices; can be raised by events like sieges.

## D. Population & society (63–69)
- [ ] 63. **Population growth model** — per-cycle `population += round(population × growthRate × (1 - unrest/100))`; default growthRate 0.001/day.
- [ ] 64. **Plague event** — calendar event of `kind: 'plague'` reduces population by a configurable %.
- [ ] 65. **Famine event** — drains treasury, raises unrest, halts growth until resolved.
- [ ] 66. **Religion tracking** — `religions: [{ name, followers, templeStoreId, influence }]`; surfaced on the Leadership tab.
- [ ] 67. **Crime / banditry level** — 0–10 setting that raises tax leakage and triggers chat-card incident reports on `dayAdvanced` rolls.
- [ ] 68. **Adventurer-hook generator** — sheet header button calls a small AI endpoint (or local table) to produce 3 plot hooks tied to the settlement's current state.
- [ ] 69. **Demographics breakdown** — race / ancestry / class mix as percentages on the Overview tab; editable so the GM can pin "elven port" vs "dwarven hold".

## E. Military & warfare (70–82)
- [ ] 70. **Army document** — separate JournalEntry kind (`flags.…settlement.kind === 'army'`) with its own sheet, holding unit counts + a stationed-at link.
- [ ] 71. **Unit roster** — `units: [{ type, count, level, equipment, morale }]` with PF2e-flavoured types (spearmen, archers, cavalry, mages, siege).
- [ ] 72. **Recruitment dialog** — drains `population` and `treasury.gp` to add units of a chosen type; respects available pool.
- [ ] 73. **Unit wages** — each unit type has a daily upkeep cost; the army's `stationedAt` settlement pays it on each `dayAdvanced`.
- [ ] 74. **Garrison vs field-army flag** — `mode: 'garrison'|'field'`; garrisons defend the settlement HP, field armies can move.
- [ ] 75. **Army movement** — assign a destination settlement; arrival time computed from calendar + distance setting, firing `armyArrived` on completion.
- [ ] 76. **Battle resolver** — modal dialog: pick attacker + defender army; rolls a deterministic mass-combat formula (numbers × level × morale × terrain) and writes casualties back.
- [ ] 77. **Siege mechanics** — attacking a settlement (not an army) drains settlement HP using `damageThreshold` + `hardness`; HP → 0 = occupied state.
- [ ] 78. **Generals (commander Actor)** — set an Actor as army commander; their level adds a flat bonus to mass-combat rolls.
- [ ] 79. **Supply chain** — armies drain food from their origin settlement per day; cut supply → morale drops until starvation.
- [ ] 80. **Surrender / occupation flag** — defeated settlements get `occupiedBy: <nationId>`; the occupier's nation gains gp drain rights.
- [ ] 81. **Mercenary companies** — buyable temporary armies tied to a contract length; auto-disband on contract end.
- [ ] 82. **War chat-cards** — every battle / siege resolution posts a styled chat card with casualty totals and a "Details" expander.

## F. Diplomacy & nations (83–92)
- [ ] 83. **Relations matrix** — Nation sheet gains a grid: every other Nation → `relation: 'ally|friendly|neutral|cold|hostile'` + numeric score (-100..100).
- [ ] 84. **Treaty objects** — `treaties: [{ partnerNationId, kind, signedOn, expiresOn, terms }]`; kinds = non-aggression / defensive / trade / vassalage.
- [ ] 85. **Treaty expiry events** — auto-schedule calendar events when treaties expire and fire a chat reminder.
- [ ] 86. **Vassal / suzerain links** — a Nation can list `vassalNationIds[]` or set `suzerainNationId`; aggregated stats roll up.
- [ ] 87. **Claims** — `claims: [{ targetSettlementId, kind: 'historical|dynastic|religious' }]` provide casus-belli filtering.
- [ ] 88. **Declare war flow** — modal dialog: pick target nation, pick claims, posts a chat card, sets relations to hostile, fires `Pf2eNationsAndCitiesMaker.warDeclared`.
- [ ] 89. **Peace negotiation dialog** — drag claims into "given up / kept / new" buckets; on accept, writes a treaty and resets relations.
- [ ] 90. **Faction influence per city** — `factions: [{ name, type: 'church|guild|nobles|criminal|crown', influence: 0..100 }]`; total must be 100.
- [ ] 91. **Succession** — monarchy / dynasty support: `heir: { actorId, name }`; on a calendar `rulerDied` event, swap the leader to the heir.
- [ ] 92. **Diplomatic gifts** — chat-button action that transfers gp between nations and bumps their relation score.

## G. Calendar features (93–101)
- [ ] 93. **Custom calendar editor UI** — modal that edits `calendarDef` (months, daysPerMonth, weekdays, monthNames).
- [ ] 94. **Golarion preset** — pre-built Inner Sea calendar (12 months, 7 weekdays, official month names).
- [ ] 95. **Gregorian preset** — Earth months / weekdays.
- [ ] 96. **Generic-fantasy preset** — neutral 10-month cycle, 5-day weeks.
- [ ] 97. **Per-day weather roll** — `weather: 'clear|overcast|rain|storm|snow'` derived from biome + month; fires `Pf2eCalendarTimeline.weatherChanged`.
- [ ] 98. **Season detection** — emit `seasonChanged` when `month` crosses spring/summer/autumn/winter boundaries; surfaced on the calendar header.
- [ ] 99. **Festival event template** — preset that boosts target settlements' morale and pulls gp from treasuries (cost of festival).
- [ ] 100. **Travel-narration helper** — "Advance with travel" button: enter days + party size, fires `inTransit` so other modules can drain rations.
- [ ] 101. **Click a day to inspect events** — popover listing every event scheduled for that day with quick edit / cancel buttons.

## H. Map & visuals (102–108)
- [ ] 102. **Settlement banner image picker** — file picker on the header; falls back to AI-generated banner when set.
- [ ] 103. **Procedural town SVG** — render a top-down layout: city wall, streets, shop dots sized by population; click a dot to scroll to that store card.
- [ ] 104. **World-map scene pins** — when a settlement has `sceneId`, the scene note uses the module icon and tooltip shows kind + population.
- [ ] 105. **Nation border overlay** — Nation sheet lets the GM draw a region on the world map; the region is rendered with the nation's colour.
- [ ] 106. **Scene-link button** — header button to set/clear `sceneId`; opens the linked scene when clicked.
- [ ] 107. **Settlement icon by kind** — village / town / city / nation get distinct icons on the journal directory.
- [ ] 108. **Embed thumbnail in chat cards** — generation-success / tax-fired chat cards include the banner thumbnail.

## I. Cross-module integration (109–115)
- [ ] 109. **Auto-tag generated NPCs** — NPCs created through a settlement's bridge get `actor.setFlag(Pf2eNationsAndCitiesMaker, 'homeSettlementId', journalId)`.
- [ ] 110. **Reverse-link from actor sheet** — button on tagged actors: "Open home settlement".
- [ ] 111. **Compendium of 10 prebuilt settlements** — ship a JSON compendium so users can drop a town in without AI credits.
- [ ] 112. **Compendium of 5 prebuilt nations** — same, but at nation tier.
- [ ] 113. **Simple Calendar adapter** — optional bridge: if `foundryvtt-simple-calendar` is active, drive our events off its date.
- [ ] 114. **Drop settlement onto scene to link** — drag a settlement journal onto a scene to set `sceneId` automatically.
- [ ] 115. **Adventurer guild integration** — a special store type that lets the GM post "jobs" (custom payloads) and roll which adventurers (NPCs) take them.

## J. Performance, DX & migrations (116–122)
- [x] 116. **Schema versioning** — stamp settlements with `_schemaVersion`; add `migrations.js` that upgrades old shapes forward on every read.
- [x] 117. **Vitest harness wired into CI** — covering `scheduler.computeNext`, `sanitizeSettlement` defaults, `economy.applyDailyTick` jitter bounds.
- [x] 118. **JSDoc typedefs** — `@typedef` blocks for `Settlement`, `Store`, `Staff`, `Rank`, `Army`, `Treaty`, `Event`.
- [x] 119. **`jsconfig.json` with `checkJs: true`** — surfaces type errors in VS Code without TypeScript.
- [x] 120. **Settings-page audit** — add "Reset welcome message", "Clear stored settlements", "Reset calendar to default" buttons.
- [x] 121. **Console logging tier** — `Pf2eNationsAndCitiesMaker.logLevel` setting (error/warn/info/debug) to silence noisy logs in production.
- [ ] 122. **Storybook-style sheet preview** — a hidden GM-only "Open sheet with fixture data" macro for QA.

## K. Accessibility & localisation (123–128)
- [ ] 123. **i18n pass (Settlements)** — wrap every user-facing string in `game.i18n.localize()`, ship `lang/en.json`.
- [ ] 124. **i18n pass (Calendar)** — same.
- [ ] 125. **High-contrast theme variant** — toggle in module settings; overrides the parchment palette with strong-contrast colours.
- [ ] 126. **Keyboard navigation** — Tab order through sheet tabs, Enter on store header toggles compact view, Esc closes the sheet.
- [ ] 127. **ARIA roles** — `role="tablist"` / `role="tabpanel"` on sheet tabs; `aria-busy` on generating buttons; visible focus rings.
- [ ] 128. **Screen-reader live region** — tax-fired / day-advanced notifications announced to assistive tech.

---

## How to pick the next task

1. Skim Phase A first — AI quality is the foundation; if it's lousy, nothing else lands well.
2. Within a section, work top-down (dependencies generally point upward).
3. Sections **E (military)** and **F (diplomacy)** depend on Section D (population) for casualty / influence math — don't start them until population growth (#63) is in.
4. Phase J (#116 schema versioning) should land **before** any backlog item that adds a new field to the settlement schema — to avoid breaking existing users.

## Definition of done (per item)

- Implementation + minimal styling consistent with existing palette.
- `./build.sh` for all seven modules still passes.
- Any new field added to the settlement schema has a safe default in `sanitizer.js`.
- PR description includes a one-line manual-test recipe.
- The item's checkbox is ticked with the PR number on merge: `- [x] 42. … (#NN)`.

