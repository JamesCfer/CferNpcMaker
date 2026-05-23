# CferNpcMaker — Feature Backlog

Work through this list one item per cycle using the instructions in `CYCLE.md`.
Check off each item when merged.

## User Experience
- [x] 1. **Retry button on failed history entries** — failed entries show a Retry button that re-submits the same form data without re-typing
- [x] 2. **Generation progress steps** — replace the spinner with labelled steps ("Sending request… Building spell mapping… Creating actor…")
- [ ] 3. **History entry deletion** — delete individual history entries or clear all with a "Clear History" button
- [ ] 4. **History search/filter** — search box above the history panel that filters entries by name
- [ ] 5. **Duplicate from history** — button on each history entry to pre-fill the form with its data
- [ ] 6. **Undo last generation** — one-click delete of the most recently created actor/item from inside the builder
- [ ] 7. **Name suggestions** — "Suggest a name" button that generates a random thematic name based on the description without a full generation
- [ ] 8. **Description templates** — dropdown of starter templates ("Town guard", "Merchant", "Cult leader") that pre-fill the description field
- [ ] 9. **Compact history panel** — toggle to collapse the history panel to a narrow strip to give the form more space
- [ ] 10. **Keyboard shortcuts** — `Ctrl+Enter` to generate, `Escape` to clear the form

## Reliability & Error Handling
- [ ] 11. **Offline detection** — detect no network and show a clear banner instead of a confusing auth error
- [ ] 12. **Session pre-validation** — on app open, silently test the stored key and proactively prompt re-auth if invalid before the user clicks Generate
- [ ] 13. **Partial failure recovery** — if actor creation fails after all retries, offer to download the raw JSON so the user doesn't lose generated data
- [ ] 14. **Rate limit countdown** — parse the reset timestamp from 429 responses and show "Limit resets in X days"
- [ ] 15. **Concurrent generation guard** — disable the Generate button (not just via CSS class) while a generation is in-flight

## Generation Quality
- [ ] 16. **Post-generation quick-edit dialog** — after creation, show a lightweight overlay (name, level, key stats) so users can fix small mistakes before the full sheet opens
- [ ] 17. **Variation slider** — a "Creativity" slider sent as a parameter to the endpoint to control how closely the AI follows the description
- [ ] 18. **Bulk generation** — generate multiple NPCs in one request and create them all as separate history entries
- [ ] 19. **Level Up for D&D 5e** — add the same CR-change sheet button that PF2e has, for 5e creatures
- [ ] 20. **Image generation for item modules** — optional portrait/icon generation for PF2e and 5e items

## Developer Experience
- [ ] 21. **Adapter validation on construction** — check all required getters return non-empty values at init time and throw descriptive errors immediately
- [ ] 22. **`Storage.clear()`** — method that wipes all stored keys for a module (useful for testing and troubleshooting)
- [ ] 23. **Build-time type checking** — add `jsconfig.json` with `checkJs: true` so VS Code surfaces type errors in the shared core
- [ ] 24. **Shared test harness** — minimal Vitest test file covering `isNewerVersion`, `escapeHtml`, `devUrl`, and `Storage`

## Performance
- [ ] 25. **Spell mapping cache** — cache the PF2e spell mapping in `sessionStorage` keyed by a hash of installed packs, skipping the 5–10 s build on repeat generations
- [ ] 26. **Template preloading at build time** — inline the Handlebars template into the entry bundle so `loadTemplates` is a no-op at runtime

## Polish & Accessibility
- [ ] 27. **Accessible focus management** — move focus to the new history entry after generation so keyboard users don't lose their place
- [ ] 28. **CSS custom properties for colours** — replace inline hex colours in dialogs with CSS variables that respect Foundry's theme
- [ ] 29. **Localisation scaffold** — wrap all user-facing strings in `game.i18n.localize()` and add a starter `en.json` lang file
- [ ] 30. **Settings page audit** — surface "Reset welcome message" and "Clear session" as buttons in the Foundry module settings UI
