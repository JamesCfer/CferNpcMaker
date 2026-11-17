# Cfer Tools for Roll20

The Roll20 version of the Cfer modules. It has two parts:

1. **A free browser extension** (Chrome, Edge, Firefox). It adds a Cfer panel to the Roll20 editor, signs in with Patreon, calls the same n8n generators the Foundry modules call, and writes the result into the game.
2. **An optional free Mod script, `CferImporter.js`**, for Pro games. With it installed, imports go through Roll20's official Mod API, and battle maps get walls, doors, windows and lights.

Uses come out of the same Patreon allowance as the Foundry modules. Requests are tagged `X-Foundry-Origin: roll20-extension` and `client: "roll20"`, so Roll20 usage can be told apart in n8n.

## What each generator makes on Roll20

| Generator | Roll20 result |
| --- | --- |
| D&D 5e NPC | A character on the **D&D 5E by Roll20** (2014) sheet, built by the sheet's own compendium drop, with spells as spell rows. On the **D&D 2024** sheet it fills the NPC stat block (abilities, AC, HP, saves, skills, defenses, senses, languages, traits, actions and attacks). |
| PF2e NPC | A character on the **Pathfinder Second Edition by Roll20** sheet, built by the sheet's own NPC drop. |
| D&D 5e / PF2e item | A handout with the item card, plus a "give" button that adds it to a character's inventory. |
| D&D 5e / PF2e store | A store handout (shopkeeper, staff, stock table) and one handout per item, each with a give button. |
| Nation, city or town | Hidden for now: n8n has no `city-builder` or `nation-builder` workflow yet (the Foundry module quietly falls back to its offline stub). The mapper is ready for when one exists. |
| Hero 6e NPC | A HeroSystem6eHeroic character with its characteristics, a button that sends the full build to the existing **HDImporter** Mod script, and the `.hdc` file. |
| AI Battle Map Maker | Mockups and final map as in Foundry. Place image uploads it to Roll20 and puts it on the map layer at 70px a square, growing the page to fit. The Mod script adds walls, doors, windows and lights. |

The PF2e calendar module is not ported: Roll20 has no calendar to plug into.

## How imports reach the game

The panel picks the first that works (Settings tab can force one):

1. **Mod script**: if `!cfer-ping` gets an answer, the plan is sent in chat as `!cfer-import` chunks and the Mod script builds it.
2. **In-page**: Roll20's own in-page campaign objects (`window.Campaign`), the same ones the editor uses. No Pro needed.
3. **File**: a `.cfer.json` download.

Every NPC import tries the sheet's own compendium-drop import first, checks that it ran, and writes the statblock directly if it did not.

## Layout

```
roll20/
  core/           pure mappers: generator JSON -> import plan (shared by extension and tests)
    catalog.js    generators, forms, endpoints, costs
    plan.js       plan format and chat transport
    mappers/      dnd5e-npc, pf2e-npc, items, store, settlement, hero6e, map
    data/         SRD 5e spell list and PF2e spell list for spell mapping
  extension/      MV3 extension: background.js (holds the key, makes all calls),
                  content.js + panel.js (UI), page-bridge.js (in-page writer)
  mod/            CferImporter.js and script.json (Roll20 one-click library format)
  tests/          vitest: mappers, plan transport, Mod script against a fake Roll20 API
  tools/          build-spell-data.py (rebuilds core/data from the Foundry system repos)
  build.mjs       npm run build:roll20 -> roll20/dist/
```

## Build and try it

```
npm run build:roll20
```

Then in Chrome open `chrome://extensions`, turn on Developer mode, choose **Load unpacked** and pick `roll20/dist/cfer-roll20-extension`. Open a Roll20 game; the **Cfer** button sits at the bottom right.

For the Mod script in a Pro game: Game Settings, Mod (API) Scripts, New Script, paste `roll20/dist/CferImporter.js`, Save. Type `!cfer-help` in chat to check it.

## Tested so far

* Unit tests for every mapper, the chat transport, and the Mod script against a fake Roll20 API.
* Live in a Roll20 game (D&D 2024 sheet, no Pro) on 2026-11-17: Patreon sign-in, uses pill, D&D 5e NPC onto the 2024 stat block, 5e item handout and give list, 5e store (14 items, 15 handouts), PF2e NPC and Hero 6e NPC imports, battle map mockup, final and image placement.

Known limits:

* D&D 2024 sheet: current HP stays at 0 until the GM sets it; giving an item to a PC on the 2024 sheet is not supported yet (NPCs get it added to Gear).
* The Mod script path and the 2014 sheet path have not been tried live yet (that needs a Pro game).
* Stores need a description: the store workflows reject an empty one, after the use has been counted.

## What James needs to do

1. Optional: try the Mod script once in a Pro game.
2. Publish the extension: Chrome Web Store developer account (one-time $5) and Firefox Add-ons (free). Upload `roll20/dist/cfer-roll20-extension.zip` to both.
3. Optional: submit the Mod script to the Roll20 one-click library with a pull request to `Roll20/roll20-api-scripts` (folder `CferImporter/1.0.0/` plus `script.json`; fill in `roll20userid`).
