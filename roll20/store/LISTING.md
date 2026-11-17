# Store upload kit: Cfer NPC Maker for Roll20

Everything to paste into the Chrome Web Store and Firefox Add-ons forms. Upload file: `roll20/dist/cfer-roll20-extension.zip` (run `npm run build:roll20`; a copy also lives in project files `roll20/`).

Privacy policy URL: https://github.com/JamesCfer/CferNpcMaker/blob/main/roll20/PRIVACY.md
Homepage / support URL: https://www.patreon.com/cw/CelestiaTools

## Shared text

**Name:** Cfer NPC Maker for Roll20

**Summary (132 characters max):**
AI-generated NPCs, magic items, shops, Hero 6e characters and battle maps, imported straight into your Roll20 game.

**Category:** Chrome: Fun (or Productivity). Firefox: Games & Entertainment.

**Description:**

```
Stop typing stat blocks by hand. Describe what you need and Cfer NPC Maker builds it and drops it straight into your Roll20 game.

WHAT IT MAKES
• D&D 5e NPCs: full stat blocks on the D&D 5E by Roll20 sheet and the D&D 2024 sheet, with attacks, traits, saves, skills and spells.
• Pathfinder 2e NPCs on the Pathfinder Second Edition by Roll20 sheet.
• Magic items for D&D 5e and PF2e, as handouts with a one-click "give to character" button.
• Shops for D&D 5e and PF2e: a shopkeeper, staff and a full stock list, each item ready to hand out.
• Hero System 6e characters, ready for the HDImporter Mod script.
• AI battle maps: pick a mockup, generate the final map and place it on your page at the right scale.

HOW IT WORKS
1. Open your Roll20 game and click the round Cfer button at the bottom right.
2. Sign in with Patreon.
3. Choose a generator, describe what you want, and press Generate.
4. The result appears in your game's Journal or on the map.

Works without Roll20 Pro. Pro users can also install the free CferImporter Mod script, which adds walls, doors, windows and lights to battle maps.

PRICING
The extension is free to install. Generating content uses your monthly uses from the CelestiaTools Patreon: the free tier includes a few uses each month to try it out, and paid tiers ($3, $8 and $15 a month) add more. The same uses work in the Cfer modules for Foundry VTT.

Not affiliated with or endorsed by Roll20.
```

## Chrome Web Store: privacy tab

**Single purpose:**
Generates tabletop RPG content (NPCs, items, shops and battle maps) from a text description and imports it into the user's Roll20 game.

**Permission justifications:**

* `storage`: Saves the user's sign-in key and their remaining monthly uses so they stay signed in between sessions.
* `downloads`: Lets the user save a generated result as a file (for example a battle map image or a Hero Designer .hdc file) when it cannot be imported directly.
* Host permission `https://foundryrelay.dedicated2.com/*`: The extension's own generation and sign-in server. All generation requests and Patreon sign-in go there.
* Content scripts on `https://app.roll20.net/editor/*`: Shows the generator panel inside the Roll20 editor and writes the generated characters, handouts and map images into the open game.

**Remote code:** No, I am not using remote code. (All scripts are inside the package. The server only returns JSON data and images.)

**Data usage: tick these boxes**
* Authentication information (the sign-in key)
* Website content (reads the character list in the open Roll20 game for "give item"; it stays in the browser)

Then tick all three certifications (no selling, no unrelated use, no creditworthiness use).

**Screenshots:** `store/screenshot-*.png` (1280x800).
**Small promo tile:** `store/promo-small-440x280.png`. **Marquee (optional):** `store/promo-marquee-1400x560.png`.

**Distribution:** Public, all regions. Free (payment happens on Patreon, not in the store).

## Firefox Add-ons

* Choose **On this site** (listed), then upload the same zip. It passes `web-ext lint` with 0 errors.
* "Do you need to submit source code?" **No.** Nothing is minified or bundled; the zip is the source.
* Data collection is already declared in the manifest (`authenticationInfo`).
* Paste the same name, summary, description, privacy policy URL and screenshots.
* Notes to reviewer: "Sign-in uses Patreon. A free Patreon membership includes test uses. The panel only appears on app.roll20.net/editor pages."

## Edge (optional, free)

Microsoft Partner Center accepts the same zip and text: https://partner.microsoft.com/dashboard/microsoftedge
