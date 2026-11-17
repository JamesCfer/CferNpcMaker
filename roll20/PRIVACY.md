# Privacy policy: Cfer NPC Maker for Roll20

Last updated: 17 November 2026

This policy covers the "Cfer NPC Maker for Roll20" browser extension and the optional `CferImporter.js` Roll20 Mod script, made by CelestiaTools (James Cfer).

## What the extension does

It adds a panel to the Roll20 editor. You describe an NPC, item, shop or battle map; the extension sends that request to the CelestiaTools generation server, then writes the result into your Roll20 game.

## Data the extension handles

| Data | Why | Where it goes |
| --- | --- | --- |
| A sign-in key issued after you sign in with Patreon | To check your Patreon membership and monthly uses | Stored in your browser's extension storage. Sent only to the CelestiaTools server (`foundryrelay.dedicated2.com`). |
| What you type into the generator forms (names, descriptions, levels and similar options) | To generate the content you asked for | Sent to the CelestiaTools server, which passes it to AI model providers to produce the result. |
| Your remaining uses for the month | To show the uses counter | Stored in your browser's extension storage. |

When you sign in, Patreon tells the CelestiaTools server your Patreon membership status, tier and the email address on your Patreon account. The server uses these only to check which tier you are on and to count your uses. The extension itself never sees your Patreon password.

The extension reads the list of characters in the Roll20 game you have open so it can offer "give item to character". That list stays in your browser and is not sent anywhere.

## What we do not do

* We do not sell or share your data with third parties, apart from the AI model providers that generate content from your prompt.
* We do not use your data for advertising or credit decisions.
* We do not track your browsing. The extension only runs on `app.roll20.net/editor` pages.

## Generated content

Generated results (for example an NPC stat block or a battle map image) may be kept in the CelestiaTools server logs and posted to a CelestiaTools Discord channel used to monitor quality. These posts contain the generated content only, not your name, email or Patreon details.

## Removing your data

Choose **Sign out** in the extension's Settings tab, or remove the extension, to delete the stored key and usage data from your browser. To ask for your server-side records to be removed, message CelestiaTools on Patreon: https://www.patreon.com/cw/CelestiaTools

## Changes

If this policy changes, the new version will be posted at this address with a new date.
