# Battle map model bake-off

Runs the same 20 battle map prompts and 4 edit tests against several image
models, saves every image with its real cost and timing, and gives you a page
to score them side by side with a grid laid over each map.

This is a test tool for choosing the battle map maker's model. It is not
shipped with any module.

## What it costs

About **$7 to $8** for the full run (20 maps and 4 edits on each of 4 models),
based on September 2026 prices. The script stops when spending passes
`--max-cost` (default $15).

## Running it

You need Node 18 or newer and an [OpenRouter](https://openrouter.ai) API key with
a few dollars of credit. One key covers all four models.

```sh
# 1. Check the setup without spending anything (writes placeholder images)
node tools/map-bakeoff/run.mjs --dry-run --out tools/map-bakeoff/out/dry

# 2. Optional: confirm the model ids are still listed
OPENROUTER_API_KEY=sk-or-... node tools/map-bakeoff/run.mjs --list-models

# 3. The real run
OPENROUTER_API_KEY=sk-or-... node tools/map-bakeoff/run.mjs
```

On Windows PowerShell, set the key with `$env:OPENROUTER_API_KEY="sk-or-..."` first.

Output goes to `tools/map-bakeoff/out/<today's date>/` (ignored by git). When it
finishes, open `report.html` in that folder in your browser.

If the run stops partway (network error, cost limit), run the same command
again. Finished images are skipped, so you only pay for what's missing.

## Options

| Option | What it does |
| --- | --- |
| `--dry-run` | No API calls. Writes flat placeholder images so you can check the pipeline and the report page. |
| `--models seedream-lite,gemini-flash` | Only run these models (keys from `models.json`). |
| `--prompts p01,p04` | Only run these prompts (ids or prefixes from `prompts.json`). |
| `--skip-edits` | Only generate new maps. |
| `--max-cost 10` | Stop once this many dollars have been spent. |
| `--concurrency 2` | How many requests run at once. |
| `--out <folder>` | Where to write images and results. |
| `--list-models` | Print the image models OpenRouter currently offers. |

## Scoring

The report page shows each prompt as a row with one image per model. Score each
image from 1 to 5:

- **Detail:** texture, clutter, and readability at table zoom.
- **Perspective:** did it hold the asked-for view (top-down, isometric, or three-quarter)?
- **Grid fit:** with the grid on, do doors, corridors, and furniture sit sensibly on the squares?
- **Matches prompt:** are the described features there?

Edits are scored on whether the change was applied, whether the rest of the
map stayed the same, and whether there are visible seams.

Scores save in your browser. Use **Save scores (JSON)** to keep them, and
**Export scores (CSV)** to put them in a spreadsheet. The summary table at the
top shows cost per map, cost per session (8 maps and 4 edits), pixels per
square, and average scores for each model.

## Files

- `prompts.json`: the 20 test maps (perspective, genre, style, mood, size in squares, description).
- `edits.json`: the 4 edit tests, each applied to one of the maps.
- `models.json`: the models to compare, with OpenRouter ids and estimated prices.
- `lib.mjs`: prompt building and image helpers (tested in `lib.test.mjs`).
- `run.mjs`: the runner.
- `report.html`: the scoring page (copied into each output folder).

## Notes

- Edits in this test send the whole map with an instruction. The real module
  will crop the selected region and blend it back, which should be cheaper and
  keep the rest of the map unchanged.
- Model ids and prices change. If a model fails with "not found", run
  `--list-models` and update `models.json`.
