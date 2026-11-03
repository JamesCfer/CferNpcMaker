#!/usr/bin/env node
/**
 * Scale test: a vision model estimates the real-world size of a map from
 * objects of known size (doors, chairs, beds, tables), and sceneScale in
 * lib.mjs turns that into Foundry scene and grid settings.
 *
 *   node tools/map-bakeoff/scale.mjs --images a.png,b.png --expect 24x18,30x20 --models google/gemini-3.7-flash --repeat 2
 *
 * --expect is the size in five-foot squares the map was asked for, used only
 * to score the estimate. Needs OPENROUTER_API_KEY. Writes out/<--out, default "scale">/scale.json.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { sceneScale } from './lib.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).reduce((a, v, i, all) => (v.startsWith('--') ? [...a, [v.slice(2), all[i + 1]]] : a), []));
const KEY = process.env.OPENROUTER_API_KEY;
if (!KEY) { console.error('Set OPENROUTER_API_KEY.'); process.exit(1); }
const images = (args.images || '').split(',').filter(Boolean);
const expect = (args.expect || '').split(',').map(s => s.split('x').map(Number));
const models = (args.models || 'google/gemini-3.7-flash').split(',');
const repeat = Number(args.repeat || 1);
const outDir = path.join(here, 'out', args.out || 'scale');
fs.mkdirSync(outDir, { recursive: true });

const PROMPT = [
  'This is a tabletop RPG battle map. Estimate the real-world size of the whole area shown in the image, edge to edge.',
  'Use objects of known size as rulers: a single door is about 0.9 m wide, a chair about 0.5 m, a bed about 2 m long, a round tavern table about 1.2 m across, a cart about 3 m long, a person about 0.5 m across from above.',
  'Measure at least three different objects, work out metres per pixel from each, and combine them.',
  'Return JSON only: {"references":[{"object":"...","size_m":0.9,"share_of_width":0.03}],"width_m":<number>,"height_m":<number>}',
  'share_of_width is how much of the image width the object spans.',
].join('\n');

let total = 0;
const results = [];
await Promise.all(images.flatMap((file, i) => models.flatMap(model => Array.from({ length: repeat }, async (_, r) => {
  const mime = file.endsWith('.png') ? 'image/png' : 'image/jpeg';
  const t0 = Date.now();
  try {
    const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map scale test' },
      body: JSON.stringify({ model, response_format: { type: 'json_object' }, usage: { include: true }, messages: [{ role: 'user', content: [
        { type: 'text', text: PROMPT },
        { type: 'image_url', image_url: { url: `data:${mime};base64,${fs.readFileSync(file).toString('base64')}` } },
      ] }] }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message}`);
    const est = JSON.parse(data.choices[0].message.content.trim().replace(/^```(?:json)?\s*/, '').replace(/```$/, ''));
    total += data.usage?.cost || 0;
    const imperial = sceneScale(est, { system: 'imperial' });
    const metric = sceneScale(est, { system: 'metric' });
    const want = expect[i];
    const errPct = want ? Math.round(((imperial.squares[0] / want[0]) - 1) * 100) : null;
    results.push({ file: path.basename(file), model, run: r + 1, costUsd: data.usage?.cost, ms: Date.now() - t0, estimate: est, imperial, metric, expected: want, widthErrorPct: errPct });
    console.log(`  ok   ${path.basename(file)} ${model} #${r + 1}: ${est.width_m.toFixed(1)} x ${est.height_m.toFixed(1)} m = ${imperial.squares.join('x')} squares (asked ${want?.join('x')}, ${errPct > 0 ? '+' : ''}${errPct}%), $${(data.usage?.cost ?? 0).toFixed(4)}`);
  } catch (err) {
    console.log(`  FAIL ${path.basename(file)} ${model} #${r + 1}: ${err.message}`);
  }
}))));
fs.writeFileSync(path.join(outDir, 'scale.json'), JSON.stringify(results, null, 2));
console.log(`Done. $${total.toFixed(4)} spent.`);
