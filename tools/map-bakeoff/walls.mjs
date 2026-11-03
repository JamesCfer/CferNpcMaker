#!/usr/bin/env node
/**
 * Wall-tracing test. Two ways to find walls on a finished map:
 *
 *   llm:  a vision model returns wall segments directly, and toFoundryWalls
 *         turns them into Foundry walls. Chosen on Nov 3, 2026.
 *   mask: an image model repaints the map as a wall mask (white walls, red
 *         doors, blue windows on black). Gemini made clean masks, but turning
 *         a mask into lines was messier and cost ten times more, so it was dropped.
 *
 *   node tools/map-bakeoff/walls.mjs mask  --images a.png --models black-forest-labs/flux.2-klein-4b
 *   node tools/map-bakeoff/walls.mjs llm   --images a.png --squares 24x18 --models anthropic/claude-haiku-5.5
 *
 * Needs OPENROUTER_API_KEY for mask and llm. Output goes to out/<--out, default "walls">/.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { toFoundryWalls, extFor } from './lib.mjs';
import { WALLS_PROMPT as LLM_PROMPT, WALLS_MASK_PROMPT as MASK_PROMPT } from '../../shared/scripts/core/maps/index.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const [step, ...rest] = process.argv.slice(2);
const args = Object.fromEntries(rest.reduce((a, v, i, all) => (v.startsWith('--') ? [...a, [v.slice(2), all[i + 1]]] : a), []));
const KEY = process.env.OPENROUTER_API_KEY;
const outDir = path.join(here, 'out', args.out || 'walls');
fs.mkdirSync(outDir, { recursive: true });
const images = (args.images || '').split(',').filter(Boolean);
const models = (args.models || '').split(',').filter(Boolean);
const squares = (args.squares || '').split(',').map(s => s.split('x').map(Number));
const dataUri = f => `data:${f.endsWith('.png') ? 'image/png' : 'image/jpeg'};base64,${fs.readFileSync(f).toString('base64')}`;
const slug = m => m.split('/').pop().replace(/[^a-z0-9.-]/gi, '-');
let total = 0;



if (step === 'mask') {
  if (!KEY) die('Set OPENROUTER_API_KEY.');
  await Promise.all(images.flatMap(file => models.map(async model => {
    const t0 = Date.now();
    try {
      const res = await fetch('https://openrouter.ai/api/v1/images', {
        method: 'POST',
        headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map wall test' },
        body: JSON.stringify({ n: 1, model, prompt: MASK_PROMPT, resolution: '2K', aspect_ratio: args.aspect || undefined, input_references: [{ type: 'image_url', image_url: { url: dataUri(file) } }] }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message}`);
      const img = data.data[0];
      const out = path.join(outDir, `${path.basename(file).replace(/\.\w+$/, '')}--mask--${slug(model)}.${extFor(img.media_type || 'image/png')}`);
      fs.writeFileSync(out, Buffer.from(img.b64_json, 'base64'));
      total += data.usage?.cost || 0;
      console.log(`  ok   ${path.basename(out)}: $${(data.usage?.cost ?? 0).toFixed(4)}, ${((Date.now() - t0) / 1000).toFixed(1)}s`);
    } catch (err) { console.log(`  FAIL ${path.basename(file)} ${model}: ${err.message}`); }
  })));
} else if (step === 'llm') {
  if (!KEY) die('Set OPENROUTER_API_KEY.');
  const results = [];
  await Promise.all(images.flatMap((file, i) => models.map(async model => {
    const t0 = Date.now();
    try {
      const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map wall test' },
        body: JSON.stringify({ model, response_format: { type: 'json_object' }, usage: { include: true }, messages: [{ role: 'user', content: [{ type: 'text', text: LLM_PROMPT }, { type: 'image_url', image_url: { url: dataUri(file) } }] }] }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message}`);
      const text = data.choices[0].message.content.trim().replace(/^```(?:json)?\s*/, '').replace(/```$/, '');
      const walls = JSON.parse(text).walls || [];
      total += data.usage?.cost || 0;
      results.push({ file: path.basename(file), model, squares: squares[i], costUsd: data.usage?.cost, ms: Date.now() - t0, walls, foundry: toFoundryWalls(walls, { squares: squares[i] }) });
      console.log(`  ok   ${path.basename(file)} ${model}: ${walls.length} segments, $${(data.usage?.cost ?? 0).toFixed(4)}, ${((Date.now() - t0) / 1000).toFixed(1)}s`);
    } catch (err) { console.log(`  FAIL ${path.basename(file)} ${model}: ${err.message}`); }
  })));
  fs.writeFileSync(path.join(outDir, 'llm-walls.json'), JSON.stringify(results, null, 2));
} else die('First argument must be mask or llm.');
console.log(`Done. $${total.toFixed(4)} spent.`);

function die(msg) { console.error(msg); process.exit(1); }
