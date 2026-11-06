#!/usr/bin/env node
/**
 * Light-placement test: a vision model finds light sources on a finished map
 * and returns them as Foundry ambient lights.
 *
 *   node tools/map-bakeoff/lights.mjs --images a.png,b.jpg --squares 24x18,40x30 --models google/gemini-3.7-flash
 *
 * Needs OPENROUTER_API_KEY. Writes out/<name>/lights.json (default name "lights").
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { toFoundryLights } from './lib.mjs';
import { LIGHTS_PROMPT as PROMPT } from '../../shared/scripts/core/maps/index.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).reduce((a, v, i, all) => (v.startsWith('--') ? [...a, [v.slice(2), all[i + 1]]] : a), []));
const KEY = process.env.OPENROUTER_API_KEY;
if (!KEY) { console.error('Set OPENROUTER_API_KEY.'); process.exit(1); }
const images = (args.images || '').split(',').filter(Boolean);
const squares = (args.squares || '').split(',').map(s => s.split('x').map(Number));
const models = (args.models || 'google/gemini-3.7-flash').split(',');
const outDir = path.join(here, 'out', args.out || 'lights');
fs.mkdirSync(outDir, { recursive: true });


let total = 0;
const results = [];
for (const [i, file] of images.entries()) {
  const bytes = fs.readFileSync(file);
  const mime = file.endsWith('.png') ? 'image/png' : 'image/jpeg';
  await Promise.all(models.map(async model => {
    const t0 = Date.now();
    try {
      const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map light test' },
        body: JSON.stringify({
          model,
          response_format: { type: 'json_object' },
          usage: { include: true },
          messages: [{ role: 'user', content: [
            { type: 'text', text: PROMPT },
            { type: 'image_url', image_url: { url: `data:${mime};base64,${bytes.toString('base64')}` } },
          ] }],
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message}`);
      const text = data.choices[0].message.content.trim().replace(/^```(?:json)?\s*/, '').replace(/```$/, '');
      const raw = JSON.parse(text).lights || [];
      const cost = data.usage?.cost ?? null;
      total += cost || 0;
      const lights = toFoundryLights(raw, { squares: squares[i] });
      results.push({ file: path.basename(file), model, squares: squares[i], ms: Date.now() - t0, costUsd: cost, raw, lights });
      console.log(`  ok   ${path.basename(file)} ${model}: ${raw.length} lights, $${(cost ?? 0).toFixed(4)}, ${((Date.now() - t0) / 1000).toFixed(1)}s`);
    } catch (err) {
      results.push({ file: path.basename(file), model, error: String(err.message || err) });
      console.log(`  FAIL ${path.basename(file)} ${model}: ${err.message}`);
    }
  }));
}
fs.writeFileSync(path.join(outDir, 'lights.json'), JSON.stringify(results, null, 2));
console.log(`Done. $${total.toFixed(4)} spent.`);
