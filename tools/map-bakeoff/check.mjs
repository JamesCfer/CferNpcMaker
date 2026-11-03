#!/usr/bin/env node
/**
 * Defect check: a vision model looks at a finished map for things Foundry
 * maps must not have (a drawn grid, labels or text, a frame or blank margin),
 * so the module can retry the map before the GM sees it.
 *
 *   node tools/map-bakeoff/check.mjs --images a.png,b.jpg --models google/gemini-3.7-flash
 *
 * Needs OPENROUTER_API_KEY. Writes out/<--out, default "check">/check.json.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { mapDefects } from './lib.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).reduce((a, v, i, all) => (v.startsWith('--') ? [...a, [v.slice(2), all[i + 1]]] : a), []));
const KEY = process.env.OPENROUTER_API_KEY;
if (!KEY) { console.error('Set OPENROUTER_API_KEY.'); process.exit(1); }
const images = (args.images || '').split(',').filter(Boolean);
const models = (args.models || 'google/gemini-3.7-flash').split(',');
const prompts = ['prompts.json', 'towns.json'].flatMap(f => JSON.parse(fs.readFileSync(path.join(here, f), 'utf8')));
// The perspective comes from the prompt id at the start of the file name, else --perspective.
const perspectiveOf = file => prompts.find(p => path.basename(file).startsWith(p.id))?.perspective || args.perspective || 'top-down';
const outDir = path.join(here, 'out', args.out || 'check');
fs.mkdirSync(outDir, { recursive: true });

const PROMPT = [
  'You are checking an AI-generated tabletop RPG battle map before it goes into Foundry VTT, which adds its own grid.',
  'Answer each question about the image:',
  'grid: are there square or hex grid lines drawn over the map? Floor tiles, paving stones, deck planks, or floor grates that are part of the scenery do not count. Only a regular overlay of thin lines across the map counts, including faint ones.',
  'text: is there any label, number, title, legend, scale bar, dimension, signature, or watermark? Writing that is part of the scene, such as a shop sign or a neon sign, does not count.',
  'frame: is there a drawn decorative border, picture frame, parchment edge, or blueprint sheet around the map? The outer walls of a building or room do not count.',
  'margin: is the map drawn on plain paper, parchment, or a white, grey, or light background, with that blank area showing around it? Dark empty space or rock outside the walls of an interior is normal and does not count. Give the share of the image that is such blank background, from 0 to 1.',
  'Return JSON only: {"grid":true|false,"text":true|false,"frame":true|false,"margin":0.0,"notes":"one short sentence"}',
].join('\n');

let total = 0;
const results = [];
await Promise.all(images.flatMap(file => models.map(async model => {
  const mime = file.endsWith('.png') ? 'image/png' : 'image/jpeg';
  const t0 = Date.now();
  try {
    const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map defect check' },
      body: JSON.stringify({ model, response_format: { type: 'json_object' }, usage: { include: true }, messages: [{ role: 'user', content: [
        { type: 'text', text: PROMPT },
        { type: 'image_url', image_url: { url: `data:${mime};base64,${fs.readFileSync(file).toString('base64')}` } },
      ] }] }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message}`);
    const raw = JSON.parse(data.choices[0].message.content.trim().replace(/^```(?:json)?\s*/, '').replace(/```$/, ''));
    total += data.usage?.cost || 0;
    const verdict = mapDefects(raw, { perspective: perspectiveOf(file) });
    results.push({ file, model, costUsd: data.usage?.cost, ms: Date.now() - t0, raw, verdict });
    console.log(`  ok   ${file} ${model}: ${verdict.retry ? 'RETRY ' + verdict.reasons.join(',') : 'clean'} (${raw.notes}) $${(data.usage?.cost ?? 0).toFixed(4)}`);
  } catch (err) {
    console.log(`  FAIL ${file} ${model}: ${err.message}`);
  }
})));
fs.writeFileSync(path.join(outDir, 'check.json'), JSON.stringify(results, null, 2));
console.log(`Done. $${total.toFixed(4)} spent.`);
