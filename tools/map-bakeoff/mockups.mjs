#!/usr/bin/env node
/**
 * Mockup-first test: a light model draws 4 quick mockups per prompt, someone
 * picks one, then the quality models build the final 2K map from that mockup.
 *
 *   node tools/map-bakeoff/mockups.mjs mockups --prompts p02-tavern,p05-spaceship
 *   node tools/map-bakeoff/mockups.mjs finals  --pick p02-tavern=2,p05-spaceship=4
 *
 * Needs OPENROUTER_API_KEY. Output goes to tools/map-bakeoff/out/mockups/.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildPrompt, pickAspect, extFor, imageSize } from './lib.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const API = 'https://openrouter.ai/api/v1';
const KEY = process.env.OPENROUTER_API_KEY;
const [step, ...rest] = process.argv.slice(2);
const args = Object.fromEntries(rest.reduce((a, v, i, all) => (v.startsWith('--') ? [...a, [v.slice(2), all[i + 1]]] : a), []));
if (!KEY) die('Set OPENROUTER_API_KEY.');

const LIGHT = { key: 'mock-klein', model: 'black-forest-labs/flux.2-klein-4b', resolution: '1K' };
const FINALS = [
  { key: 'final-gemini', model: 'google/gemini-3.1-flash-image', resolution: '2K', label: 'Standard (Gemini)' },
  { key: 'final-seedream-pro', model: 'bytedance-seed/seedream-5-0-pro', resolution: '2K', label: 'High detail (Seedream Pro)' },
];

const outDir = path.join(here, 'out', 'mockups');
fs.mkdirSync(outDir, { recursive: true });
const statePath = path.join(outDir, 'results.json');
const state = fs.existsSync(statePath) ? JSON.parse(fs.readFileSync(statePath, 'utf8')) : { results: [] };
const prompts = JSON.parse(fs.readFileSync(path.join(here, 'prompts.json'), 'utf8'));
const byId = Object.fromEntries(prompts.map(p => [p.id, p]));
let spent = 0;

if (step === 'mockups') {
  const ids = (args.prompts || '').split(',').filter(Boolean);
  if (!ids.length) die('Pass --prompts id1,id2');
  state.prompts = ids;
  await Promise.all(ids.map(async id => {
    const p = byId[id] || die(`Unknown prompt ${id}`);
    const prompt = `Quick rough mockup for layout only. ${buildPrompt(p)}`;
    await Promise.all([1, 2, 3, 4].map(n => gen({ kind: 'mockup', id, n, model: LIGHT, prompt, aspect: pickAspect(p.squares), squares: p.squares })));
  }));
} else if (step === 'finals') {
  const picks = Object.fromEntries((args.pick || '').split(',').filter(Boolean).map(s => s.split('=')));
  if (!Object.keys(picks).length) die('Pass --pick id=n,...');
  state.picks = picks;
  const jobs = [];
  for (const [id, n] of Object.entries(picks)) {
    const p = byId[id];
    const mock = state.results.find(r => r.kind === 'mockup' && r.id === id && r.n === Number(n) && r.ok);
    if (!mock) die(`No mockup ${n} for ${id}`);
    const bytes = fs.readFileSync(path.join(outDir, mock.file));
    const ref = `data:${mock.mediaType};base64,${bytes.toString('base64')}`;
    const prompt = [
      'The attached image is a rough layout mockup. Redraw it as a finished, highly detailed battle map.',
      'Keep the same layout: room shapes, walls, doors, paths, and where every feature sits, and keep the same camera angle.',
      'Ignore and remove any grid lines, text, or labels in the mockup.',
      buildPrompt(p),
    ].join(' ');
    for (const m of FINALS) jobs.push(gen({ kind: 'final', id, n: Number(n), model: m, prompt, aspect: mock.aspect, squares: p.squares, ref }));
  }
  await Promise.all(jobs);
} else die('First argument must be "mockups" or "finals".');

fs.writeFileSync(statePath, JSON.stringify(state, null, 2));
console.log(`Done. $${spent.toFixed(3)} spent this step.`);

async function gen({ kind, id, n, model, prompt, aspect, squares, ref }) {
  const t0 = Date.now();
  const body = { n: 1, model: model.model, prompt, aspect_ratio: aspect, resolution: model.resolution };
  if (ref) body.input_references = [{ type: 'image_url', image_url: { url: ref } }];
  const rec = { kind, id, n, modelKey: model.key, model: model.model, aspect, squares };
  try {
    const res = await fetch(`${API}/images`, { method: 'POST', headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map mockup test' }, body: JSON.stringify(body) });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message || JSON.stringify(data).slice(0, 300)}`);
    const img = data.data?.[0];
    if (!img?.b64_json) throw new Error('Response had no image.');
    const bytes = Buffer.from(img.b64_json, 'base64');
    const mediaType = img.media_type || 'image/png';
    const file = path.join(model.key, `${id}${kind === 'mockup' ? '-' + n : ''}.${extFor(mediaType)}`);
    fs.mkdirSync(path.join(outDir, model.key), { recursive: true });
    fs.writeFileSync(path.join(outDir, file), bytes);
    const costUsd = typeof data.usage?.cost === 'number' ? data.usage.cost : null;
    spent += costUsd || 0;
    Object.assign(rec, { ok: true, file, mediaType, ...(imageSize(bytes) || {}), costUsd, ms: Date.now() - t0 });
    console.log(`  ok   ${kind} ${id}${kind === 'mockup' ? ' #' + n : ''} ${model.key}: $${(costUsd ?? 0).toFixed(4)}, ${((Date.now() - t0) / 1000).toFixed(1)}s`);
  } catch (err) {
    Object.assign(rec, { ok: false, error: String(err.message || err), ms: Date.now() - t0 });
    console.log(`  FAIL ${kind} ${id} ${model.key}: ${rec.error}`);
  }
  state.results = state.results.filter(r => !(r.kind === kind && r.id === id && r.n === n && r.modelKey === model.key));
  state.results.push(rec);
}

function die(msg) { console.error(msg); process.exit(1); }
