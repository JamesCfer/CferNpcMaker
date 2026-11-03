#!/usr/bin/env node
/**
 * Mockup-first test: a light model draws 4 quick mockups per prompt, someone
 * picks one, then the quality models build the final 2K map from that mockup.
 *
 *   node tools/map-bakeoff/mockups.mjs mockups --prompts p02-tavern,p05-spaceship
 *   node tools/map-bakeoff/mockups.mjs finals  --pick p02-tavern=2,p05-spaceship=4
 *   node tools/map-bakeoff/mockups.mjs dials   --pick p02-tavern=2 --models final-gemini --variants "detail=1;detail=5;color=1,tint=5"
 *
 * Needs OPENROUTER_API_KEY. Output goes to tools/map-bakeoff/out/mockups/, or out/<name>/ with --out <name>.
 * Prompts come from prompts.json and towns.json.
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
const ALL_FINALS = [
  { key: 'final-gemini', model: 'google/gemini-3.1-flash-image', resolution: '2K', label: 'Standard (Gemini)' },
  { key: 'final-seedream-pro', model: 'bytedance-seed/seedream-5-0-pro', resolution: '2K', label: 'High detail (Seedream Pro)' },
  { key: 'final-gemini-4k', model: 'google/gemini-3.1-flash-image', resolution: '4K', label: 'Upscale candidate (Gemini at 4K)' },
  { key: 'final-gemini-1k', model: 'google/gemini-3.1-flash-image', resolution: '1K', label: 'Budget candidate (Gemini at 1K)' },
  { key: 'final-seedream-lite', model: 'bytedance-seed/seedream-5-0-lite', resolution: '2K', label: 'Budget candidate (Seedream Lite)' },
  { key: 'final-klein-2k', model: 'black-forest-labs/flux.2-klein-4b', resolution: '2K', label: 'Budget candidate (FLUX.2 klein at 2K)' },
];
// --models final-gemini,final-seedream-lite picks which finals to run (default: the two launch options).
const FINALS = ALL_FINALS.filter(m => (args.models ? args.models.split(',') : ['final-gemini', 'final-seedream-pro']).includes(m.key));

const outDir = path.join(here, 'out', args.out || 'mockups');
fs.mkdirSync(outDir, { recursive: true });
const statePath = path.join(outDir, 'results.json');
const state = fs.existsSync(statePath) ? JSON.parse(fs.readFileSync(statePath, 'utf8')) : { results: [] };
const prompts = ['prompts.json', 'towns.json'].flatMap(f => JSON.parse(fs.readFileSync(path.join(here, f), 'utf8')));
const byId = Object.fromEntries(prompts.map(p => [p.id, p]));
let spent = 0;

if (step === 'mockups') {
  const ids = (args.prompts || '').split(',').filter(Boolean);
  if (!ids.length) die('Pass --prompts id1,id2');
  state.prompts = ids;
  const styleRef = args['style-ref'] ? `data:image/${args['style-ref'].endsWith('.png') ? 'png' : 'jpeg'};base64,${fs.readFileSync(args['style-ref']).toString('base64')}` : null;
  // --dials "clutter=1;clutter=5" makes one set of mockups per dial variant, named <id>--<variant>.
  const variants = args.dials ? args.dials.split(';').map(v => Object.fromEntries(v.split(',').map(kv => { const [k, x] = kv.split('='); return [k, Number(x)]; }))) : [null];
  await Promise.all(ids.flatMap(baseId => variants.map(async dials => {
    const p = byId[baseId] || die(`Unknown prompt ${baseId}`);
    let id = dials ? `${baseId}--${Object.entries(dials).map(([k, v]) => `${k}${v}`).join('-')}` : baseId;
    // --view plan|sat|combo tries another top-down wording; --style-ref <image> adds a reference for the camera angle only.
    if (args.view) id += `--view-${args.view}`;
    if (styleRef) id += '--ref';
    const prompt = [
      'Quick rough mockup for layout only.',
      ...(styleRef ? ['Use the attached image only as a reference for the straight-down camera angle and how flat walls and furniture look. Do not copy its content, layout, or colors.'] : []),
      buildPrompt({ ...p, ...(dials ? { dials } : {}), ...(args.view ? { view: args.view } : {}) }),
    ].join(' ');
    await Promise.all([1, 2, 3, 4].map(n => gen({ kind: 'mockup', id, n, model: LIGHT, prompt, aspect: pickAspect(p.squares), squares: p.squares, ref: styleRef, ...(dials ? { dials } : {}), ...(args.view ? { view: args.view } : {}) })));
  })));
} else if (step === 'finals') {
  const picks = Object.fromEntries((args.pick || '').split(',').filter(Boolean).map(s => s.split('=')));
  if (!Object.keys(picks).length) die('Pass --pick id=n,...');
  state.picks = picks;
  const jobs = [];
  for (const [id, n] of Object.entries(picks)) {
    const mock = state.results.find(r => r.kind === 'mockup' && r.id === id && r.n === Number(n) && r.ok);
    if (!mock) die(`No mockup ${n} for ${id}`);
    // A mockup made with --dials is named <id>--<variant>; the final keeps its content dials.
    const p = { ...byId[id.split('--')[0]], ...(mock.dials ? { dials: mock.dials } : {}), ...(mock.view ? { view: mock.view } : {}) };
    const bytes = fs.readFileSync(path.join(outDir, mock.file));
    const ref = `data:${mock.mediaType};base64,${bytes.toString('base64')}`;
    for (const m of FINALS) jobs.push(gen({ kind: 'final', id, n: Number(n), model: m, prompt: finalPrompt(p), aspect: mock.aspect, squares: p.squares, ref }));
  }
  await Promise.all(jobs);
} else if (step === 'dials') {
  // One final per dial variant, all from the same picked mockup, so only the look changes.
  const [id, n] = (args.pick || '').split('=');
  const p = byId[id] || die('Pass --pick id=n');
  const mock = state.results.find(r => r.kind === 'mockup' && r.id === id && r.n === Number(n) && r.ok) || die(`No mockup ${n} for ${id}`);
  const ref = `data:${mock.mediaType};base64,${fs.readFileSync(path.join(outDir, mock.file)).toString('base64')}`;
  const variants = (args.variants || '').split(';').filter(Boolean).map(v => Object.fromEntries(v.split(',').map(kv => { const [k, x] = kv.split('='); return [k, Number(x)]; })));
  if (!variants.length) die('Pass --variants "detail=1;detail=5"');
  await Promise.all(variants.flatMap(dials => {
    const slug = Object.entries(dials).map(([k, v]) => `${k}${v}`).join('-');
    return FINALS.map(m => gen({ kind: 'dial', id: `${id}--${slug}`, n: Number(n), model: m, prompt: finalPrompt({ ...p, dials }), aspect: mock.aspect, squares: p.squares, ref, dials }));
  }));
} else die('First argument must be "mockups", "finals", or "dials".');

function finalPrompt(p) {
  return [
    'The attached image is a rough layout mockup. Redraw it as a finished battle map.',
    'Keep the same layout: room shapes, walls, doors, paths, and where every feature sits, and keep the same camera angle.',
    'Ignore and remove any grid lines, text, or labels in the mockup.',
    'Keep every object the same size relative to the map as in the mockup, and never enlarge furniture, buildings, or other objects.',
    ...(p.perspective === 'top-down' ? ['If the mockup is tilted or shows the sides of walls, flatten it into a true straight-down view.'] : []),
    buildPrompt(p),
  ].join(' ');
}

fs.writeFileSync(statePath, JSON.stringify(state, null, 2));
console.log(`Done. $${spent.toFixed(3)} spent this step.`);

async function gen({ kind, id, n, model, prompt, aspect, squares, ref, dials, view }) {
  const t0 = Date.now();
  const body = { n: 1, model: model.model, prompt, aspect_ratio: aspect, resolution: model.resolution };
  if (ref) body.input_references = [{ type: 'image_url', image_url: { url: ref } }];
  const rec = { kind, id, n, modelKey: model.key, model: model.model, aspect, squares, ...(dials ? { dials } : {}), ...(view ? { view } : {}) };
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
