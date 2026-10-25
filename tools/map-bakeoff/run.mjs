#!/usr/bin/env node
/**
 * Battle map model bake-off runner.
 *
 * Sends the same test prompts to several image models through OpenRouter's
 * Images API, saves every image, and records real cost, time, and size so the
 * models can be compared on the scoring page (report.html).
 *
 *   OPENROUTER_API_KEY=... node tools/map-bakeoff/run.mjs
 *   node tools/map-bakeoff/run.mjs --dry-run          # no API calls, placeholder images
 *
 * See README.md for every option.
 */

import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { buildPrompt, pickAspect, imageSize, extFor } from './lib.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const API = 'https://openrouter.ai/api/v1';

const args = parseArgs(process.argv.slice(2));
if (args.help) { console.log(fs.readFileSync(path.join(HERE, 'README.md'), 'utf8')); process.exit(0); }

const KEY = process.env.OPENROUTER_API_KEY;
const DRY = !!args['dry-run'];
if (!DRY && !KEY && !args['list-models']) die('Set OPENROUTER_API_KEY, or use --dry-run to test without API calls.');

if (args['list-models']) {
  const res = await fetch(`${API}/images/models`, { headers: KEY ? { Authorization: `Bearer ${KEY}` } : {} });
  const body = await res.json();
  for (const m of body.data || body.models || []) console.log(m.id || m.slug || JSON.stringify(m));
  process.exit(0);
}

const allModels = readJson('models.json');
const allPrompts = readJson('prompts.json');
const allEdits = readJson('edits.json');

const models = pick(allModels, args.models, m => m.key);
const prompts = pick(allPrompts, args.prompts, p => p.id);
const edits = args['skip-edits'] ? [] : allEdits.filter(e => prompts.some(p => p.id === e.source));
const maxCost = Number(args['max-cost'] ?? 15);
const concurrency = Math.max(1, Number(args.concurrency ?? 2));

const outDir = path.resolve(args.out || path.join(HERE, 'out', new Date().toISOString().slice(0, 10)));
fs.mkdirSync(outDir, { recursive: true });
fs.copyFileSync(path.join(HERE, 'report.html'), path.join(outDir, 'report.html'));

const resultsPath = path.join(outDir, 'results.json');
const state = fs.existsSync(resultsPath)
  ? JSON.parse(fs.readFileSync(resultsPath, 'utf8'))
  : { createdAt: new Date().toISOString(), models: [], prompts: [], edits: [], results: [] };
state.models = mergeBy(state.models, models, 'key');
state.prompts = mergeBy(state.prompts, prompts, 'id');
state.edits = mergeBy(state.edits, edits, 'id');
state.dryRun = DRY || state.dryRun === true;

const done = r => state.results.some(x => x.kind === r.kind && x.id === r.id && x.modelKey === r.modelKey && x.ok);
const genJobs = models.flatMap(m => prompts.map(p => ({ kind: 'map', id: p.id, modelKey: m.key, m, p })));
const editJobs = models.flatMap(m => edits.map(e => ({ kind: 'edit', id: e.id, modelKey: m.key, m, e })));
const todoGen = genJobs.filter(j => !done(j));
const todoEdit = editJobs.filter(j => !done(j));

const estimate = [...todoGen, ...todoEdit].reduce((s, j) => s + (j.m.estUsd || 0.05) * (j.kind === 'edit' ? 1.5 : 1), 0);
console.log(`${todoGen.length} maps and ${todoEdit.length} edits to run (${genJobs.length + editJobs.length - todoGen.length - todoEdit.length} already done).`);
console.log(`Estimated cost: about $${estimate.toFixed(2)}. Stops early if spending passes $${maxCost.toFixed(2)} (--max-cost).`);
if (DRY) console.log('Dry run: no API calls, placeholder images only.');
console.log(`Output: ${outDir}\n`);

let spent = state.results.reduce((s, r) => s + (r.costUsd || 0), 0);
let stopped = false;

await pool(todoGen, runMap);
await pool(todoEdit, runEdit);
save();

const ok = state.results.filter(r => r.ok).length;
const failed = state.results.filter(r => !r.ok).length;
console.log(`\nFinished. ${ok} images saved, ${failed} failed, $${spent.toFixed(3)} spent${DRY ? ' (dry run)' : ''}.`);
console.log(`Open ${path.join(outDir, 'report.html')} in a browser to score the results.`);

async function runMap(job) {
  const { m, p } = job;
  const prompt = buildPrompt(p);
  const aspect = pickAspect(p.squares);
  const file = path.join(m.key, `${p.id}`);
  await call(job, { model: m.model, prompt, aspect_ratio: aspect, resolution: m.resolution }, file, { prompt, aspect, squares: p.squares });
}

async function runEdit(job) {
  const { m, e } = job;
  const src = state.results.find(r => r.kind === 'map' && r.id === e.source && r.modelKey === m.key && r.ok);
  if (!src) return record(job, { ok: false, error: `No source map ${e.source} for this model.` });
  const bytes = fs.readFileSync(path.join(outDir, src.file));
  const p = state.prompts.find(x => x.id === e.source);
  const prompt = `This image is a tabletop RPG battle map of ${p.description}. Edit it: ${e.instruction} Keep the same perspective, art style, scale, and framing. Do not add grid lines or text.`;
  const ref = `data:${src.mediaType};base64,${bytes.toString('base64')}`;
  const file = path.join(m.key, `${e.id}`);
  await call(job, {
    model: m.model, prompt, aspect_ratio: src.aspect, resolution: m.resolution,
    input_references: [{ type: 'image_url', image_url: { url: ref } }],
  }, file, { prompt, aspect: src.aspect, squares: src.squares, source: e.source });
}

async function call(job, body, fileBase, extra) {
  if (stopped) return;
  if (spent >= maxCost) { stopped = true; console.log(`Stopping: spent $${spent.toFixed(2)}, limit $${maxCost}.`); return; }
  const t0 = Date.now();
  try {
    let data, note;
    if (DRY) {
      data = fakeResponse(extra.squares, job);
    } else {
      let res = await post(body);
      if (res.status === 400 && body.resolution) {
        note = 'Provider refused resolution 2K, retried with aspect ratio only.';
        const { resolution, ...rest } = body;
        res = await post(rest);
      }
      data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${data?.error?.message || JSON.stringify(data).slice(0, 300)}`);
    }
    const img = data.data?.[0];
    if (!img?.b64_json) throw new Error('Response had no image.');
    const bytes = Buffer.from(img.b64_json, 'base64');
    const mediaType = img.media_type || 'image/png';
    const file = `${fileBase}.${extFor(mediaType)}`;
    fs.mkdirSync(path.join(outDir, path.dirname(file)), { recursive: true });
    fs.writeFileSync(path.join(outDir, file), bytes);
    const size = imageSize(bytes) || { width: null, height: null };
    const costUsd = typeof data.usage?.cost === 'number' ? data.usage.cost : null;
    spent += costUsd || 0;
    record(job, {
      ok: true, file, mediaType, ...size, costUsd, ms: Date.now() - t0, note,
      pxPerSquare: size.width ? Math.round(size.width / extra.squares[0]) : null, ...extra,
    });
    console.log(`  ok   ${job.kind} ${job.id} on ${job.m.label}: ${size.width}x${size.height}, ${costUsd == null ? 'cost not reported' : '$' + costUsd.toFixed(4)}, ${((Date.now() - t0) / 1000).toFixed(1)}s`);
  } catch (err) {
    record(job, { ok: false, error: String(err.message || err), ms: Date.now() - t0, ...extra });
    console.log(`  FAIL ${job.kind} ${job.id} on ${job.m.label}: ${err.message || err}`);
  }
}

function post(body) {
  return fetch(`${API}/images`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${KEY}`, 'Content-Type': 'application/json', 'X-Title': 'Battle map bake-off' },
    body: JSON.stringify({ n: 1, ...body }),
  });
}

function record(job, fields) {
  state.results = state.results.filter(r => !(r.kind === job.kind && r.id === job.id && r.modelKey === job.modelKey));
  state.results.push({ kind: job.kind, id: job.id, modelKey: job.modelKey, model: job.m.model, estUsd: job.m.estUsd, ...fields });
  save();
}

function save() {
  state.updatedAt = new Date().toISOString();
  const json = JSON.stringify(state, null, 2);
  fs.writeFileSync(resultsPath, json);
  fs.writeFileSync(path.join(outDir, 'results.js'), `window.BAKEOFF = ${json};\n`);
}

async function pool(jobs, fn) {
  let i = 0;
  const workers = Array.from({ length: Math.min(concurrency, jobs.length) }, async () => {
    while (i < jobs.length && !stopped) await fn(jobs[i++]);
  });
  await Promise.all(workers);
}

/** Placeholder response for --dry-run: a flat-colour PNG at the size a 2K model would return. */
function fakeResponse([w, h], job) {
  const scale = 2048 / Math.max(w, h);
  const width = Math.round(w * scale), height = Math.round(h * scale);
  const shade = [...job.modelKey + job.id].reduce((s, c) => s + c.charCodeAt(0), 0) % 120 + 80;
  return { data: [{ b64_json: solidPng(width, height, [shade, shade - 30, shade - 60]).toString('base64'), media_type: 'image/png' }], usage: { cost: 0 } };
}

function solidPng(width, height, [r, g, b]) {
  const row = Buffer.alloc(1 + width * 3);
  for (let x = 0; x < width; x++) row.set([r, g, b], 1 + x * 3);
  const raw = Buffer.concat(Array.from({ length: height }, () => row));
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr.set([8, 2, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}

function crc32(buf) {
  let c = ~0;
  for (const byte of buf) { c ^= byte; for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1)); }
  return ~c >>> 0;
}

function parseArgs(list) {
  const out = {};
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    if (!a.startsWith('--')) continue;
    const [k, v] = a.slice(2).split('=');
    if (v !== undefined) out[k] = v;
    else if (list[i + 1] && !list[i + 1].startsWith('--')) out[k] = list[++i];
    else out[k] = true;
  }
  return out;
}

function pick(list, filter, keyOf) {
  if (!filter || filter === true) return list;
  const want = String(filter).split(',').map(s => s.trim());
  const found = list.filter(x => want.some(w => keyOf(x).startsWith(w)));
  if (!found.length) die(`Nothing matched "${filter}". Options: ${list.map(keyOf).join(', ')}`);
  return found;
}

function mergeBy(old, add, key) {
  const map = new Map((old || []).map(x => [x[key], x]));
  for (const x of add) map.set(x[key], x);
  return [...map.values()];
}

function readJson(name) {
  return JSON.parse(fs.readFileSync(path.join(HERE, name), 'utf8'));
}

function die(msg) {
  console.error(msg);
  process.exit(1);
}
