/**
 * The Cfer panel on the Roll20 editor page (isolated world, ES module).
 *
 * It shows the generators, sends requests through the background worker (which
 * holds the Patreon key), turns answers into import plans with the shared
 * mappers, and writes them into the game one of three ways:
 *   1. "mod":  chat commands to the companion Cfer Importer Mod script (Pro games)
 *   2. "page": Roll20's in-page campaign objects through page-bridge.js
 *   3. "file": a JSON or image download the GM imports by hand
 */
import { GENERATORS, getGenerator, endpointFor, readForm, ENDPOINTS, MAP_COSTS } from './core/catalog.js';
import { planToChatCommands } from './core/plan.js';
import { buildMapRequest, mapFinal } from './core/mappers/map.js';

const api = globalThis.browser ?? globalThis.chrome;
const MAX_RESULTS = 15;
const MAX_MOCKUPS = 8;

/* ── messaging ───────────────────────────────────────── */

const bg = msg => new Promise((resolve, reject) => {
  try { api.runtime.sendMessage(msg, res => { const e = api.runtime.lastError; e ? reject(new Error(e.message)) : resolve(res); }); }
  catch (err) { reject(new Error('The extension was updated. Reload this Roll20 tab.')); }
});

let seq = 0;
const waiting = new Map();
window.addEventListener('message', ev => {
  const m = ev.data;
  if (ev.source !== window || !m || m.cfer !== 'res' || !waiting.has(m.id)) return;
  const { resolve, reject, timer } = waiting.get(m.id);
  waiting.delete(m.id); clearTimeout(timer);
  m.ok ? resolve(m.result) : reject(new Error(m.error));
});
function bridge(op, args = {}, timeoutMs = 60000) {
  return new Promise((resolve, reject) => {
    const id = `c${Date.now()}-${++seq}`;
    const timer = setTimeout(() => { waiting.delete(id); reject(new Error('Roll20 did not answer. Reload the tab and try again.')); }, timeoutMs);
    waiting.set(id, { resolve, reject, timer });
    window.postMessage({ cfer: 'req', id, op, args }, '*');
  });
}

/* ── tiny DOM helper ─────────────────────────────────── */

function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k in el && typeof v !== 'string') el[k] = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat()) if (kid !== null && kid !== undefined && kid !== false) el.append(kid.nodeType ? kid : String(kid));
  return el;
}
const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;
const dataUrl = img => `data:${img.mediaType || 'image/png'};base64,${img.b64}`;

/* ── state ───────────────────────────────────────────── */

const state = {
  signedIn: false, usage: null, busy: false,
  writer: null,          // detected: { mod, page }
  method: 'auto',        // auto | mod | page | file
  results: [],           // [{ at, plan }]
  map: { mockups: [], selected: -1, final: null, plan: null },
};

let root, els = {};

/* ── writer detection ────────────────────────────────── */

async function detectWriter(force = false) {
  if (state.writer && !force) return state.writer;
  let probe = {};
  try { probe = await bridge('probe', {}, 5000); } catch (_) { /* bridge missing */ }
  let mod = false;
  if (probe.chat) {
    try {
      await bridge('sendChat', { lines: ['!cfer-ping'] }, 5000);
      mod = await bridge('waitForChat', { text: 'Cfer Importer', timeoutMs: 4000 }, 8000);
    } catch (_) { mod = false; }
  }
  state.writer = { mod, page: !!(probe.characters && probe.handouts), chat: !!probe.chat, probe };
  renderWriter();
  return state.writer;
}

async function chooseMethod(kind) {
  if (state.method === 'file') return 'file';
  const w = await detectWriter();
  if (state.method === 'mod') { if (!w.chat) throw new Error('Could not find the Roll20 chat box.'); return 'mod'; }
  if (state.method === 'page') { if (!w.page) throw new Error('This Roll20 page does not allow in-page imports. Try the Mod script or a download.'); return 'page'; }
  if (kind === 'map') return w.mod ? 'mod' : 'file';
  // In-page writes handle both the 2014 and the 2024 D&D sheets; the Mod script only the 2014 one.
  if (w.page) return 'page';
  if (w.mod) return 'mod';
  return 'file';
}

/* ── import ──────────────────────────────────────────── */

/** The Mod script needs only the fields it applies; big extras stay here. */
function forMod(plan) {
  const { hdc, heroImport, ...rest } = plan;
  return rest;
}

function download(text, filename, type = 'application/json') {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = h('a', { href: url, download: filename });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
const safeName = s => String(s || 'cfer').replace(/[^\w-]+/g, '-').replace(/^-|-$/g, '').toLowerCase() || 'cfer';

async function importPlan(plan, card) {
  const method = await chooseMethod(plan.kind);
  const note = card.querySelector('.cfer-card-note');
  if (method === 'file') {
    download(JSON.stringify(plan, null, 2), `${safeName(plan.character?.name || plan.handouts?.[0]?.name || plan.summary)}.cfer.json`);
    note.textContent = state.writer?.chat
      ? 'Downloaded the import file. Install the free Cfer Importer Mod script (Pro games) to import in one click.'
      : 'Downloaded the import file.';
    return;
  }
  if (method === 'mod') {
    const lines = planToChatCommands(forMod(plan));
    note.textContent = `Sending to the Cfer Importer script (${plural(lines.length, 'message', 'messages')})…`;
    await bridge('sendChat', { lines }, 30000 + lines.length * 500);
    note.textContent = 'Sent. The Cfer Importer script whispers you a link when it is done.';
    return;
  }
  note.textContent = 'Importing…';
  if (plan.kind === 'character') {
    const res = await bridge('createCharacter', { plan: forMod(plan) }, 60000);
    note.textContent = `Created ${res.name}${res.mode === 'sheet2024' ? ' on the D&D 2024 sheet' : res.mode === 'direct' && plan.drop ? ' (written directly; the sheet\'s own import did not run)' : ''}. Find it in the Journal.`;
  } else if (plan.kind === 'handouts') {
    const res = await bridge('createHandouts', { plan }, 60000);
    note.textContent = `Created ${plural(res.handouts.length, 'handout', 'handouts')} in the Journal.`;
    const givable = plan.handouts.filter(x => x.give);
    if (givable.length) card.append(await giveBox(givable));
  }
}

async function giveBox(handouts) {
  let chars = [];
  try { chars = await bridge('listCharacters'); } catch (_) { /* none */ }
  const pick = h('select', { class: 'cfer-input' }, chars.map(c => h('option', { value: c.id }, c.name || '(unnamed)')));
  const box = h('div', { class: 'cfer-give' }, h('div', { class: 'cfer-label' }, 'Give to'), pick);
  for (const ho of handouts) {
    const status = h('span', { class: 'cfer-muted' });
    box.append(h('div', { class: 'cfer-give-row' }, h('span', {}, ho.give.name || ho.name),
      h('button', { class: 'cfer-btn small', onclick: async ev => {
        ev.target.disabled = true;
        try { const r = await bridge('give', { recipe: ho.give, characterId: pick.value }, 30000); status.textContent = `Given to ${r.name}.`; }
        catch (err) { status.textContent = err.message; }
        finally { ev.target.disabled = false; }
      } }, 'Give'), status));
  }
  if (!chars.length) box.append(h('div', { class: 'cfer-muted' }, 'No characters found in this game yet.'));
  return box;
}

/* ── result cards ────────────────────────────────────── */

function resultCard(plan) {
  const card = h('div', { class: 'cfer-card' },
    h('div', { class: 'cfer-card-title' }, plan.summary || plan.generator),
    h('div', { class: 'cfer-card-note cfer-muted' }));
  const run = fn => async ev => {
    const btn = ev.currentTarget; btn.disabled = true;
    try { await fn(); } catch (err) { card.querySelector('.cfer-card-note').textContent = err.message; }
    finally { btn.disabled = false; }
  };
  const buttons = h('div', { class: 'cfer-row' },
    h('button', { class: 'cfer-btn primary', onclick: run(() => importPlan(plan, card)) }, 'Import into game'),
    h('button', { class: 'cfer-btn', onclick: () => download(JSON.stringify(plan, null, 2), `${safeName(plan.summary)}.cfer.json`) }, 'Download'));
  if (plan.hdc) buttons.append(h('button', { class: 'cfer-btn', onclick: () => download(plan.hdc.xml, plan.hdc.filename, 'application/xml') }, '.hdc file'));
  if (plan.heroImport) buttons.append(h('button', { class: 'cfer-btn', title: 'Needs the HeroSystem6eHeroic HDImporter Mod script', onclick: run(async () => {
    await bridge('sendChat', { lines: [plan.heroImport] }, 30000);
    card.querySelector('.cfer-card-note').textContent = 'Sent to HD Importer. It makes a full HeroSystem6eHeroic sheet.';
  }) }, 'Send to HD Importer'));
  card.append(buttons);
  return card;
}

async function saveResult(plan) {
  state.results.unshift({ at: Date.now(), plan });
  state.results = state.results.slice(0, MAX_RESULTS);
  try { await api.storage.local.set({ cferResults: state.results }); } catch (_) { /* storage full: keep in memory */ }
  renderResults();
}

/* ── generate ────────────────────────────────────────── */

async function call(url, payload, timeoutMs) {
  const res = await bg({ type: 'generate', url, payload, timeoutMs });
  if (res?.usage) setUsage(res.usage);
  if (!res?.ok) {
    if (res?.status === 401 || res?.status === 403) { state.signedIn = false; renderAccount(); }
    throw new Error(res?.error || 'The request failed.');
  }
  return res.body ?? res.text;
}

async function generate() {
  const gen = getGenerator(els.genPick.value);
  const values = Object.fromEntries(new FormData(els.genForm).entries());
  for (const f of gen.fields) if (f.type === 'checkbox') values[f.name] = els.genForm.elements[f.name].checked;
  const form = readForm(gen, values);
  await busy(`Generating ${gen.label}… this can take a minute or two.`, async () => {
    const data = await call(endpointFor(gen, form), gen.payload(form), 300000);
    if (!data || typeof data !== 'object' || !Object.keys(data).length) throw new Error('The server sent back an empty answer. Please try again.');
    const plan = gen.map(data, form);
    await saveResult(plan);
    els.genOut.replaceChildren(resultCard(plan));
    say('Ready. Press Import into game.');
  });
}

async function busy(message, fn) {
  if (state.busy) return;
  state.busy = true; say(message); renderAccount();
  const started = Date.now();
  const clock = setInterval(() => say(`${message} ${Math.floor((Date.now() - started) / 1000)}s`), 1000);
  try { await fn(); }
  catch (err) { say(err.message, true); }
  finally { clearInterval(clock); state.busy = false; renderAccount(); renderMapButtons(); }
}

function say(text, error = false) {
  els.status.textContent = text || '';
  els.status.classList.toggle('error', !!error);
}

/* ── battle map ──────────────────────────────────────── */

function mapValues() {
  const f = els.mapForm.elements;
  const v = Object.fromEntries(new FormData(els.mapForm).entries());
  v.walls = f.walls.checked; v.autoPlan = f.autoPlan.checked;
  return v;
}

function mapCost() {
  const v = mapValues();
  const mock = MAP_COSTS.mockups[Number(v.mockups)] ?? 3;
  const fin = (MAP_COSTS.final[v.quality] ?? 6) + (v.walls && v.perspective === 'top-down' ? MAP_COSTS.walls : 0);
  return { mock, fin };
}

function renderMapButtons() {
  if (!els.mapMock) return;
  const { mock, fin } = mapCost();
  els.mapMock.textContent = `${state.map.mockups.length ? 'More mockups' : 'Generate mockups'} (${plural(mock, 'use', 'uses')})`;
  els.mapFinal.textContent = `Make final map (${plural(fin, 'use', 'uses')})`;
  els.mapMock.disabled = state.busy || !state.signedIn;
  els.mapFinal.disabled = state.busy || !state.signedIn || state.map.selected < 0;
}

function renderMockups() {
  els.mapThumbs.replaceChildren(...state.map.mockups.map((m, i) => h('img', {
    class: `cfer-thumb${i === state.map.selected ? ' selected' : ''}`, src: dataUrl(m), alt: `Mockup ${i + 1}`,
    onclick: () => { state.map.selected = i; renderMockups(); renderMapButtons(); },
  })));
}

async function makeMockups() {
  const v = mapValues();
  const map = buildMapRequest(v);
  const count = Number(v.mockups) || 4;
  await busy(`Drawing ${plural(count, 'mockup', 'mockups')}…`, async () => {
    const data = await call(ENDPOINTS.mapBuilder, { action: 'mockups', map, count, ...(!map.layout && v.autoPlan ? { autoPlan: true } : {}) }, 300000);
    if (typeof data?.layout === 'string' && data.layout.trim()) els.mapForm.elements.layout.value = data.layout.trim();
    if (!Array.isArray(data?.images) || !data.images.length) throw new Error('No mockups were returned.');
    const made = { ...map, ...(data.layout ? { layout: String(data.layout).trim().slice(0, 1200) } : {}) };
    state.map.mockups = [...data.images.map(img => ({ b64: img.b64, mediaType: img.mediaType, map: made })), ...state.map.mockups].slice(0, MAX_MOCKUPS);
    state.map.selected = -1;
    renderMockups();
    say(data.failed ? `${data.images.length} mockups ready, ${data.failed} failed. Pick one.` : 'Pick a mockup, then make the final map.');
  });
}

async function makeFinal() {
  const mock = state.map.mockups[state.map.selected];
  if (!mock) return say('Pick a mockup first.', true);
  const v = mapValues();
  await busy('Making the final map. This can take several minutes.', async () => {
    const wantWalls = v.walls && mock.map.perspective === 'top-down';
    const data = await call(ENDPOINTS.mapBuilder, { action: 'final', quality: v.quality, map: mock.map,
      mockup: { b64: mock.b64, mediaType: mock.mediaType }, units: v.units, walls: wantWalls }, 330000);
    const name = (v.name || '').trim() || 'Battle map';
    state.map.final = data;
    state.map.plan = mapFinal(data, { name, walls: wantWalls });
    renderFinal();
    say(data.warning || 'Final map ready.');
  });
}

function renderFinal() {
  const { final, plan } = state.map;
  if (!final || !plan) { els.mapResult.replaceChildren(); return; }
  const m = plan.map;
  const note = h('div', { class: 'cfer-muted' });
  const doors = m.walls.filter(w => w.kind === 'door').length;
  const run = fn => async ev => { const b = ev.currentTarget; b.disabled = true; try { await fn(); } catch (err) { note.textContent = err.message; } finally { b.disabled = false; } };
  els.mapResult.replaceChildren(
    h('img', { class: 'cfer-final', src: dataUrl(final.image), alt: m.name }),
    h('div', {}, `${m.squares[0]} x ${m.squares[1]} squares, ${plural(m.walls.length - doors, 'wall', 'walls')}, ${plural(doors, 'door', 'doors')}, ${plural(m.lights.length, 'light', 'lights')}`),
    h('ol', { class: 'cfer-steps' },
      h('li', {}, 'Place the image on the map layer (or drag the downloaded file onto the page).'),
      h('li', {}, 'Select the image, then press Add walls and lights. This needs the Cfer Importer Mod script.')),
    h('div', { class: 'cfer-row' },
      h('button', { class: 'cfer-btn primary', onclick: run(async () => {
        try {
          note.textContent = 'Uploading the image to Roll20…';
          const res = await bridge('dropImage', { dataUrl: dataUrl(final.image), filename: m.image.filename, squares: m.squares }, 60000);
          note.textContent = res.placed ? `Placed on the map layer at ${m.squares[0]} x ${m.squares[1]} squares.` : 'Dropped on the page. If nothing appeared, use Download image and drag it in.';
        }
        catch (err) { download64(final.image, m.image.filename); note.textContent = `${err.message} Downloaded the image instead; drag it onto the page.`; }
      }) }, 'Place image'),
      h('button', { class: 'cfer-btn', onclick: () => download64(final.image, m.image.filename) }, 'Download image'),
      h('button', { class: 'cfer-btn primary', onclick: run(async () => {
        const w = await detectWriter();
        if (!w.mod) { download(JSON.stringify(plan, null, 2), `${safeName(m.name)}.cfer.json`); note.textContent = 'Walls and lights need the free Cfer Importer Mod script (Pro games). Downloaded the map data instead.'; return; }
        await bridge('sendChat', { lines: planToChatCommands(plan) }, 60000);
        note.textContent = 'Sent. The script sizes the page and adds walls, doors and lights to the selected image.';
      }) }, 'Add walls and lights')),
    note);
}

function download64(img, filename) {
  const bin = atob(img.b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const url = URL.createObjectURL(new Blob([bytes], { type: img.mediaType || 'image/png' }));
  const a = h('a', { href: url, download: filename }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

/* ── forms ───────────────────────────────────────────── */

function field(f, prefix = 'gen') {
  const id = `cfer-${prefix}-${f.name}`;
  let input;
  if (f.type === 'select') input = h('select', { name: f.name, id, class: 'cfer-input' }, f.options.map(o => h('option', { value: o, selected: o === f.value }, o === '' ? '(any)' : o)));
  else if (f.type === 'textarea') input = h('textarea', { name: f.name, id, class: 'cfer-input', rows: 4, placeholder: f.placeholder, required: f.required });
  else if (f.type === 'checkbox') return h('label', { class: 'cfer-check' }, h('input', { type: 'checkbox', name: f.name, checked: !!f.value }), f.label);
  else input = h('input', { type: f.type, name: f.name, id, class: 'cfer-input', placeholder: f.placeholder, min: f.min, max: f.max, value: f.value ?? '' });
  return h('div', { class: 'cfer-field' }, h('label', { for: id, class: 'cfer-label' }, f.label), input);
}

function renderGenForm() {
  const gen = getGenerator(els.genPick.value);
  els.genForm.replaceChildren(...gen.fields.map(f => field(f)),
    h('div', { class: 'cfer-muted' }, `Sheet: ${gen.sheet}`),
    h('button', { type: 'submit', class: 'cfer-btn primary wide', disabled: !state.signedIn || state.busy }, `Generate (${plural(gen.cost, 'use', 'uses')})`));
}

const MAP_FIELDS = [
  { name: 'name', label: 'Map name', type: 'text', placeholder: 'Battle map' },
  { name: 'description', label: 'Describe the map', type: 'textarea', required: true, placeholder: 'A ruined chapel in a misty graveyard, with a crypt entrance…' },
  { name: 'perspective', label: 'Perspective', type: 'select', options: ['top-down', 'isometric', 'three-quarter'], value: 'top-down' },
  { name: 'genre', label: 'Genre', type: 'select', options: ['fantasy', 'sci-fi', 'modern', 'post-apocalyptic', 'horror', 'western', 'steampunk'], value: 'fantasy' },
  { name: 'style', label: 'Style', type: 'select', options: ['painted', 'realistic', 'hand-drawn', 'cartoon', 'pixel art'], value: 'painted' },
  { name: 'mood', label: 'Time of day', type: 'select', options: ['day', 'dusk', 'night'], value: 'day' },
  { name: 'width', label: 'Width (squares)', type: 'number', min: 5, max: 60, value: 24 },
  { name: 'height', label: 'Height (squares)', type: 'number', min: 5, max: 60, value: 18 },
  { name: 'detail', label: 'Detail (1 to 5)', type: 'number', min: 1, max: 5, value: 3 },
  { name: 'clutter', label: 'Props (1 to 5)', type: 'number', min: 1, max: 5, value: 3 },
  { name: 'wear', label: 'Wear (1 to 5)', type: 'number', min: 1, max: 5, value: 3 },
  { name: 'mockups', label: 'Mockups', type: 'select', options: ['1', '2', '4'], value: '4' },
  { name: 'quality', label: 'Final quality', type: 'select', options: ['standard', 'economy', 'draft', 'high'], value: 'standard' },
  { name: 'units', label: 'Units', type: 'select', options: ['imperial', 'metric'], value: 'imperial' },
  { name: 'autoPlan', label: 'Plan the layout first', type: 'checkbox', value: true },
  { name: 'walls', label: 'Walls, doors and lights (+1 use, top-down only)', type: 'checkbox', value: true },
];

/* ── account ─────────────────────────────────────────── */

function setUsage(u) { if (u && (u.remaining !== null || u.limit !== null)) { state.usage = u; renderAccount(); } }

function renderAccount() {
  if (!els.account) return;
  const u = state.usage;
  els.pill.textContent = state.signedIn ? (u?.remaining !== null && u?.remaining !== undefined ? `${u.remaining}${u.limit ? ` / ${u.limit}` : ''} uses` : 'Signed in') : 'Signed out';
  els.account.replaceChildren(state.signedIn
    ? h('div', { class: 'cfer-row' },
      h('span', { class: 'cfer-muted' }, u?.tier ? `Patreon: ${u.tier}` : 'Signed in with Patreon'),
      h('button', { class: 'cfer-btn small', onclick: async () => { await bg({ type: 'signOut' }); state.signedIn = false; state.usage = null; renderAccount(); } }, 'Sign out'))
    : h('div', { class: 'cfer-signin' },
      h('p', {}, 'Sign in with Patreon to generate. Your uses are shared with the Foundry modules.'),
      h('div', { class: 'cfer-row' },
        h('button', { class: 'cfer-btn primary', onclick: signIn }, 'Sign in with Patreon'),
        h('button', { class: 'cfer-btn', onclick: () => bg({ type: 'openPatreon' }) }, 'Join on Patreon'))));
  if (els.genForm) {
    const submit = els.genForm.querySelector('button[type=submit]');
    if (submit) submit.disabled = !state.signedIn || state.busy;
  }
  renderMapButtons();
}

async function signIn() {
  say('A Patreon tab opened. Finish signing in there.');
  const res = await bg({ type: 'signIn' });
  if (!res?.ok) return say(res?.error || 'Sign-in failed.', true);
  await refreshStatus();
  say('Signed in.');
}

async function refreshStatus() {
  const s = await bg({ type: 'status' });
  state.signedIn = !!s?.signedIn;
  state.usage = s?.usage || null;
  renderAccount();
  // Fresh counts when the panel opens, at most once a minute.
  if (state.signedIn && (!state.usage?.at || Date.now() - state.usage.at > 60000)) {
    const res = await bg({ type: 'usage' }).catch(() => null);
    if (res?.usage) setUsage({ ...res.usage, at: Date.now() });
    if (res && (res.status === 401 || res.status === 403)) { state.signedIn = false; renderAccount(); }
  }
}

function renderWriter() {
  if (!els.writer) return;
  const w = state.writer;
  els.writer.textContent = !w ? 'Import method not checked yet.'
    : w.mod ? 'Cfer Importer Mod script found: imports go through it.'
      : w.page ? 'Imports are written straight into this game.'
        : 'Imports download as files (no in-game import found).';
}

function renderResults() {
  if (!els.results) return;
  els.results.replaceChildren(...(state.results.length
    ? state.results.map(r => resultCard(r.plan))
    : [h('p', { class: 'cfer-muted' }, 'Nothing generated yet.')]));
}

/* ── shell ───────────────────────────────────────────── */

function tab(name) {
  for (const [k, el] of Object.entries(els.tabs)) el.hidden = k !== name;
  for (const b of els.tabBar.children) b.classList.toggle('active', b.dataset.tab === name);
  if (name === 'results') renderResults();
}

function build() {
  const host = h('div', { id: 'cfer-roll20-host' });
  // Roll20's keyboard shortcuts listen on the document; keep typing in the panel to the panel.
  for (const type of ['keydown', 'keyup', 'keypress']) host.addEventListener(type, ev => ev.stopPropagation());
  root = host.attachShadow({ mode: 'open' });
  root.append(h('link', { rel: 'stylesheet', href: api.runtime.getURL('panel.css') }));

  const launcher = h('button', { class: 'cfer-launcher', title: 'Cfer NPC Maker', 'aria-label': 'Cfer NPC Maker', onclick: () => { panel.hidden = !panel.hidden; if (!panel.hidden) refreshStatus(); } },
    h('img', { src: api.runtime.getURL('icons/icon128.png'), alt: '' }));
  els.pill = h('span', { class: 'cfer-pill' });
  els.status = h('div', { class: 'cfer-status', role: 'status' });
  els.account = h('div', { class: 'cfer-account' });

  els.genPick = h('select', { class: 'cfer-input', onchange: renderGenForm }, GENERATORS.filter(g => g.available !== false).map(g => h('option', { value: g.id }, g.label)));
  els.genForm = h('form', { class: 'cfer-form', onsubmit: ev => { ev.preventDefault(); generate().catch(err => say(err.message, true)); } });
  els.genOut = h('div');

  els.mapForm = h('form', { class: 'cfer-form', onsubmit: ev => ev.preventDefault(), oninput: renderMapButtons }, MAP_FIELDS.map(f => field(f, 'map')),
    h('details', {}, h('summary', {}, 'Layout plan (optional)'), h('textarea', { name: 'layout', class: 'cfer-input', rows: 4, maxlength: 1200, placeholder: 'Filled in for you when "Plan the layout first" is on. Edit it to steer the drawing.' })));
  els.mapMock = h('button', { class: 'cfer-btn primary', onclick: () => makeMockups().catch(err => say(err.message, true)) });
  els.mapFinal = h('button', { class: 'cfer-btn primary', onclick: () => makeFinal().catch(err => say(err.message, true)) });
  els.mapThumbs = h('div', { class: 'cfer-thumbs' });
  els.mapResult = h('div', { class: 'cfer-map-result' });

  els.writer = h('div', { class: 'cfer-muted' });
  const method = h('select', { class: 'cfer-input', onchange: async ev => { state.method = ev.target.value; await api.storage.local.set({ cferMethod: state.method }); } },
    [['auto', 'Automatic'], ['mod', 'Cfer Importer Mod script'], ['page', 'Write into this game directly'], ['file', 'Download files']].map(([v, l]) => h('option', { value: v, selected: v === state.method }, l)));

  els.results = h('div');
  els.tabs = {
    create: h('section', {}, h('div', { class: 'cfer-field' }, h('label', { class: 'cfer-label' }, 'Generator'), els.genPick), els.genForm, els.genOut),
    map: h('section', { hidden: true }, els.mapForm, h('div', { class: 'cfer-row' }, els.mapMock, els.mapFinal), els.mapThumbs, els.mapResult),
    results: h('section', { hidden: true }, els.results),
    settings: h('section', { hidden: true },
      h('div', { class: 'cfer-field' }, h('label', { class: 'cfer-label' }, 'Import method'), method), els.writer,
      h('button', { class: 'cfer-btn', onclick: () => detectWriter(true) }, 'Check again'),
      h('p', { class: 'cfer-muted' }, 'Pro games: install the free Cfer Importer Mod script for the most reliable imports and for battle map walls and lights. Other games import directly from this panel or by file.')),
  };
  els.tabBar = h('nav', { class: 'cfer-tabs' }, [['create', 'Create'], ['map', 'Battle map'], ['results', 'Results'], ['settings', 'Settings']]
    .map(([k, l]) => h('button', { 'data-tab': k, class: k === 'create' ? 'active' : '', onclick: () => tab(k) }, l)));

  const panel = h('div', { class: 'cfer-panel', hidden: true },
    h('header', {}, h('strong', {}, 'Cfer NPC Maker'), els.pill, h('button', { class: 'cfer-close', title: 'Close', onclick: () => { panel.hidden = true; } }, '×')),
    els.account, els.tabBar, h('div', { class: 'cfer-body' }, ...Object.values(els.tabs)), els.status);

  root.append(launcher, panel);
  document.body.append(host);
  renderGenForm(); renderAccount(); renderWriter(); renderMapButtons();
}

export async function start() {
  try {
    const saved = await api.storage.local.get(['cferResults', 'cferMethod']);
    state.results = Array.isArray(saved.cferResults) ? saved.cferResults : [];
    state.method = saved.cferMethod || 'auto';
  } catch (_) { /* fresh install */ }
  build();
  await refreshStatus().catch(() => {});
  api.storage.onChanged?.addListener((changes, area) => {
    if (area === 'local' && changes.cferUsage?.newValue) setUsage(changes.cferUsage.newValue);
  });
}
