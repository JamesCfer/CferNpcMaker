import { describe, it, expect, beforeEach, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { planToChatCommands, fromBase64Url, toBase64Url, CHUNK_SIZE } from '../core/plan.js';
import { mapDnd5eNpc } from '../core/mappers/dnd5e-npc.js';
import { mapStore } from '../core/mappers/store.js';

const here = dirname(fileURLToPath(import.meta.url));
const source = readFileSync(join(here, '..', 'mod', 'CferImporter.js'), 'utf8');

/** A small fake of the Roll20 Mod sandbox. */
function sandbox({ gm = true } = {}) {
  let n = 0;
  const objs = [];
  const chat = [];
  const handlers = {};
  const make = (type, props) => {
    const o = { id: `${type}-${++n}`, type, props: { ...props }, worker: [],
      get(k) { return k === '_id' ? this.id : this.props[k]; },
      set(k, v) { if (typeof k === 'object') Object.assign(this.props, k); else this.props[k] = v; },
      setWithWorker(u) { this.worker.push(u); Object.assign(this.props, u); } };
    objs.push(o);
    return o;
  };
  const env = {
    state: {},
    on: (ev, fn) => { handlers[ev] = fn; },
    log: () => {},
    sendChat: (who, text) => chat.push(text),
    createObj: (type, props) => make(type, props),
    findObjs: q => objs.filter(o => (!q.type || o.type === q.type) && Object.entries(q).every(([k, v]) => k === 'type' || o.props[k] === v)),
    getObj: (type, id) => objs.find(o => o.type === type && o.id === id),
    playerIsGM: () => gm,
  };
  const api = new Function(...Object.keys(env), `${source}\nreturn CferImporter;`)(...Object.values(env));
  handlers.ready?.();
  const say = (content, selected) => handlers['chat:message']({ type: 'api', content, playerid: 'p1', selected });
  return { api, env, objs, chat, say, make };
}

const fixture = JSON.parse(readFileSync(join(here, 'fixtures', 'dnd5e-npc.json'), 'utf8'));

describe('plan transport', () => {
  it('round trips unicode through base64url chunks', () => {
    const plan = { v: 1, kind: 'handouts', summary: 'Café ✨ 龍', handouts: [{ name: 'x'.repeat(CHUNK_SIZE * 2) }] };
    const lines = planToChatCommands(plan, 'abc');
    expect(lines.length).toBeGreaterThan(1);
    for (const l of lines) expect(l).toMatch(/^!cfer-import abc \d+\/\d+ [A-Za-z0-9_-]+$/);
    expect(JSON.parse(fromBase64Url(lines.map(l => l.split(' ')[3]).join('')))).toEqual(plan);
  });

  it('the Mod decoder matches the extension encoder', () => {
    const { api } = sandbox();
    const text = 'Plain, Ünïcödé, emoji 🐉 and "quotes"';
    expect(api.decodeBase64Url(toBase64Url(text))).toBe(text);
  });
});

describe('Cfer Importer Mod script', () => {
  beforeEach(() => vi.useFakeTimers());

  it('answers the ping the extension uses to find it', () => {
    const s = sandbox();
    s.say('!cfer-ping');
    expect(s.chat.at(-1)).toContain('Cfer Importer 1.0.0 is installed.');
  });

  it('ignores players who are not the GM', () => {
    const s = sandbox({ gm: false });
    s.say('!cfer-ping');
    for (const l of planToChatCommands({ v: 1, kind: 'handouts', handouts: [{ name: 'Secret' }] }, 'x')) s.say(l);
    expect(s.chat).toEqual([]);
    expect(s.objs.some(o => o.type === 'handout')).toBe(false);
  });

  it('imports a 5e NPC through the sheet drop, chunks in any order', () => {
    const s = sandbox();
    const plan = mapDnd5eNpc(fixture);
    const lines = planToChatCommands(plan, 'id1');
    for (const l of [...lines].reverse()) s.say(l);
    const char = s.objs.find(o => o.type === 'character');
    expect(char.props.name).toBe(plan.character.name);
    const trigger = s.objs.find(o => o.type === 'attribute' && o.props.name === 'drop_category');
    expect(trigger.worker[0].current).toBe('Monsters');
    // Pretend the sheet worker built the NPC.
    s.env.createObj('attribute', { characterid: char.id, name: 'npc_name', current: plan.character.name });
    vi.advanceTimersByTime(plan.drop.timeoutMs + 10);
    expect(s.chat.at(-1)).toContain('Created');
    expect(s.chat.at(-1)).not.toContain('written directly');
  });

  it('writes the fallback when the sheet drop does not run', () => {
    const s = sandbox();
    const plan = mapDnd5eNpc(fixture);
    for (const l of planToChatCommands(plan, 'id2')) s.say(l);
    vi.advanceTimersByTime(plan.drop.timeoutMs + 10);
    expect(s.chat.at(-1)).toContain('written directly');
    const char = s.objs.find(o => o.type === 'character');
    const npcName = s.objs.find(o => o.type === 'attribute' && o.props.characterid === char.id && o.props.name === 'npc_name');
    expect(npcName.props.current).toBe(plan.character.name);
  });

  it('creates store handouts and gives an item to the selected token', () => {
    const s = sandbox();
    const plan = mapStore('dnd5e')({ store: { name: 'Shop', inventory: [{ name: 'Rope', type: 'loot', system: {} }] } });
    for (const l of planToChatCommands(plan, 'id3')) s.say(l);
    expect(s.objs.filter(o => o.type === 'handout').map(o => o.props.name)).toEqual(['Shop', 'Shop: Rope']);
    const menu = s.chat.at(-1);
    const key = menu.match(/!cfer-give (\S+)\)/)[1];
    const pc = s.env.createObj('character', { name: 'Hero' });
    const token = s.env.createObj('graphic', { represents: pc.id });
    s.say(`!cfer-give ${key}`, [{ _id: token.id, _type: 'graphic' }]);
    expect(s.chat.at(-1)).toContain('Gave Rope to 1 character');
    expect(s.objs.some(o => o.type === 'attribute' && o.props.characterid === pc.id && o.props.name === 'drop_category')).toBe(true);
  });

  it('sizes the page and adds walls, doors and lights to the selected map', () => {
    const s = sandbox();
    const page = s.env.createObj('page', {});
    const img = s.env.createObj('graphic', { pageid: page.id, imgsrc: 'https://s3.amazonaws.com/files.d20.io/images/1/thumb.png' });
    const plan = { v: 1, kind: 'map', summary: 'Crypt', map: { name: 'Crypt', squares: [10, 8], grid: { distance: 5, units: 'ft' }, darkness: true,
      walls: [{ kind: 'wall', c: [0, 0, 700, 0] }, { kind: 'door', c: [350, 0, 350, 70] }],
      lights: [{ kind: 'torch', x: 100, y: 100, bright: 20, dim: 40, color: '#ff9900' }] } };
    for (const l of planToChatCommands(plan, 'id4')) s.say(l, [{ _id: img.id, _type: 'graphic' }]);
    expect(page.props).toMatchObject({ width: 10, height: 8, scale_number: 5, dynamic_lighting_enabled: true });
    expect(img.props).toMatchObject({ layer: 'map', width: 700, height: 560 });
    const path = s.objs.find(o => o.type === 'path');
    expect(path.props.layer).toBe('walls');
    const door = s.objs.find(o => o.type === 'door');
    expect(door.props.y).toBe(-35);
    expect(door.props.path.handle0).toEqual({ x: 0, y: 35 });
    const light = s.objs.find(o => o.type === 'graphic' && o.props.name === 'Light: torch');
    expect(light.props).toMatchObject({ layer: 'walls', bright_light_distance: 20, low_light_distance: 40 });
    expect(s.chat.at(-1)).toContain('1 walls, 1 doors');
  });

  it('asks for a selected map image first', () => {
    const s = sandbox();
    for (const l of planToChatCommands({ v: 1, kind: 'map', map: { name: 'X', squares: [5, 5], grid: { distance: 5, units: 'ft' }, walls: [], lights: [] } }, 'id5')) s.say(l);
    expect(s.chat.at(-1)).toContain('Select the map image');
  });
});
