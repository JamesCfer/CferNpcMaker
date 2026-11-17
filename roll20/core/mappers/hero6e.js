/**
 * Hero System 6e: the generator returns a HERO Designer .hdc file (XML).
 *
 * Roll20's HeroSystem6eHeroic sheet has a companion Mod, "HeroSystem6eHeroic
 * HDImporter", that imports a character from `!hero --import {json}` where the
 * JSON is what HERO Designer's HeroSystem6eHeroic.hde export template writes.
 * We convert the .hdc into that JSON. Characteristics are exact; power, skill
 * and complication text is rebuilt from the .hdc fields and is close to, not
 * identical to, HERO Designer's own wording.
 *
 * Without that Mod (or without Pro) the plan still writes characteristics and
 * a full text summary straight onto the sheet.
 */
import { makePlan } from '../plan.js';
import { escapeHtml, num } from '../util.js';

/** 6e base values (HERO Designer's Heroic6E template). */
export const BASES = { STR: 10, DEX: 10, CON: 10, INT: 10, EGO: 10, PRE: 10, OCV: 3, DCV: 3, OMCV: 3, DMCV: 3, SPD: 2,
  PD: 2, ED: 2, REC: 4, END: 20, BODY: 10, STUN: 20, RUNNING: 12, LEAPING: 4, SWIMMING: 4 };
const KEYS = { STR: 'strength', DEX: 'dexterity', CON: 'constitution', INT: 'intelligence', EGO: 'ego', PRE: 'presence', OCV: 'ocv',
  DCV: 'dcv', OMCV: 'omcv', DMCV: 'dmcv', SPD: 'speed', PD: 'pd', ED: 'ed', REC: 'recovery', END: 'endurance', BODY: 'body',
  STUN: 'stun', RUNNING: 'running', LEAPING: 'leaping', SWIMMING: 'swimming' };
const ATTACKS = { RKA: 'killing', HKA: 'killing', ENERGYBLAST: 'normal', BLAST: 'normal', HANDTOHANDATTACK: 'normal', EGOATTACK: 'normal' };

const attr = (el, name) => el?.getAttribute?.(name) ?? '';
const kids = (el, tag) => (el ? [...el.children].filter(c => !tag || c.tagName === tag) : []);
const frac = v => {
  const n = Math.abs(v), whole = Math.floor(n), rest = n - whole;
  const f = { 0.25: '1/4', 0.5: '1/2', 0.75: '3/4' }[Math.round(rest * 4) / 4] || '';
  return `${v < 0 ? '-' : '+'}${whole ? whole : ''}${whole && f ? ' ' : ''}${f || (!whole ? '0' : '')}`;
};

/** "Alias: Input (option)" plus selected adders, the way HERO Designer starts its descriptions. */
function describe(el) {
  const name = attr(el, 'ALIAS') || attr(el, 'XMLID');
  const input = attr(el, 'INPUT');
  const option = attr(el, 'OPTION_ALIAS');
  const levels = num(attr(el, 'LEVELS'), 0);
  const adders = kids(el, 'ADDER').filter(a => attr(a, 'DISPLAYINSTRING') !== 'No').map(a => attr(a, 'OPTION_ALIAS') || attr(a, 'ALIAS')).filter(Boolean);
  let head = name;
  const xml = attr(el, 'XMLID');
  if (ATTACKS[xml]) head += ` ${levels}d6`;
  else if (attr(el, 'PDLEVELS') || attr(el, 'EDLEVELS')) head += ` (${num(attr(el, 'PDLEVELS'))} PD/${num(attr(el, 'EDLEVELS'))} ED)`;
  else if (levels && !['SKILL', 'PERK', 'TALENT', 'DISAD'].includes(el.tagName)) head += ` ${levels}`;
  if (input) head += `: ${input}`;
  const inner = [option, ...adders].filter(Boolean).join('; ');
  if (inner) head += ` (${inner})`;
  const mods = kids(el, 'MODIFIER').map(m => {
    const v = num(attr(m, 'BASECOST'), 0) + kids(m, 'ADDER').reduce((s, a) => s + num(attr(a, 'BASECOST'), 0), 0);
    const bits = [attr(m, 'OPTION_ALIAS'), attr(m, 'COMMENTS')].filter(Boolean).join(', ');
    return `${attr(m, 'ALIAS')}${bits ? ` (${bits}; ${frac(v)})` : ` (${frac(v)})`}`;
  });
  return mods.length ? `${head}, ${mods.join(', ')}` : head;
}

function realCost(el) {
  const base = num(attr(el, 'BASECOST'), 0) + num(attr(el, 'LEVELS'), 0) * (ATTACKS[attr(el, 'XMLID')] === 'killing' ? 15 : ATTACKS[attr(el, 'XMLID')] ? 5 : 1)
    + kids(el, 'ADDER').reduce((s, a) => s + num(attr(a, 'BASECOST'), 0), 0);
  const mods = kids(el, 'MODIFIER').map(m => num(attr(m, 'BASECOST'), 0) + kids(m, 'ADDER').reduce((s, a) => s + num(attr(a, 'BASECOST'), 0), 0));
  const adv = mods.filter(v => v > 0).reduce((a, b) => a + b, 0);
  const lim = -mods.filter(v => v < 0).reduce((a, b) => a + b, 0);
  const active = Math.round(base * (1 + adv));
  return { base: Math.round(base), active, real: Math.max(1, Math.round(active / (1 + lim))) };
}

function pad(prefix, i) { return `${prefix}${String(i + 1).padStart(2, '0')}`; }

/** Parse the .hdc into the JSON HERO Designer's HeroSystem6eHeroic.hde export writes. */
export function hdcToHdImport(xmlText, parser = new DOMParser()) {
  const doc = parser.parseFromString(xmlText.replace(/^﻿/, '').replace(/encoding="UTF-16"/i, 'encoding="UTF-8"'), 'text/xml');
  if (doc.querySelector('parsererror')) throw new Error('The Hero file from the server could not be read.');
  const root = doc.documentElement;
  const info = root.querySelector('CHARACTER_INFO');
  const text = tag => info?.querySelector(tag)?.textContent?.trim() || '';
  const c = {
    character_name: attr(info, 'CHARACTER_NAME') || 'Hero NPC', character_title: attr(info, 'ALTERNATE_IDENTITIES'),
    height: attr(info, 'HEIGHT') ? `${(num(attr(info, 'HEIGHT')) * 0.0254).toFixed(2)} m` : '',
    weight: attr(info, 'WEIGHT') ? `${(num(attr(info, 'WEIGHT')) * 0.4536).toFixed(2)} kg` : '',
    eyes: attr(info, 'EYE_COLOR'), hair: attr(info, 'HAIR_COLOR'),
    backgroundText: text('BACKGROUND'), historyText: text('PERSONALITY'), appearance: text('APPEARANCE'), tactics: text('TACTICS'),
    campaignUse: text('CAMPAIGN_USE'), quote: text('QUOTE'), experience: '0', experienceBenefit: '0', size: '',
  };
  const chars = root.querySelector('CHARACTERISTICS');
  for (const [k, key] of Object.entries(KEYS)) {
    const el = chars ? kids(chars, k)[0] : null;
    c[key] = String(BASES[k] + num(attr(el, 'LEVELS'), 0));
  }
  const section = tag => kids(root.querySelector(tag));
  c.equipment = Object.fromEntries(section('EQUIPMENT').map((el, i) => [pad('equipment', i), {
    name: attr(el, 'NAME') || attr(el, 'ALIAS'), text: describe(el),
    damage: ATTACKS[attr(el, 'XMLID')] ? `${attr(el, 'LEVELS')}d6` : '', end: '0', range: '', mass: attr(el, 'WEIGHT') ? `${(num(attr(el, 'WEIGHT')) * 0.4536).toFixed(2)}kg` : '',
    attack: ATTACKS[attr(el, 'XMLID')] ? 'true' : '', defense: attr(el, 'PDLEVELS') || attr(el, 'EDLEVELS') ? 'true' : '', notes: el.querySelector('NOTES')?.textContent?.trim() || '',
  }]));
  c.maneuvers = {};
  c.perks = Object.fromEntries(section('PERKS').map((el, i) => [pad('perk', i), { type: attr(el, 'ALIAS'), points: String(realCost(el).real), text: `${attr(el, 'NAME') ? `${attr(el, 'NAME')} ` : ''}${describe(el)}`, notes: '' }]));
  c.talents = Object.fromEntries(section('TALENTS').map((el, i) => [pad('talent', i), { type: attr(el, 'ALIAS'), points: String(realCost(el).real), text: `${attr(el, 'NAME') ? `${attr(el, 'NAME')}: ` : ''}${describe(el)}`, notes: '' }]));
  c.complications = Object.fromEntries(section('DISADVANTAGES').map((el, i) => [pad('complication', i), {
    type: attr(el, 'ALIAS'), points: String(kids(el, 'ADDER').reduce((s, a) => s + num(attr(a, 'BASECOST'), 0), num(attr(el, 'BASECOST'), 0))), text: describe(el), notes: '',
  }]));
  c.powers = Object.fromEntries(section('POWERS').map((el, i) => {
    const cost = realCost(el);
    return [pad('power', i), {
      name: attr(el, 'NAME') || attr(el, 'ALIAS'), base: String(cost.base), text: `${describe(el)} (${cost.active} Active Points)`, notes: el.querySelector('NOTES')?.textContent?.trim() || '',
      cost: String(cost.real), endurance: String(Math.round(cost.active / 10)), damage: ATTACKS[attr(el, 'XMLID')] ? `${attr(el, 'LEVELS')}d6` : '', compound: 'false',
    }];
  }));
  c.skills = Object.fromEntries(section('SKILLS').map((el, i) => {
    const charName = attr(el, 'CHARACTERISTIC');
    const roll = charName && charName !== 'GENERAL' && KEYS[charName] ? `${9 + Math.round(num(c[KEYS[charName]], 10) / 5) + num(attr(el, 'LEVELS'), 0)}-` : '';
    return [pad('skill', i), {
      name: attr(el, 'NAME'), enhancer: '', text: `${describe(el)}${roll ? ` ${roll}` : ''}`, display: attr(el, 'ALIAS'), attribute: charName || 'GENERAL',
      base: attr(el, 'BASECOST').replace(/\.0$/, ''), levels: attr(el, 'LEVELS') || '0', cost: String(realCost(el).real),
    }];
  }));
  Object.assign(c, { playerName: attr(info, 'PLAYER_NAME'), gmName: attr(info, 'GM'), characterFile: `${c.character_name}.hdc`,
    template: attr(root, 'TEMPLATE') || 'builtIn.Heroic6E.hdt', versionHD: '20220801', timeStamp: new Date().toUTCString(),
    genre: attr(info, 'GENRE'), campaign: attr(info, 'CAMPAIGN_NAME'), version: '2.4', HeroSystem6eHeroic: 'true' });
  return { character: c };
}

export function mapHero6e(response, form = {}, parser) {
  const xml = response.hdcXml || (typeof response === 'string' ? response : '');
  if (!xml) throw new Error('The server returned no Hero file.');
  const data = hdcToHdImport(xml, parser);
  const c = data.character;
  const listHtml = (title, obj) => {
    const rows = Object.values(obj).filter(v => v && (v.text || v.name));
    return rows.length ? `<h3>${title}</h3><ul>${rows.map(v => `<li>${v.name ? `<strong>${escapeHtml(v.name)}</strong>: ` : ''}${escapeHtml(v.text)}</li>`).join('')}</ul>` : '';
  };
  const summary = `<p>${escapeHtml(c.backgroundText)}</p>` + listHtml('Powers', c.powers) + listHtml('Skills', c.skills)
    + listHtml('Perks', c.perks) + listHtml('Talents', c.talents) + listHtml('Complications', c.complications) + listHtml('Equipment', c.equipment);
  return makePlan('character', 'hero6e-npc', {
    summary: `${c.character_name} (${form.level || ''} points)`.replace(' ( points)', ''),
    sheet: 'hero6e-heroic',
    character: { name: c.character_name, bio: summary },
    // The HD Importer mod builds the whole sheet from this; it creates its own character.
    heroImport: `!hero --import ${JSON.stringify(data)}`,
    attributes: Object.fromEntries(Object.values(KEYS).filter(k => k !== 'running' && k !== 'leaping' && k !== 'swimming').map(k => [k, c[k]])),
    hdc: { filename: `${c.character_name}.hdc`, xml },
  });
}
