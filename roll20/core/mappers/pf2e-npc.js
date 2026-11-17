/**
 * PF2e NPC: Foundry pf2e actor JSON (what /webhook/npc-builder returns)
 * → a plan for the "Pathfinder Second Edition by Roll20" sheet.
 *
 * That sheet builds an NPC when `npcdrop_data` changes (updateNpcDrop /
 * calcNpcDrop in the sheet), and fills spell details from the Roll20
 * compendium by spell name. We send compendium-shaped data the same way.
 */
import { makePlan } from '../plan.js';
import { plainText, rowId, signed, titleCase, num, list, escapeHtml } from '../util.js';
import PF2E_SPELLS from '../data/pf2e-spells.js';

const SKILLS = ['acrobatics', 'arcana', 'athletics', 'crafting', 'deception', 'diplomacy', 'intimidation', 'medicine',
  'nature', 'occultism', 'performance', 'religion', 'society', 'stealth', 'survival', 'thievery'];
const SIZES = { tiny: 'tiny', sm: 'small', med: 'medium', lg: 'large', huge: 'huge', grg: 'gargantuan' };
const ACTIONS = { 1: '1 action', 2: '2 actions', 3: '3 actions' };

/** Spell list sent to the generator as `spellMapping`. */
export function pf2eSpellMapping() {
  return PF2E_SPELLS.map(s => ({ name: s.name, id: s.id, packId: 'pf2e.spells-srd', level: s.level }));
}

const val = v => (v && typeof v === 'object' && 'value' in v ? v.value : v);

function abilityMod(sys, k) {
  const a = sys.abilities?.[k];
  if (a && 'mod' in a) return num(a.mod, 0);
  if (a && 'value' in a) return Math.floor((num(a.value, 10) - 10) / 2);
  return 0;
}

function damageText(item) {
  const rolls = Object.values(item.system?.damageRolls || {});
  const [first, ...rest] = rolls;
  return {
    damage: first?.damage || '',
    type: first?.damageType || '',
    extra: rest.map(r => `plus ${r.damage} ${r.damageType || ''}`.trim()).join(' '),
  };
}

function actionCost(item) {
  const type = val(item.system?.actionType);
  if (type === 'reaction') return 'reaction';
  if (type === 'free') return 'free action';
  if (type === 'passive') return '';
  return ACTIONS[num(val(item.system?.actions), 1)] || '1 action';
}

export function mapPf2eNpc(response) {
  const npc = response.foundryNpc || response.npcDesign || response.actor || response;
  if (!npc?.name) throw new Error('The server returned no NPC.');
  const sys = npc.system || {};
  const attrs = sys.attributes || {};
  const items = npc.items || [];
  const level = num(val(sys.details?.level), 1);

  const perception = num(sys.perception?.mod ?? sys.perception?.value ?? attrs.perception?.value, 0);
  const senses = Array.isArray(sys.perception?.senses)
    ? sys.perception.senses.map(s => [titleCase(s.type).toLowerCase(), s.acuity && s.acuity !== 'precise' ? `(${s.acuity})` : '', s.range ? `${s.range} feet` : ''].filter(Boolean).join(' ')).join(', ')
    : String(val(sys.traits?.senses) || '');
  const languages = list(val(sys.details?.languages) ?? val(sys.traits?.languages)).map(titleCase).join(', ');

  // Skills: pf2e v6 keeps them on system.skills; older data uses "lore" items.
  const skillMods = {};
  for (const [k, s] of Object.entries(sys.skills || {})) skillMods[k] = num(s?.base ?? s?.value ?? s?.mod, 0);
  const lores = [];
  for (const it of items.filter(i => i.type === 'lore')) {
    const key = it.name.toLowerCase();
    const v = num(val(it.system?.mod), 0);
    if (SKILLS.includes(key)) skillMods[key] = v; else lores.push({ lore_name: it.name.replace(/\s*lore$/i, '') + ' Lore', lore: signed(v), lore_notes: '' });
  }
  for (const [k, v] of Object.entries(skillMods)) if (!SKILLS.includes(k)) lores.push({ lore_name: titleCase(k), lore: signed(v), lore_notes: '' });

  const iwr = (arr, withValue) => list(arr).map(e => `${titleCase(e.type || e).toLowerCase()}${withValue && e.value ? ` ${e.value}` : ''}`).join(', ');
  const speeds = sys.movement?.speeds || null;
  const land = num(val(attrs.speed) ?? speeds?.land?.value, 25);
  const other = speeds ? Object.entries(speeds).filter(([k, v]) => k !== 'land' && v?.value).map(([k, v]) => `${k} ${v.value} feet`)
    : list(attrs.speed?.otherSpeeds).map(s => `${s.type} ${s.value} feet`);

  // Strikes, abilities, gear.
  const melee = [], ranged = [], activities = [], reactive = [], interaction = [], gear = [];
  for (const it of items) {
    const desc = plainText(it.system?.description?.value);
    if (it.type === 'melee') {
      const d = damageText(it);
      const traits = list(val(it.system?.traits)).filter(t => !/^range/.test(t));
      const isRanged = val(it.system?.weaponType) === 'ranged' || list(val(it.system?.traits)).some(t => /^range/.test(t));
      const row = {
        weapon: it.name, weapon_strike: signed(num(val(it.system?.bonus), 0)), weapon_traits: traits.map(t => t.replace(/-/g, ' ')).join(', '),
        weapon_strike_damage: d.damage, weapon_strike_damage_type: d.type, weapon_strike_damage_additional: d.extra, weapon_notes: desc,
      };
      if (isRanged) {
        const inc = list(val(it.system?.traits)).find(t => /^range/.test(t));
        ranged.push({ ...row, weapon_range: inc ? inc.replace(/\D+/g, ' ').trim() + ' feet' : '' });
      } else melee.push(row);
    } else if (it.type === 'action') {
      const cost = actionCost(it);
      const traits = list(val(it.system?.traits)).join(', ');
      if (cost === 'reaction' || cost === 'free action') {
        const trigger = (desc.match(/Trigger\s*:?\s*([^\n]*?)(?:\s*Effect|$)/i) || [])[1] || '';
        reactive.push({ name: it.name, traits, trigger, description: desc, ...(cost === 'reaction' ? { reaction: 'reaction' } : { free_action: 'free action' }) });
      } else if (!cost) {
        (it.system?.category === 'interaction' ? interaction : reactive).push({ name: it.name, traits, description: desc });
      } else activities.push({ name: it.name, actions: cost, traits, description: desc });
    } else if (['weapon', 'armor', 'equipment', 'consumable', 'shield', 'backpack', 'treasure'].includes(it.type)) {
      gear.push({ item: it.name, description: desc });
    }
  }

  // Spellcasting.
  const entries = items.filter(i => i.type === 'spellcastingEntry');
  const spellRows = { innate: [], focus: [], cantrip: [], spell: [] };
  let spellDc = 0, spellAttack = 0, prepared = false, spontaneous = false;
  for (const e of entries) {
    spellDc = Math.max(spellDc, num(e.system?.spelldc?.dc, 0));
    spellAttack = Math.max(spellAttack, num(e.system?.spelldc?.value, 0));
    const p = val(e.system?.prepared);
    if (p === 'prepared') prepared = true;
    if (p === 'spontaneous') spontaneous = true;
  }
  for (const sp of items.filter(i => i.type === 'spell')) {
    const entry = entries.find(e => e._id === val(sp.system?.location)) || entries[0];
    const known = PF2E_SPELLS.find(s => s.name.toLowerCase() === sp.name.toLowerCase());
    const traits = list(val(sp.system?.traits));
    const isCantrip = traits.includes('cantrip') || known?.cantrip;
    const isFocus = traits.includes('focus') || known?.focus || val(entry?.system?.prepared) === 'focus';
    const isInnate = val(entry?.system?.prepared) === 'innate';
    const row = {
      name: sp.name, spelllevel: String(num(val(sp.system?.level), known?.level ?? 1)), spelldc: String(num(entry?.system?.spelldc?.dc, spellDc)),
      magic_tradition: val(entry?.system?.tradition) || known?.traditions?.[0] || '', description: plainText(sp.system?.description?.value),
    };
    if (isInnate) spellRows.innate.push(row);
    else if (isFocus) spellRows.focus.push(row);
    else if (isCantrip) spellRows.cantrip.push(row);
    else spellRows.spell.push(row);
  }

  const cdata = {
    Category: 'Monsters',
    npc_type: 'creature', level: String(level),
    size: SIZES[sys.traits?.size?.value] || 'medium',
    traits: [sys.traits?.rarity && sys.traits.rarity !== 'common' ? sys.traits.rarity : '', ...list(val(sys.traits))].filter(Boolean).map(titleCase).join(', '),
    perception: signed(perception), senses, languages,
    ...Object.fromEntries(SKILLS.filter(k => k in skillMods).map(k => [k, signed(skillMods[k])])),
    ...Object.fromEntries(['strength', 'dexterity', 'constitution', 'intelligence', 'wisdom', 'charisma'].map(k => [`${k}_modifier`, signed(abilityMod(sys, k.slice(0, 3)))])),
    armor_class: String(num(attrs.ac?.value, 10 + level)), armor_class_notes: attrs.ac?.details || '',
    saving_throws_fortitude: signed(num(sys.saves?.fortitude?.value, 0)),
    saving_throws_reflex: signed(num(sys.saves?.reflex?.value, 0)),
    saving_throws_will: signed(num(sys.saves?.will?.value, 0)),
    saving_throws_notes: sys.attributes?.allSaves?.value || '',
    hit_points: String(num(attrs.hp?.max, num(attrs.hp?.value, 10))),
    hit_points_max: String(num(attrs.hp?.max, num(attrs.hp?.value, 10))), hit_points_notes: attrs.hp?.details || '',
    immunities: iwr(attrs.immunities, false), weaknesses: iwr(attrs.weaknesses, true), resistances: iwr(attrs.resistances, true),
    speed: String(land), speed_notes: other.join(', '),
    ...(spellDc ? { spell_dc: String(spellDc), spell_attack: signed(spellAttack) } : {}),
    ...(spellRows.focus.length ? { casts_focus: 'true', focus_points_max: String(Math.min(3, spellRows.focus.length)) } : {}),
    ...(spellRows.cantrip.length ? { casts_cantrips: 'true' } : {}),
    ...(prepared ? { spellcaster_prepared: 'true' } : {}),
    ...(spontaneous ? { spellcaster_spontaneous: 'true' } : {}),
  };
  const jsonField = (key, rows) => { if (rows.length) cdata[key] = JSON.stringify(rows); };
  jsonField('data-lore', lores);
  jsonField('data-items', gear);
  jsonField('data-interaction-abilities', interaction);
  jsonField('data-free-actions-reactions', reactive);
  jsonField('data-melee-strikes', melee);
  jsonField('data-ranged-strikes', ranged);
  jsonField('data-actions-activities', activities);
  jsonField('data-spellinnate', spellRows.innate);
  jsonField('data-spellfocus', spellRows.focus);
  jsonField('data-cantrip', spellRows.cantrip);
  jsonField('data-spell', spellRows.spell);
  for (const k of Object.keys(cdata)) if (cdata[k] === '') delete cdata[k];

  // Direct fallback, mirroring calcNpcDrop.
  const fallback = { sheet_type: 'npc', character_name: npc.name };
  for (const [k, v] of Object.entries(cdata)) if (!k.startsWith('data-') && k !== 'Category' && !k.startsWith('casts_') && !k.startsWith('spellcaster_')) fallback[k] = v;
  if (cdata.spellcaster_prepared) fallback.spellcaster_prepared = 'prepared';
  if (cdata.spellcaster_spontaneous) fallback.spellcaster_spontaneous = 'spontaneous';
  const rows = (section, arr, fn) => arr.forEach(o => { const r = `repeating_${section}_${rowId()}_`; fallback[`${r}toggles`] = 'display,'; for (const [k, v] of Object.entries(fn(o))) fallback[r + k] = v; });
  rows('lore', lores, o => o);
  rows('items-worn', gear, o => ({ worn_item: o.item, description: o.description }));
  rows('interaction-abilities', interaction, o => o);
  rows('free-actions-reactions', reactive, o => ({ ...o, action_type: o.reaction ? 'Reaction' : o.free_action ? 'Free Action' : '' }));
  rows('melee-strikes', melee, o => o);
  rows('ranged-strikes', ranged, o => o);
  rows('actions-activities', activities, o => o);
  const spellRow = o => ({ name: o.name, spelllevel: o.spelllevel, current_level: o.spelllevel, spelldc: o.spelldc, ...(o.magic_tradition ? { magic_tradition: o.magic_tradition } : {}), description: o.description });
  rows('spellinnate', spellRows.innate, spellRow);
  rows('spellfocus', spellRows.focus, spellRow);
  rows('cantrip', spellRows.cantrip, spellRow);
  rows('normalspells', spellRows.spell, spellRow);
  if (spellRows.spell.length) for (let l = 1; l <= 10; l++) {
    const n = spellRows.spell.filter(s => Number(s.spelllevel) === l).length;
    Object.assign(fallback, { [`casts_level_${l}`]: n ? '1' : '0', [`level_${l}_per_day`]: n, [`level_${l}_per_day_max`]: n });
  }

  const notes = plainText(sys.details?.publicNotes) || plainText(sys.details?.blurb);
  const gm = plainText(sys.details?.privateNotes);
  return makePlan('character', 'pf2e-npc', {
    summary: `${npc.name} (Creature ${level})`,
    sheet: 'pf2e',
    character: {
      name: npc.name,
      bio: notes ? `<p>${escapeHtml(notes).replace(/\n+/g, '</p><p>')}</p>` : '',
      gmnotes: gm ? `<p>${escapeHtml(gm).replace(/\n+/g, '</p><p>')}</p>` : '',
    },
    attributes: { sheet_type: 'npc' },
    drop: {
      attributes: { npcdrop_name: npc.name, npcdrop_uniq: `Monsters:${npc.name}` },
      trigger: { name: 'npcdrop_data', value: JSON.stringify(cdata) },
      verify: 'hit_points_max',
      timeoutMs: 6000,
    },
    fallback: { attributes: fallback },
    token: { bar1: 'hit_points', size: { tiny: 0.5, sm: 1, med: 1, lg: 2, huge: 3, grg: 4 }[sys.traits?.size?.value] || 1 },
  });
}
