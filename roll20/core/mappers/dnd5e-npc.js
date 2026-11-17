/**
 * D&D 5e NPC: Foundry dnd5e actor JSON (what /webhook/dnd5e-npc-builder returns)
 * → a plan for the "D&D 5E by Roll20" (2014) sheet.
 *
 * The sheet builds an NPC itself when `drop_category` is set to "Monsters" with
 * `drop_name` / `drop_data` filled in, exactly as a compendium drag does
 * (see handle_drop / processDrop in the sheet). We hand it compendium-shaped
 * data and keep a direct attribute write as the fallback.
 */
import { makePlan } from '../plan.js';
import { plainText, rowId, signed, titleCase, num, list, escapeHtml } from '../util.js';
import SRD_SPELLS from '../data/dnd5e-spells.js';

const ABILITIES = ['str', 'dex', 'con', 'int', 'wis', 'cha'];
const ABILITY_NAMES = { str: 'strength', dex: 'dexterity', con: 'constitution', int: 'intelligence', wis: 'wisdom', cha: 'charisma' };
const SKILLS = {
  acr: 'acrobatics', ani: 'animal_handling', arc: 'arcana', ath: 'athletics', dec: 'deception', his: 'history',
  ins: 'insight', itm: 'intimidation', inv: 'investigation', med: 'medicine', nat: 'nature', prc: 'perception',
  prf: 'performance', per: 'persuasion', rel: 'religion', slt: 'sleight_of_hand', ste: 'stealth', sur: 'survival',
};
const SIZES = { tiny: 'Tiny', sm: 'Small', med: 'Medium', lg: 'Large', huge: 'Huge', grg: 'Gargantuan' };
const XP = { 0: 10, 0.125: 25, 0.25: 50, 0.5: 100, 1: 200, 2: 450, 3: 700, 4: 1100, 5: 1800, 6: 2300, 7: 2900, 8: 3900, 9: 5000, 10: 5900,
  11: 7200, 12: 8400, 13: 10000, 14: 11500, 15: 13000, 16: 15000, 17: 18000, 18: 20000, 19: 22000, 20: 25000, 21: 33000, 22: 41000,
  23: 50000, 24: 62000, 25: 75000, 26: 90000, 27: 105000, 28: 120000, 29: 135000, 30: 155000 };

export const crText = cr => ({ 0.125: '1/8', 0.25: '1/4', 0.5: '1/2' }[cr] ?? String(cr));
export const profBonus = cr => Math.max(2, Math.floor((Math.max(1, cr) - 1) / 4) + 2);
const mod = score => Math.floor((num(score, 10) - 10) / 2);
const ordinal = n => `${n}${n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th'}`;

/** Spell list sent to the generator as `spellMapping` (it picks from these). */
export function dnd5eSpellMapping() {
  return SRD_SPELLS.map(s => ({ name: s.name, id: s.id, packId: 'dnd5e.spells', level: s.level }));
}

function firstActivity(item) {
  return Object.values(item.system?.activities || {})[0] || null;
}

/** Which statblock section a feature belongs in. */
export function featureSection(item) {
  const act = firstActivity(item);
  const type = act?.activation?.type || item.system?.activation?.type || '';
  const desc = plainText(item.system?.description?.value);
  const first = desc.split(/(?<=\.)\s/)[0] || '';
  if (type === 'legendary') return 'legendary';
  if (type === 'reaction' || /\b(as a|uses? its|use a) reaction\b/i.test(first)) return 'reaction';
  if (type === 'bonus' || /\bbonus action\b/i.test(desc)) return 'bonus';
  if (/multiattack/i.test(item.name)) return 'action';
  if (type === 'action' && /(\bmakes?\b.*\battacks?\b|\bDC \d+|\bas an action\b|\brecharge\b|\bhit:)/i.test(desc)) return 'action';
  if (['action', 'minute', 'hour'].includes(type) && act?.type && act.type !== 'utility') return 'action';
  return 'trait';
}

function attackOf(item, abilities, prof) {
  const act = Object.values(item.system?.activities || {}).find(a => a.type === 'attack');
  if (!act && item.type !== 'weapon') return null;
  const atk = act?.attack || act?.attach || {};
  const kind = (atk.type?.value || (item.system?.range?.long ? 'ranged' : 'melee')) === 'ranged' ? 'Ranged' : 'Melee';
  const props = list(item.system?.properties);
  let ability = atk.ability;
  if (!ability) {
    ability = kind === 'Ranged' ? 'dex' : 'str';
    if (props.includes('fin') && mod(abilities.dex) > mod(abilities.str)) ability = 'dex';
  }
  const abilityMod = mod(abilities[ability]);
  const toHit = abilityMod + prof + num(String(atk.bonus || '').replace(/[^\d-]/g, ''), 0) + num(item.system?.magicalBonus, 0);
  const parts = act?.damage?.parts?.length ? act.damage.parts : [];
  const base = item.system?.damage?.base;
  const dmg = parts[0] || (base?.number && base?.denomination ? { number: base.number, denomination: base.denomination, types: base.types, bonus: '@mod' } : null);
  const dmgText = d => {
    if (!d) return '';
    if (d.custom?.enabled && d.custom.formula) return d.custom.formula.replace(/@mod/g, String(abilityMod));
    const bonusRaw = String(d.bonus ?? '');
    const bonus = bonusRaw.includes('@mod') ? abilityMod + num(bonusRaw.replace('@mod', '').replace(/[^\d-]/g, ''), 0) : num(bonusRaw.replace(/[^\d-]/g, ''), 0);
    return `${d.number || 1}d${d.denomination || 6}${bonus ? (bonus > 0 ? `+${bonus}` : bonus) : ''}`;
  };
  const r = item.system?.range || {};
  const range = kind === 'Ranged' ? `${r.value || 30}${r.long ? `/${r.long}` : ''} ft.` : `${r.reach || r.value || 5} ft.`;
  return {
    kind, toHit, range,
    damage: dmgText(dmg), damageType: list(dmg?.types)[0] || list(base?.types)[0] || '',
    damage2: dmgText(parts[1]), damage2Type: list(parts[1]?.types)[0] || '',
  };
}

function abilityLine(npc) {
  const sys = npc.system || {};
  const a = sys.abilities || {};
  return ABILITIES.map(k => num(a[k]?.value, 10));
}

export function mapDnd5eNpc(response) {
  const npc = response.foundryNpc || response.npcDesign || response.actor || response;
  if (!npc?.name) throw new Error('The server returned no NPC.');
  const sys = npc.system || {};
  const abilities = Object.fromEntries(ABILITIES.map(k => [k, num(sys.abilities?.[k]?.value, 10)]));
  const cr = num(sys.details?.cr, 1);
  const prof = profBonus(cr);
  const attrs = sys.attributes || {};
  const traits = sys.traits || {};

  // Saves and skills, in the "Str +5, Dex +3" form the drop handler parses.
  const saves = ABILITIES.filter(k => num(sys.abilities?.[k]?.proficient) > 0)
    .map(k => [k, mod(abilities[k]) + prof + num(sys.abilities?.[k]?.bonuses?.save, 0)]);
  const skills = Object.entries(SKILLS).filter(([k]) => num(sys.skills?.[k]?.value) > 0)
    .map(([k, name]) => [name, mod(abilities[sys.skills[k].ability || 'dex']) + Math.floor(prof * num(sys.skills[k].value))]);
  const perception = 10 + mod(abilities.wis) + Math.floor(prof * num(sys.skills?.prc?.value, 0));

  const armor = (npc.items || []).find(i => i.type === 'equipment' && i.system?.armor?.value && i.system?.type?.value !== 'shield');
  const acValue = num(attrs.ac?.flat, 0) || num(attrs.ac?.value, 0) || (armor ? num(armor.system.armor.value) + Math.min(mod(abilities.dex), armor.system.type?.value === 'heavy' ? 0 : armor.system.type?.value === 'medium' ? 2 : 99) : 10 + mod(abilities.dex));
  const acType = attrs.ac?.calc === 'natural' ? 'natural armor' : armor ? armor.name.toLowerCase() : '';

  const mv = attrs.movement || {};
  const speed = [`${num(mv.walk, 30)} ft.`, ...['burrow', 'climb', 'fly', 'swim'].filter(k => num(mv[k]) > 0)
    .map(k => `${k} ${num(mv[k])} ft.${k === 'fly' && mv.hover ? ' (hover)' : ''}`)].join(', ');
  const senses = ['blindsight', 'darkvision', 'tremorsense', 'truesight'].filter(k => num(attrs.senses?.[k]) > 0)
    .map(k => `${k} ${num(attrs.senses[k])} ft.`).concat(attrs.senses?.special ? [attrs.senses.special] : []).join(', ');
  const listOf = t => [...list(t?.value).map(titleCase), ...(t?.custom ? String(t.custom).split(/;\s*/) : [])].filter(Boolean).join(', ');
  const languages = listOf(traits.languages);
  const typeLine = [SIZES[traits.size] || 'Medium', titleCase(sys.details?.type?.value || 'humanoid').toLowerCase() + (sys.details?.type?.subtype ? ` (${sys.details.type.subtype})` : '')].join(' ');
  const alignment = sys.details?.alignment || 'unaligned';

  // Sort features and weapons into statblock sections.
  const sections = { trait: [], action: [], bonus: [], reaction: [], legendary: [] };
  for (const item of npc.items || []) {
    if (item.type === 'spell') continue;
    if (item.type === 'equipment' && !Object.keys(item.system?.activities || {}).length) continue;
    const atk = item.type === 'weapon' ? attackOf(item, abilities, prof) : null;
    const desc = plainText(item.system?.description?.value);
    if (atk) { sections.action.push({ name: item.name, desc, atk }); continue; }
    const section = featureSection(item);
    sections[section].push({ name: section === 'bonus' ? `${item.name} (Bonus Action)` : item.name, desc });
  }
  const actions = [...sections.action, ...sections.bonus].sort((a, b) => /multiattack/i.test(b.name) - /multiattack/i.test(a.name));

  // Spells: generator-chosen spells plus any spell items, resolved against the SRD list.
  const spellNames = [...list(response.chosenSpells).map(s => s.name), ...(npc.items || []).filter(i => i.type === 'spell').map(i => i.name)];
  const spells = [...new Set(spellNames.filter(Boolean))].map(name => SRD_SPELLS.find(s => s.name.toLowerCase() === String(name).toLowerCase()) || { name, level: num((npc.items || []).find(i => i.name === name)?.system?.level, 1), desc: '' });
  const castAbility = attrs.spellcasting || 'int';
  if (spells.length) {
    const casterLevel = Math.max(1, num(attrs.spell?.level, 0) || Math.min(20, Math.max(1, Math.round(cr))));
    const dc = 8 + prof + mod(abilities[castAbility]);
    const byLevel = {};
    for (const s of spells) (byLevel[s.level] ||= []).push(s.name.toLowerCase());
    const lines = Object.keys(byLevel).sort((a, b) => a - b).map(l => l === '0' ? `Cantrips (at will): ${byLevel[l].join(', ')}` : `${ordinal(Number(l))} level: ${byLevel[l].join(', ')}`);
    sections.trait.push({
      name: 'Spellcasting',
      desc: `The ${npc.name} is a ${ordinal(casterLevel)}-level spellcaster. Its spellcasting ability is ${titleCase(ABILITY_NAMES[castAbility])} (spell save DC ${dc}, ${signed(prof + mod(abilities[castAbility]))} to hit with spell attacks). It has the following spells prepared:\n${lines.join('\n')}`,
    });
  }

  const legendaryCount = num(sys.resources?.legact?.max, 0);
  const hp = num(attrs.hp?.max, num(attrs.hp?.value, 10));
  const hpFormula = attrs.hp?.formula || '';
  const crString = crText(cr);

  // ── compendium-shaped drop data ──
  const scores = abilityLine(npc);
  const dropData = {
    Category: 'Monsters',
    STR: scores[0], DEX: scores[1], CON: scores[2], INT: scores[3], WIS: scores[4], CHA: scores[5],
    AC: acType ? `${acValue} (${acType})` : String(acValue),
    HP: hpFormula ? `${hp} (${hpFormula})` : String(hp),
    Speed: speed,
    Size: SIZES[traits.size] || 'Medium',
    Type: titleCase(sys.details?.type?.value || 'humanoid') + (sys.details?.type?.subtype ? ` (${sys.details.type.subtype})` : ''),
    Alignment: alignment,
    'Challenge Rating': crString,
    'data-XP': (XP[cr] ?? num(sys.details?.xp?.value, 0)),
    Senses: senses,
    'Passive Perception': perception,
    Languages: languages,
    Resistances: listOf(traits.dr),
    Immunities: listOf(traits.di),
    Vulnerabilities: listOf(traits.dv),
    'Condition Immunities': listOf(traits.ci),
    'Saving Throws': saves.map(([k, v]) => `${titleCase(k)} ${signed(v)}`).join(', '),
    Skills: skills.map(([k, v]) => `${titleCase(k)} ${signed(v)}`).join(', '),
    'data-Traits': JSON.stringify(sections.trait.map(t => ({ Name: t.name, Desc: t.desc }))),
    'data-Actions': JSON.stringify(actions.map(a => a.atk ? {
      Name: a.name, Desc: a.desc, 'Type Attack': `${a.atk.kind} Weapon Attack`, Type: a.atk.kind,
      'Hit Bonus': a.atk.toHit, Reach: a.atk.range, Target: 'one target', Damage: a.atk.damage, 'Damage Type': a.atk.damageType,
      ...(a.atk.damage2 ? { 'Damage 2': a.atk.damage2, 'Damage 2 Type': a.atk.damage2Type } : {}),
    } : { Name: a.name, Desc: a.desc })),
    ...(sections.reaction.length ? { 'data-Reactions': JSON.stringify(sections.reaction.map(r => ({ Name: r.name, Desc: r.desc }))) } : {}),
    ...(sections.legendary.length ? { 'data-Legendary Actions': JSON.stringify(sections.legendary.map(r => ({ Name: r.name, Desc: r.desc }))), 'data-LANum': legendaryCount || 3 } : {}),
  };
  for (const k of Object.keys(dropData)) if (dropData[k] === '') delete dropData[k];

  // ── spell rows: written directly in both paths (the drop needs a licensed compendium for spells) ──
  const attributes = {};
  for (const s of spells) {
    const lvl = s.level > 0 ? s.level : 'cantrip';
    const row = `repeating_spell-${lvl}_${rowId()}_`;
    Object.assign(attributes, {
      [`${row}spellname`]: s.name, [`${row}spelllevel`]: lvl, [`${row}spellschool`]: (s.school || '').toLowerCase(),
      [`${row}spellcastingtime`]: s.time || '1 action', [`${row}spellrange`]: s.range || '', [`${row}spellduration`]: s.duration || '',
      [`${row}spellcomp_v`]: !s.components || s.components.includes('V') ? '{{v=1}}' : '0',
      [`${row}spellcomp_s`]: !s.components || s.components.includes('S') ? '{{s=1}}' : '0',
      [`${row}spellcomp_m`]: s.components?.includes('M') ? '{{m=1}}' : '0',
      [`${row}spellcomp_materials`]: s.material || '', [`${row}spellconcentration`]: s.concentration ? '{{concentration=1}}' : '0',
      [`${row}spellritual`]: s.ritual ? '{{ritual=1}}' : '0', [`${row}spelldescription`]: s.desc || '',
      [`${row}spell_ability`]: `@{${ABILITY_NAMES[castAbility]}_mod}+`, [`${row}spelloutput`]: 'SPELLCARD', [`${row}spellprepared`]: '1',
    });
  }
  if (spells.length) Object.assign(attributes, { npcspellcastingflag: '1', spellcasting_ability: `@{${ABILITY_NAMES[castAbility]}_mod}+` });

  // ── direct fallback: what processDrop would have written ──
  const fallback = {
    npc: '1', 'npc_options-flag': '0', npc_name: npc.name,
    npc_type: `${typeLine}, ${alignment}`, npc_ac: acValue, npc_actype: acType,
    hp: { current: hp, max: hp }, hp_max: hp, npc_hpformula: hpFormula, npc_speed: speed,
    npc_senses: [senses, `passive Perception ${perception}`].filter(Boolean).join(', '),
    npc_languages: languages, npc_challenge: crString, npc_xp: (XP[cr] ?? num(sys.details?.xp?.value, 0)),
    npc_resistances: dropData.Resistances || '', npc_immunities: dropData.Immunities || '',
    npc_vulnerabilities: dropData.Vulnerabilities || '', npc_condition_immunities: dropData['Condition Immunities'] || '',
    pb: prof, npc_saving_flag: saves.length ? '1' : '', npc_skills_flag: skills.length ? '1' : '',
  };
  ABILITIES.forEach((k, i) => {
    const full = ABILITY_NAMES[k];
    Object.assign(fallback, { [`${full}_base`]: scores[i], [full]: scores[i], [`${full}_mod`]: mod(scores[i]) });
  });
  for (const [k, v] of saves) Object.assign(fallback, { [`npc_${k}_save_base`]: v, [`npc_${k}_save`]: v, [`npc_${k}_save_flag`]: '1' });
  for (const [k, v] of skills) Object.assign(fallback, { [`npc_${k}_base`]: v, [`npc_${k}`]: v, [`npc_${k}_flag`]: '1' });
  for (const t of sections.trait) {
    const row = `repeating_npctrait_${rowId()}_`;
    Object.assign(fallback, { [`${row}name`]: `${t.name}.`, [`${row}desc`]: t.desc });
  }
  for (const a of actions) {
    const row = `repeating_npcaction_${rowId()}_`;
    Object.assign(fallback, { [`${row}name`]: a.name, [`${row}description`]: a.desc, [`${row}npc_options-flag`]: '0' });
    if (a.atk) Object.assign(fallback, {
      [`${row}attack_flag`]: 'on', [`${row}attack_display_flag`]: '{{attack=1}}', [`${row}attack_options`]: '{{attack=1}}',
      [`${row}attack_type`]: a.atk.kind, [`${row}attack_range`]: a.atk.range, [`${row}attack_target`]: 'one target',
      [`${row}attack_tohit`]: a.atk.toHit, [`${row}attack_damage`]: a.atk.damage, [`${row}attack_damagetype`]: a.atk.damageType,
      ...(a.atk.damage2 ? { [`${row}attack_damage2`]: a.atk.damage2, [`${row}attack_damagetype2`]: a.atk.damage2Type } : {}),
    });
  }
  if (sections.reaction.length) fallback.npcreactionsflag = '1';
  for (const r of sections.reaction) {
    const row = `repeating_npcreaction_${rowId()}_`;
    Object.assign(fallback, { [`${row}name`]: `${r.name}.`, [`${row}desc`]: r.desc });
  }
  if (sections.legendary.length) fallback.npc_legendary_actions = legendaryCount || 3;
  for (const r of sections.legendary) {
    const row = `repeating_npcaction-l_${rowId()}_`;
    Object.assign(fallback, { [`${row}name`]: r.name, [`${row}description`]: r.desc, [`${row}npc_options-flag`]: '0' });
  }

  const beacon = beaconNpc({ npc, abilities, cr: crString, prof, acValue, acType, hp, hpFormula, mv, attrs, traits, typeLine, alignment,
    languages, saves, skills, sections, legendaryCount });

  const bio = plainText(sys.details?.biography?.value);
  return makePlan('character', 'dnd5e-npc', {
    beacon,
    summary: `${npc.name} (CR ${crString} ${typeLine.toLowerCase()})`,
    sheet: 'dnd5e-2014',
    character: { name: npc.name, bio: bio ? `<p>${escapeHtml(bio).replace(/\n+/g, '</p><p>')}</p>` : '' },
    attributes,
    drop: {
      attributes: { drop_name: npc.name, drop_data: JSON.stringify(dropData), drop_content: '' },
      trigger: { name: 'drop_category', value: 'Monsters' },
      verify: 'npc_name',
      timeoutMs: 5000,
    },
    fallback: { attributes: fallback },
    token: { bar1: 'hp', size: { tiny: 0.5, sm: 1, med: 1, lg: 2, huge: 3, grg: 4 }[traits.size] || 1 },
  });
}

/* ── D&D 2024 sheet ("Beacon" sheet, data in the `store` attribute) ──
 * Shapes copied from what the 2024 sheet's own NPC editor writes. The page
 * bridge gives each integrant its id and links Damage rows to their Attack. */

const DAMAGE_TYPES = ['Acid', 'Bludgeoning', 'Cold', 'Fire', 'Force', 'Lightning', 'Necrotic', 'Piercing', 'Poison', 'Psychic', 'Radiant', 'Slashing', 'Thunder'];
const CONDITIONS = ['Blinded', 'Charmed', 'Deafened', 'Exhaustion', 'Frightened', 'Grappled', 'Incapacitated', 'Invisible', 'Paralyzed', 'Petrified', 'Poisoned', 'Prone', 'Restrained', 'Stunned', 'Unconscious'];

function parseDice(text) {
  const m = String(text || '').replace(/\s+/g, '').match(/^(\d*)d(\d+)([+-]\d+)?$/i);
  if (!m) return null;
  return { _diceCount: Number(m[1] || 1), diceSize: `d${m[2]}`, _bonus: Number(m[3] || 0) };
}

export function beaconNpc({ npc, abilities, cr, prof, acValue, acType, hp, hpFormula, mv, attrs, traits, typeLine, alignment, languages, saves, skills, sections, legendaryCount }) {
  const ints = [];
  const add = (fields, ref) => { ints.push({ ...(ref ? { _ref: ref } : {}), ...fields }); };
  ABILITIES.forEach(k => add({ type: 'Ability Score', ability: titleCase(ABILITY_NAMES[k]), calculation: 'Set Value', label: 'Override (Custom)', name: '', source: 'Custom', valueFormula: { flatValue: abilities[k] } }));
  add({ _id: 'armor-class-override-total', type: 'Armor Class', calculation: 'Set Value', defaultAbility: false, label: 'Unknown', name: '', source: 'Custom', valueFormula: { flatValue: acValue } });
  add({ type: 'Hit Points', calculation: 'Set Value', hitpointType: 'Maximum', isFixed: false, isTemp: false, label: 'Override Max HP', name: '', source: 'Custom', valueFormula: { flatValue: hp } });
  for (const [k, label] of [['walk', 'Walk'], ['burrow', 'Burrow'], ['climb', 'Climb'], ['fly', 'Fly'], ['swim', 'Swim']]) {
    const v = num(mv[k], k === 'walk' ? 30 : 0);
    if (v > 0) add({ type: 'Speed', calculation: 'Set Value', name: '', source: 'Custom', speed: label, valueFormula: { flatValue: v } });
  }
  for (const k of ['blindsight', 'darkvision', 'tremorsense', 'truesight']) {
    const v = num(attrs.senses?.[k], 0);
    if (v > 0) add({ type: 'Sense', name: titleCase(k), calculation: 'Set Base', ignoreValue: false, source: '', valueFormula: { flatValue: v } });
  }
  for (const lang of String(languages || '').split(/,\s*/).filter(Boolean)) add({ type: 'Language', name: lang, source: '' });
  const defenses = (t, defense) => [...list(t?.value), ...(t?.custom ? String(t.custom).split(/;\s*/) : [])].map(titleCase).filter(Boolean)
    .forEach(name => {
      if (DAMAGE_TYPES.includes(name)) add({ type: 'Defense', damage: name, defense, name: `${defense} to ${name}`, source: '' });
      else if (CONDITIONS.includes(name)) add({ type: 'Defense', condition: name, defense, name: `${defense} to ${name}`, source: '' });
    });
  defenses(traits.dr, 'Resistance');
  defenses(traits.di, 'Immunity');
  defenses(traits.dv, 'Vulnerability');
  defenses(traits.ci, 'Immunity');
  for (const [k] of saves) add({ type: 'Proficiency', category: 'Saving Throw', proficiency: titleCase(ABILITY_NAMES[k]), proficiencyLevel: 'Proficient', rollAbility: 'Query Attribute', increaseIfAlreadyAt: false, notes: '', name: '', source: 'Custom' });
  for (const [name, total] of skills) {
    const key = Object.keys(SKILLS).find(k => SKILLS[k] === name);
    const ability = { acr: 'dex', ani: 'wis', arc: 'int', ath: 'str', dec: 'cha', his: 'int', ins: 'wis', itm: 'cha', inv: 'int', med: 'wis', nat: 'int', prc: 'wis', prf: 'cha', per: 'cha', rel: 'int', slt: 'dex', ste: 'dex', sur: 'wis' }[key] || 'dex';
    const level = total - mod(abilities[ability]) >= prof * 2 ? 'Expertise' : 'Proficient';
    add({ type: 'Proficiency', category: 'Skill', proficiency: titleCase(name.replace(/_/g, ' ')), proficiencyLevel: level, rollAbility: 'Query Attribute', increaseIfAlreadyAt: false, notes: '', name: '', source: 'Custom' });
  }
  const traitRefs = [];
  sections.trait.forEach((t, i) => { traitRefs.push(`trait${i}`); add({ type: 'Features', name: t.name, description: t.desc, source: 'Species' }, `trait${i}`); });
  const actionOf = (a, actionType, i) => {
    if (!a.atk) return add({ type: 'Action', actionType, name: a.name.replace(/ \(Bonus Action\)$/, ''), description: a.desc, excludeFamilialResources: false, source: '' });
    const ref = `atk${i}`;
    add({ type: 'Attack', actionType, attack: { bonusFlat: Number(a.atk.toHit) || 0, type: a.atk.kind }, name: a.name, description: a.desc,
      range: String(a.atk.range || '').replace(/[^\d/]/g, '') || '5', source: '' }, ref);
    [[a.atk.damage, a.atk.damageType], [a.atk.damage2, a.atk.damage2Type]].forEach(([d, type], j) => {
      const dice = parseDice(d);
      if (dice) add({ type: 'Damage', _parentRef: ref, ability: 'none', critDiceSize: '', overrideCrit: false, damageType: titleCase(type || ''), name: `${a.name} Damage${j ? ' 2' : ''}`, source: '', ...dice });
    });
  };
  const ordered = [...sections.action].sort((a, b) => /multiattack/i.test(b.name) - /multiattack/i.test(a.name));
  ordered.forEach((a, i) => actionOf(a, 'Action', i));
  sections.bonus.forEach((a, i) => actionOf(a, 'Bonus Action', 100 + i));
  sections.reaction.forEach((a, i) => actionOf(a, 'Reaction', 200 + i));
  sections.legendary.forEach((a, i) => actionOf(a, 'Legendary', 300 + i));
  return {
    integrants: ints,
    traitRefs,
    store: {
      npc: { challengeRating: String(cr), acNotes: acType || '', rollHP: hpFormula || '', ...(sections.legendary.length ? { legendaryActionNum: legendaryCount || 3 } : {}) },
      character: { creatureType: titleCase(typeLine.split(' ').slice(1).join(' ')) },
      about: { characteristics: { size: typeLine.split(' ')[0], alignment: titleCase(alignment) } },
    },
  };
}
