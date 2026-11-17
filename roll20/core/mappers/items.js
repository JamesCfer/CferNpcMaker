/**
 * Items (D&D 5e and PF2e) → a handout card plus a "give" recipe that adds the
 * item to a character's sheet.
 *
 * Roll20 has no world-level item list like Foundry's Items directory, so the
 * item lives as a handout and is added to a character on demand.
 */
import { makePlan } from '../plan.js';
import { plainText, rowId, titleCase, num, list, escapeHtml } from '../util.js';

/* ── D&D 5e ─────────────────────────────────────────────── */

const DND_PROPS = { ada: 'Adamantine', amm: 'Ammunition', fin: 'Finesse', fir: 'Firearm', foc: 'Focus', hvy: 'Heavy', lgt: 'Light',
  lod: 'Loading', mgc: 'Magical', rch: 'Reach', rel: 'Reload', ret: 'Returning', sil: 'Silvered', spc: 'Special', thr: 'Thrown',
  two: 'Two-Handed', ver: 'Versatile', stealthDisadvantage: 'Stealth Disadvantage' };
const DND_RARITY = { common: 'Common', uncommon: 'Uncommon', rare: 'Rare', veryRare: 'Very Rare', legendary: 'Legendary', artifact: 'Artifact' };

export function dnd5eItemType(item) {
  const t = item.system?.type?.value || '';
  if (item.type === 'weapon') return /R$/.test(t) || t === 'simpleR' || t === 'martialR' || item.system?.range?.long ? 'Ranged Weapon' : 'Melee Weapon';
  if (item.type === 'equipment') {
    if (['light', 'medium', 'heavy'].includes(t)) return `${titleCase(t)} Armor`;
    if (t === 'shield') return 'Shield';
    return 'Wondrous Item';
  }
  if (item.type === 'consumable') return t === 'potion' ? 'Potion' : t === 'scroll' ? 'Scroll' : t === 'ammo' ? 'Ammunition' : 'Adventuring Gear';
  if (item.type === 'tool') return 'Tool';
  if (item.type === 'container') return 'Adventuring Gear';
  return 'Adventuring Gear';
}

export function dnd5eItemFacts(item) {
  const s = item.system || {};
  const props = list(s.properties).map(p => DND_PROPS[p] || titleCase(p));
  const base = s.damage?.base;
  const damage = base?.number && base?.denomination ? `${base.number}d${base.denomination}` : '';
  const vers = s.damage?.versatile;
  const price = s.price?.value ? `${s.price.value} ${s.price.denomination || 'gp'}` : '';
  const weight = typeof s.weight === 'object' ? num(s.weight?.value, 0) : num(s.weight, 0);
  const range = s.range?.long ? `${s.range.value}/${s.range.long}` : '';
  return {
    type: dnd5eItemType(item), rarity: DND_RARITY[s.rarity] || (s.rarity ? titleCase(s.rarity) : ''),
    attunement: s.attunement === 'required' || s.attunement === 1 ? 'requires attunement' : '',
    damage, damageType: list(base?.types)[0] || '', versatile: vers?.number ? `${vers.number}d${vers.denomination}` : (props.includes('Versatile') && base?.denomination ? `${base.number || 1}d${num(base.denomination) + 2}` : ''),
    props, price, weight, range, ac: s.armor?.value || '', magicBonus: num(s.magicalBonus, 0),
    quantity: Math.max(1, num(s.quantity, 1)), desc: plainText(s.description?.value),
  };
}

function dnd5eGive(item) {
  const f = dnd5eItemFacts(item);
  const mods = [f.magicBonus ? `${f.type.includes('Weapon') ? 'Weapon Attacks' : 'AC'} +${f.magicBonus}` : ''].filter(Boolean).join(', ');
  const data = {
    Category: 'Items', 'Item Type': f.type, itemcount: f.quantity,
    ...(f.damage ? { Damage: f.damage, 'Damage Type': f.damageType } : {}),
    ...(f.versatile ? { 'Alternate Damage': f.versatile, 'Alternate Damage Type': f.damageType } : {}),
    ...(f.props.length ? { Properties: f.props.join(', ') } : {}),
    ...(f.weight ? { Weight: f.weight } : {}), ...(f.range ? { Range: f.range } : {}), ...(f.ac ? { AC: f.ac } : {}),
    ...(mods ? { Modifiers: mods } : {}),
  };
  const row = `repeating_inventory_${rowId()}_`;
  return {
    sheet: 'dnd5e-2014',
    drop: { attributes: { drop_name: item.name, drop_data: JSON.stringify(data), drop_content: f.desc }, trigger: { name: 'drop_category', value: 'Items' } },
    rows: {
      [`${row}itemname`]: item.name, [`${row}itemcount`]: f.quantity, [`${row}itemweight`]: f.weight || '',
      [`${row}itemproperties`]: f.props.join(', '), [`${row}itemcontent`]: f.desc,
      [`${row}itemmodifiers`]: [`Item Type: ${f.type}`, f.damage && `Damage: ${f.damage}`, f.damageType && `Damage Type: ${f.damageType}`, f.ac && `AC: ${f.ac}`, mods].filter(Boolean).join(', '),
    },
  };
}

/* ── PF2e ───────────────────────────────────────────────── */

export function pf2eItemFacts(item) {
  const s = item.system || {};
  const price = s.price?.value && typeof s.price.value === 'object'
    ? Object.entries(s.price.value).filter(([, v]) => v).map(([k, v]) => `${v} ${k}`).join(', ') : '';
  const bulk = s.bulk?.value ?? s.weight?.value ?? '';
  const traits = [s.traits?.rarity && s.traits.rarity !== 'common' ? s.traits.rarity : '', ...list(s.traits?.value)].filter(Boolean).map(titleCase);
  const dmg = s.damage?.die ? `${Math.max(1, num(s.runes?.striking, 0) + 1) * Math.max(1, num(s.damage.dice, 1))}${s.damage.die} ${s.damage.damageType || ''}`.trim() : '';
  return {
    type: titleCase(item.type), level: num(s.level?.value, 0), traits, price, bulk: bulk === 0.1 ? 'L' : String(bulk || '-'),
    damage: dmg, group: s.group ? titleCase(s.group) : '', ac: s.acBonus ? `+${s.acBonus}` : '',
    quantity: Math.max(1, num(s.quantity, 1)), desc: plainText(s.description?.value),
  };
}

function pf2eGive(item) {
  const f = pf2eItemFacts(item);
  const notes = [f.damage && `Damage ${f.damage}`, f.ac && `AC ${f.ac}`, f.price && `Price ${f.price}`, f.desc].filter(Boolean).join('\n');
  return {
    sheet: 'pf2e',
    rowsBySheetType: {
      npc: { section: 'items-worn', fields: { toggles: 'display,', worn_item: item.name, traits: f.traits.join(', '), description: notes } },
      pc: { section: 'items-other', fields: { toggles: 'display,', other_item: item.name, other_quantity: f.quantity, other_bulk: f.bulk === '-' ? '' : f.bulk, other_item_notes: notes } },
    },
  };
}

/* ── shared card + plans ────────────────────────────────── */

export function itemCardHtml(item, system) {
  const f = system === 'pf2e' ? pf2eItemFacts(item) : dnd5eItemFacts(item);
  const sub = system === 'pf2e'
    ? [`${f.type} ${f.level}`, f.traits.join(', ')].filter(Boolean).join(' · ')
    : [f.type, f.rarity && f.rarity.toLowerCase(), f.attunement && `(${f.attunement})`].filter(Boolean).join(', ');
  const facts = system === 'pf2e'
    ? [f.price && ['Price', f.price], ['Bulk', f.bulk], f.damage && ['Damage', f.damage], f.group && ['Group', f.group], f.ac && ['AC bonus', f.ac]]
    : [f.price && ['Price', f.price], f.weight && ['Weight', `${f.weight} lb.`], f.damage && ['Damage', `${f.damage} ${f.damageType}`], f.range && ['Range', `${f.range} ft.`], f.ac && ['AC', f.ac], f.props.length && ['Properties', f.props.join(', ')]];
  return `<h3>${escapeHtml(item.name)}</h3><p><em>${escapeHtml(sub)}</em></p>`
    + facts.filter(Boolean).map(([k, v]) => `<p><strong>${escapeHtml(k)}:</strong> ${escapeHtml(v)}</p>`).join('')
    + (f.desc ? `<p>${escapeHtml(f.desc).replace(/\n+/g, '</p><p>')}</p>` : '');
}

export function giveRecipe(item, system) {
  return { name: item.name, ...(system === 'pf2e' ? pf2eGive(item) : dnd5eGive(item)) };
}

export function mapItem(system) {
  return response => {
    const item = response.foundryItem || response.item || response;
    if (!item?.name) throw new Error('The server returned no item.');
    return makePlan('handouts', `${system}-item`, {
      summary: item.name,
      handouts: [{ name: item.name, notes: itemCardHtml(item, system), give: giveRecipe(item, system) }],
    });
  };
}
