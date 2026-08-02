/**
 * Battle resolver (#76) — deterministic mass-combat formula (numbers × level
 * × morale × terrain) for a fight between two Army documents. Casualties are
 * written back onto each army's unit roster. Also resolves sieges (#77)
 * against Settlement documents and applies commander bonuses (#78).
 */

import { FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeArmy } from './army.js';
import { sanitizeSettlement } from './sanitizer.js';

export const TERRAIN_TYPES = ['plains', 'forest', 'hills', 'mountains', 'urban', 'swamp'];

/** Flat power added per commander level (#78) — small next to a real roster's power. */
const COMMANDER_POWER_PER_LEVEL = 10;

/** Per-terrain power multiplier by unit type; unlisted types default to 1. */
export const TERRAIN_MODIFIERS = {
  plains:    { cavalry: 1.25 },
  forest:    { cavalry: 0.6, archers: 1.2, siege: 0.7 },
  hills:     { archers: 1.2, cavalry: 0.85, siege: 0.85 },
  mountains: { spearmen: 1.1, cavalry: 0.5, siege: 0.5 },
  urban:     { spearmen: 1.1, siege: 1.3, cavalry: 0.6, archers: 0.9 },
  swamp:     { cavalry: 0.5, spearmen: 0.9, mages: 1.1, siege: 0.6 },
};

/**
 * Total combat power of an army on a given terrain: Σ(count × level × morale% × terrainMul),
 * plus a flat bonus from the army's commander level, if any (#78).
 */
export function armyPower(army, terrain, commanderLevel = 0) {
  const mods = TERRAIN_MODIFIERS[terrain] || {};
  const rosterPower = (army.units || []).reduce((sum, u) => {
    const count = Math.max(0, Number(u.count) || 0);
    const level = Math.max(1, Number(u.level) || 1);
    const morale = Math.max(0, Math.min(100, Number(u.morale) || 0)) / 100;
    const terrainMul = mods[u.type] ?? 1;
    return sum + count * level * morale * terrainMul;
  }, 0);
  return rosterPower + Math.max(0, Number(commanderLevel) || 0) * COMMANDER_POWER_PER_LEVEL;
}

function applyCasualties(army, pct) {
  return (army.units || []).map(u => {
    const count = Math.max(0, Number(u.count) || 0);
    const lost = Math.min(count, Math.round(count * pct));
    return { id: u.id, type: u.type, lost, remaining: count - lost };
  });
}

/**
 * Resolve a battle between an attacking and defending army on a terrain.
 * Deterministic given the two rosters — no dice. The losing side takes
 * heavier casualties; both sides bleed in proportion to the opposing power.
 *
 * @param {import('./types.js').Army} attacker
 * @param {import('./types.js').Army} defender
 * @param {string} terrain
 * @param {number} attackerCommanderLevel  flat power bonus from the attacker's commander (#78)
 * @param {number} defenderCommanderLevel  flat power bonus from the defender's commander (#78)
 */
export function resolveBattle(attacker, defender, terrain = 'plains', attackerCommanderLevel = 0, defenderCommanderLevel = 0) {
  const attackerPower = armyPower(attacker, terrain, attackerCommanderLevel);
  const defenderPower = armyPower(defender, terrain, defenderCommanderLevel);
  const totalPower = attackerPower + defenderPower;
  const winner = attackerPower === defenderPower
    ? 'draw'
    : (attackerPower > defenderPower ? 'attacker' : 'defender');

  const attackerCasualtyPct = totalPower > 0
    ? Math.min(1, (defenderPower / totalPower) * (winner === 'attacker' ? 0.4 : 0.75))
    : 0;
  const defenderCasualtyPct = totalPower > 0
    ? Math.min(1, (attackerPower / totalPower) * (winner === 'defender' ? 0.4 : 0.75))
    : 0;

  return {
    terrain,
    attackerPower,
    defenderPower,
    winner,
    attackerCasualtyPct,
    defenderCasualtyPct,
    attackerCasualties: applyCasualties(attacker, attackerCasualtyPct),
    defenderCasualties: applyCasualties(defender, defenderCasualtyPct),
  };
}

async function writeCasualties(doc, casualties) {
  if (!doc) return;
  const cur = foundry.utils.deepClone(getSettlement(doc) || {});
  cur.units = (cur.units || []).map(u => {
    const c = casualties.find(x => x.id === u.id);
    return c ? { ...u, count: c.remaining } : u;
  });
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, sanitizeArmy(cur));
}

/** Write a resolveBattle() result's casualties back onto both army documents. */
export async function applyBattleResult(attackerDoc, defenderDoc, result) {
  await writeCasualties(attackerDoc, result.attackerCasualties);
  await writeCasualties(defenderDoc, result.defenderCasualties);
}

function gmWhisper() {
  return game.users?.filter(u => u.isGM).map(u => u.id) ?? [];
}

/** Post a GM-whispered chat card summarizing a resolved battle (#82). */
export function postBattleChatCard(attackerDoc, defenderDoc, result) {
  const summarize = (casualties) => casualties.filter(c => c.lost > 0)
    .map(c => `${c.lost} ${c.type}`).join(', ') || 'none';
  const winnerName = result.winner === 'draw'
    ? 'Neither side'
    : (result.winner === 'attacker' ? attackerDoc.name : defenderDoc.name);
  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-shield-halved"></i> Battle: ${attackerDoc.name} vs ${defenderDoc.name}</h3>
      <p>Terrain: ${result.terrain}. <strong>${winnerName}</strong> ${result.winner === 'draw' ? 'fought to a draw' : 'prevailed'}.</p>
      <details>
        <summary>Casualties</summary>
        <p><strong>${attackerDoc.name}:</strong> ${summarize(result.attackerCasualties)}</p>
        <p><strong>${defenderDoc.name}:</strong> ${summarize(result.defenderCasualties)}</p>
      </details>`,
    whisper: gmWhisper(),
  }).catch(() => {});
}

/**
 * Resolve a siege: an attacking army assaulting a settlement (not another
 * army) (#77). Damage is the attacker's power minus the settlement's
 * hardness (per-hit reduction); if that doesn't clear the damageThreshold
 * (the settlement shrugs off small attacks entirely, mirroring PF2e object
 * damage rules) no damage gets through. HP hitting 0 flips `occupied`.
 *
 * @param {import('./types.js').Army} attacker
 * @param {import('./types.js').Settlement} settlement
 * @param {string} terrain
 * @param {number} commanderLevel  flat power bonus from the attacker's commander (#78)
 */
export function resolveSiege(attacker, settlement, terrain = 'urban', commanderLevel = 0) {
  const attackerPower = armyPower(attacker, terrain, commanderLevel);
  const stats = settlement.stats || {};
  const rawDamage = Math.max(0, Math.round(attackerPower - (Number(stats.hardness) || 0)));
  const damage = rawDamage >= (Number(stats.damageThreshold) || 0) ? rawDamage : 0;
  const hpBefore = Number(stats.hp) || 0;
  const hpAfter = Math.max(0, hpBefore - damage);

  return {
    terrain,
    attackerPower,
    damage,
    hpBefore,
    hpAfter,
    occupied: hpAfter <= 0,
  };
}

/** Write a resolveSiege() result's HP and occupied state back onto the settlement document. */
export async function applySiegeResult(settlementDoc, result) {
  if (!settlementDoc) return;
  const cur = sanitizeSettlement(getSettlement(settlementDoc) || {});
  cur.stats.hp = result.hpAfter;
  cur.stats.occupied = result.occupied;
  await settlementDoc.setFlag(FLAG_SCOPE, FLAG_KEY, cur);
}

/** Post a GM-whispered chat card summarizing a resolved siege (#82). */
export function postSiegeChatCard(attackerDoc, settlementDoc, result) {
  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-chess-rook"></i> Siege: ${attackerDoc.name} vs ${settlementDoc.name}</h3>
      <p>Terrain: ${result.terrain}. ${result.occupied ? `<strong>${settlementDoc.name} has fallen.</strong>` : `${settlementDoc.name} holds.`}</p>
      <details>
        <summary>Damage</summary>
        <p>${result.damage} damage dealt. HP: ${result.hpBefore} → ${result.hpAfter}.</p>
      </details>`,
    whisper: gmWhisper(),
  }).catch(() => {});
}
