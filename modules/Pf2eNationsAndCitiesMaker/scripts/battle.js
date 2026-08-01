/**
 * Battle resolver (#76) — deterministic mass-combat formula (numbers × level
 * × morale × terrain) for a fight between two Army documents. Casualties are
 * written back onto each army's unit roster.
 */

import { FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeArmy } from './army.js';

export const TERRAIN_TYPES = ['plains', 'forest', 'hills', 'mountains', 'urban', 'swamp'];

/** Per-terrain power multiplier by unit type; unlisted types default to 1. */
export const TERRAIN_MODIFIERS = {
  plains:    { cavalry: 1.25 },
  forest:    { cavalry: 0.6, archers: 1.2, siege: 0.7 },
  hills:     { archers: 1.2, cavalry: 0.85, siege: 0.85 },
  mountains: { spearmen: 1.1, cavalry: 0.5, siege: 0.5 },
  urban:     { spearmen: 1.1, siege: 1.3, cavalry: 0.6, archers: 0.9 },
  swamp:     { cavalry: 0.5, spearmen: 0.9, mages: 1.1, siege: 0.6 },
};

/** Total combat power of an army on a given terrain: Σ(count × level × morale% × terrainMul). */
export function armyPower(army, terrain) {
  const mods = TERRAIN_MODIFIERS[terrain] || {};
  return (army.units || []).reduce((sum, u) => {
    const count = Math.max(0, Number(u.count) || 0);
    const level = Math.max(1, Number(u.level) || 1);
    const morale = Math.max(0, Math.min(100, Number(u.morale) || 0)) / 100;
    const terrainMul = mods[u.type] ?? 1;
    return sum + count * level * morale * terrainMul;
  }, 0);
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
 */
export function resolveBattle(attacker, defender, terrain = 'plains') {
  const attackerPower = armyPower(attacker, terrain);
  const defenderPower = armyPower(defender, terrain);
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
