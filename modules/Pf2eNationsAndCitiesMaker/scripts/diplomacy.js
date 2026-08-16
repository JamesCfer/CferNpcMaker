/**
 * Treaty lifecycle (#84, #85) and war/peace/succession flows (#88, #89, #91).
 * Treaties live on a Nation document's own `treaties` list. Expiry is checked
 * on every calendar day-advance tick rather than through Pf2eCalendarTimeline's
 * own event scheduler, mirroring how mercenary contract expiry already
 * piggybacks on the same hook.
 */

import { FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { cmpDate } from './army.js';

export function gmWhisper() {
  return game.users?.filter(u => u.isGM).map(u => u.id) ?? [];
}

function setRelationEntry(s, nationId, relation, score) {
  s.relations = s.relations || [];
  let entry = s.relations.find(r => r.nationId === nationId);
  if (!entry) { entry = { nationId, relation, score }; s.relations.push(entry); }
  else { entry.relation = relation; entry.score = score; }
}

/** Drops any treaty whose expiresOn has passed and posts a chat reminder for each (#85). */
export async function processTreatyExpiry(doc, s, currentDate) {
  if (!doc || !currentDate || !Array.isArray(s.treaties) || !s.treaties.length) return;

  const remaining = [];
  const expired = [];
  for (const treaty of s.treaties) {
    if (treaty.expiresOn && cmpDate(currentDate, treaty.expiresOn) >= 0) expired.push(treaty);
    else remaining.push(treaty);
  }
  if (!expired.length) return;

  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, { ...s, treaties: remaining });

  for (const treaty of expired) {
    const partner = game.journal?.get(treaty.partnerNationId);
    ChatMessage.create({
      content: `<h3><i class="fa-solid fa-file-circle-xmark"></i> Treaty Expired</h3>
        <p>The <strong>${treaty.kind}</strong> treaty between <strong>${doc.name}</strong> and
        <strong>${partner?.name || 'an unknown nation'}</strong> has expired.</p>`,
      whisper: gmWhisper(),
    }).catch(() => {});
    Hooks.callAll('Pf2eNationsAndCitiesMaker.treatyExpired', {
      nationId: doc.id, partnerNationId: treaty.partnerNationId, treatyId: treaty.id,
    });
  }
}

/**
 * Declare war (#88): sets both nations' relations to hostile, posts a chat
 * card citing the claims the declaring nation invoked as casus belli, and
 * fires `warDeclared` for other modules/macros to react to.
 */
export async function declareWar(doc, targetDoc, claimIds = []) {
  if (!doc || !targetDoc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  const targetS = foundry.utils.deepClone(getSettlement(targetDoc) || {});

  const claims = (s.claims || []).filter(c => claimIds.includes(c.id));

  setRelationEntry(s, targetDoc.id, 'hostile', -80);
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);

  setRelationEntry(targetS, doc.id, 'hostile', -80);
  await targetDoc.setFlag(FLAG_SCOPE, FLAG_KEY, targetS);

  const citing = claims.length
    ? claims.map(c => `${c.kind}${c.notes ? ` (${c.notes})` : ''}`).join(', ')
    : 'no formal claims';
  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-flag-checkered"></i> War Declared</h3>
      <p><strong>${doc.name}</strong> declares war on <strong>${targetDoc.name}</strong>, citing ${citing}.</p>`,
    whisper: gmWhisper(),
  }).catch(() => {});
  Hooks.callAll('Pf2eNationsAndCitiesMaker.warDeclared', {
    nationId: doc.id, targetNationId: targetDoc.id, claimIds: claims.map(c => c.id),
  });
}

/**
 * Negotiate peace (#89): claims not chosen to be kept are given up, a
 * no-expiry non-aggression treaty is signed, and both nations' relations
 * reset to neutral.
 */
export async function negotiatePeace(doc, targetDoc, keepClaimIds = []) {
  if (!doc || !targetDoc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  const targetS = foundry.utils.deepClone(getSettlement(targetDoc) || {});

  const allClaims = s.claims || [];
  const kept = allClaims.filter(c => keepClaimIds.includes(c.id));
  const givenUp = allClaims.filter(c => !keepClaimIds.includes(c.id));
  s.claims = kept;

  setRelationEntry(s, targetDoc.id, 'neutral', 0);
  s.treaties = s.treaties || [];
  s.treaties.push({
    id: `treaty-${Math.random().toString(36).slice(2, 10)}`,
    partnerNationId: targetDoc.id, kind: 'non-aggression', signedOn: null, expiresOn: null,
    terms: 'Peace treaty ending hostilities.',
  });
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);

  setRelationEntry(targetS, doc.id, 'neutral', 0);
  await targetDoc.setFlag(FLAG_SCOPE, FLAG_KEY, targetS);

  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-dove"></i> Peace Negotiated</h3>
      <p><strong>${doc.name}</strong> and <strong>${targetDoc.name}</strong> sign a peace treaty.
      ${givenUp.length ? `${doc.name} relinquishes ${givenUp.length} claim(s) as part of the terms.` : ''}</p>`,
    whisper: gmWhisper(),
  }).catch(() => {});
  Hooks.callAll('Pf2eNationsAndCitiesMaker.peaceNegotiated', {
    nationId: doc.id, targetNationId: targetDoc.id, claimsGivenUp: givenUp.map(c => c.id),
  });

  return { kept, givenUp };
}

/**
 * Succession (#91): on a `rulerDied` calendar event, the nation's heir
 * ascends to head of state and the heir slot is cleared. A no-op if the
 * nation has no heir set.
 */
export async function applySuccession(doc) {
  if (!doc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  if (s.kind !== 'nation') return;
  const heir = s.heir || {};
  if (!heir.name && !heir.actorId) return;

  const oldLeaderName = s.government?.leaderName || 'The ruler';
  s.government = s.government || {};
  s.government.leaderName = heir.name || 'New Ruler';
  s.government.leaderActorId = heir.actorId || null;
  s.heir = { actorId: null, name: '' };
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);

  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-crown"></i> Succession in ${doc.name}</h3>
      <p><strong>${oldLeaderName}</strong> has died. <strong>${s.government.leaderName}</strong> ascends to power.</p>`,
    whisper: gmWhisper(),
  }).catch(() => {});
  Hooks.callAll('Pf2eNationsAndCitiesMaker.rulerSucceeded', {
    nationId: doc.id, newLeaderName: s.government.leaderName, newLeaderActorId: s.government.leaderActorId,
  });
}
