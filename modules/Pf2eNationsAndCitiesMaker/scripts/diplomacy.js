/**
 * Treaty lifecycle (#84, #85). Treaties live on a Nation document's own
 * `treaties` list. Expiry is checked on every calendar day-advance tick
 * rather than through Pf2eCalendarTimeline's own event scheduler, mirroring
 * how mercenary contract expiry already piggybacks on the same hook.
 */

import { FLAG_SCOPE, FLAG_KEY } from './constants.js';
import { cmpDate } from './army.js';

function gmWhisper() {
  return game.users?.filter(u => u.isGM).map(u => u.id) ?? [];
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
