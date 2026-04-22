/**
 * Hero System 6e — actor data sanitization helpers.
 *
 * IMPORTANT: n8n Step 4 injects critical fields onto each item.system before
 * returning the response (INPUT, OPTIONID, OPTION, OPTION_ALIAS, CHARACTERISTIC,
 * ADDER, xmlTag, is5e). This function preserves those fields.
 *
 * CHARACTERISTIC STORAGE — two paths must both be populated:
 *   system.characteristics[KEY]  — uppercase key, read at runtime.
 *   system[KEY]                  — EmbeddedDataField HeroItemCharacteristic.
 *
 * hero6efoundryvttv2's _preCreate iterates all uppercase system keys and, when
 * XMLID is missing, replaces the embedded object with a blank version that
 * wipes LEVELS back to 0. We write { LEVELS, XMLID, xmlTag } so the guard
 * is satisfied and _preCreate skips it.
 */

export function sanitizeActorDataHero6e(actorData) {
  const generateId = () => foundry.utils.randomID(16);

  if (!actorData._id || actorData._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(actorData._id)) {
    actorData._id = generateId();
  }

  if (actorData.type !== 'npc') actorData.type = 'npc';

  if (!actorData.img)     actorData.img     = 'icons/svg/mystery-man.svg';
  if (!actorData.flags)   actorData.flags   = {};
  if (!actorData.effects) actorData.effects = [];

  if (!actorData.system) actorData.system = {};
  const sys = actorData.system;

  if (typeof sys.is5e === 'undefined') sys.is5e = false;

  const CHAR_BASES = {
    STR:10, DEX:10, CON:10, INT:10, EGO:10, PRE:10,
    OCV:3,  DCV:3,  OMCV:3, DMCV:3,
    SPD:2,  PD:2,   ED:2,
    REC:4,  END:20, BODY:10, STUN:20,
  };

  if (!sys.characteristics) sys.characteristics = {};
  const chars = sys.characteristics;

  for (const [key, base] of Object.entries(CHAR_BASES)) {
    const lkey  = key.toLowerCase();
    const upper = chars[key];
    const lower = chars[lkey];

    let levels;
    if (upper && upper.LEVELS !== undefined) {
      levels = Math.max(0, parseInt(upper.LEVELS) || 0);
    } else if (lower && lower.max !== undefined) {
      levels = Math.max(0, (parseInt(lower.max) || base) - base);
    } else if (upper && upper.max !== undefined) {
      levels = Math.max(0, (parseInt(upper.max) || base) - base);
    } else if (sys[key] && sys[key].LEVELS !== undefined) {
      levels = Math.max(0, parseInt(sys[key].LEVELS) || 0);
    } else {
      levels = 0;
    }

    chars[key] = {
      LEVELS: levels,
      max:    base + levels,
      value:  base + levels,
    };

    if (lkey !== key && lower !== undefined) delete chars[lkey];
    sys[key] = { LEVELS: levels, XMLID: key, xmlTag: key };
  }

  for (const k of Object.keys(chars)) {
    if (!CHAR_BASES[k]) delete chars[k];
  }

  if (!Array.isArray(actorData.items)) actorData.items = [];

  const VALID_ITEM_TYPES = new Set([
    'power', 'skill', 'talent', 'complication', 'equipment',
    'perk', 'martialart', 'maneuver', 'characteristic',
  ]);

  actorData.items = actorData.items.map(item => {
    if (!item._id || item._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(item._id)) {
      item._id = generateId();
    }

    if (!VALID_ITEM_TYPES.has(item.type)) item.type = 'power';

    if (!item.system) item.system = {};
    const s = item.system;

    if (!s.XMLID) {
      const xmlidDefaults = {
        power:        'CUSTOMPOWER',
        skill:        'CUSTOMSKILL',
        talent:       'CUSTOMTALENT',
        complication: 'GENERICDISADVANTAGE',
      };
      s.XMLID = xmlidDefaults[item.type] || 'CUSTOMPOWER';
    }

    if (typeof s.LEVELS === 'number') s.LEVELS = String(s.LEVELS);
    if (!s.LEVELS) s.LEVELS = '1';

    if (!s.ALIAS)       s.ALIAS       = item.name || s.XMLID;
    if (!s.description) s.description = item.name || '';

    if (s.active_points !== undefined) s.active_points = parseInt(s.active_points) || 0;
    if (s.real_cost     !== undefined) s.real_cost     = parseInt(s.real_cost)     || 0;
    if (s.ENDCOST       !== undefined) s.ENDCOST       = parseInt(s.ENDCOST)       || 0;

    if (item.type === 'complication') {
      s.POINTS = parseInt(s.POINTS) || 10;
    }

    item.system = s;
    return item;
  });
}
