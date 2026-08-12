/**
 * Sanitizer for AI-returned store payloads. Coerces missing or malformed
 * fields into safe defaults so the store sheet never sees holes — a partial
 * response still renders a usable shop.
 *
 * The shape is deliberately compatible with the `Store` typedef in
 * Pf2eNationsAndCitiesMaker (modules/Pf2eNationsAndCitiesMaker/scripts/types.js),
 * so a store generated here can be pasted into a settlement's `stores` array.
 */

/** Store types offered on the form, shared with the settlement module. */
export const STORE_TYPES = [
  'blacksmith', 'armorer', 'weapons', 'alchemist',
  'tavern', 'inn', 'general', 'grocer',
  'tailor', 'jeweler', 'temple', 'magic',
  'stable', 'apothecary', 'bookbinder', 'guild', 'other',
];

export const PRICE_TIERS = ['low', 'standard', 'high', 'luxury'];

/** Multiplier applied to displayed prices per wealth tier (mirrors #42). */
export const PRICE_TIER_MULS = { low: 0.75, standard: 1.0, high: 1.5, luxury: 2.0 };

const SHIFTS   = new Set(['morning', 'day', 'evening', 'night', 'graveyard']);
const RARITIES = new Set(['common', 'uncommon', 'rare', 'very rare', 'legendary', 'artifact']);

export function storeTypeLabel(type) {
  if (!type) return 'Other';
  return type.charAt(0).toUpperCase() + type.slice(1);
}

function safeNum(n, def, min = -Infinity, max = Infinity) {
  const v = Number(n);
  if (!Number.isFinite(v)) return def;
  return Math.min(max, Math.max(min, v));
}

function safeString(s, def = '') {
  return (typeof s === 'string' && s.trim()) ? s.trim() : def;
}

function shortId(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

/** 'HH:MM', falling back to the supplied default when unparseable. */
function safeTime(value, def) {
  const s = safeString(value, '');
  return /^([01]?\d|2[0-3]):[0-5]\d$/.test(s) ? s.padStart(5, '0') : def;
}

/**
 * @param {object} raw         Payload from the relay (or a stub).
 * @param {object} [formData]  The form values, used to fill gaps.
 * @returns {object} A complete store object.
 */
export function sanitizeStore(raw, formData = {}) {
  const s = (raw && typeof raw === 'object') ? raw : {};
  const src = (s.store && typeof s.store === 'object') ? s.store : s;

  const type = STORE_TYPES.includes(src.type) ? src.type : (formData.storeType || 'general');
  const owner = (src.owner && typeof src.owner === 'object') ? src.owner : {};
  const hours = (src.hours && typeof src.hours === 'object') ? src.hours : {};
  const income = (src.income && typeof src.income === 'object') ? src.income : {};

  return {
    id:          safeString(src.id, shortId('store')),
    name:        safeString(src.name, formData.name || 'New Shop'),
    type,
    description: safeString(src.description, ''),
    settlement:  safeString(src.settlement, formData.settlement || ''),
    priceTier:   PRICE_TIERS.includes(src.priceTier)
      ? src.priceTier
      : (PRICE_TIERS.includes(formData.priceTier) ? formData.priceTier : 'standard'),
    level:       safeNum(src.level, formData.level ?? 1, 0, 30),
    closed:        !!src.closed,
    isBlackMarket: !!src.isBlackMarket,
    marketWeekday: Number.isInteger(src.marketWeekday) ? Math.min(6, Math.max(0, src.marketWeekday)) : null,

    owner: {
      name:        safeString(owner.name, 'Unknown Owner'),
      description: safeString(owner.description, ''),
      actorId:     owner.actorId || null,
    },

    staff: Array.isArray(src.staff) ? src.staff.slice(0, 20).map(m => ({
      id:      safeString(m?.id, shortId('staff')),
      name:    safeString(m?.name, 'Unnamed'),
      role:    safeString(m?.role, 'Assistant'),
      shift:   SHIFTS.has(m?.shift) ? m.shift : 'day',
      actorId: m?.actorId || null,
    })) : [],

    hours: {
      open:  safeTime(hours.open,  '08:00'),
      close: safeTime(hours.close, '20:00'),
      daysClosed: Array.isArray(hours.daysClosed)
        ? hours.daysClosed.map(d => safeString(d)).filter(Boolean).slice(0, 7)
        : [],
    },

    inventory: Array.isArray(src.inventory) ? src.inventory.slice(0, 60).map(i => ({
      id:          safeString(i?.id, shortId('inv')),
      name:        safeString(i?.name, 'Unnamed Good'),
      description: safeString(i?.description, ''),
      price:       safeNum(i?.price, 1, 0, 9_999_999),
      stock:       safeNum(i?.stock, 1, 0, 9999),
      level:       safeNum(i?.level, 0, 0, 30),
      rarity:      RARITIES.has(i?.rarity) ? i.rarity : 'common',
      itemId:      i?.itemId || null,
    })) : [],

    services: Array.isArray(src.services) ? src.services.slice(0, 20).map(v => ({
      id:          safeString(v?.id, shortId('svc')),
      name:        safeString(v?.name, 'Service'),
      description: safeString(v?.description, ''),
      price:       safeNum(v?.price, 0, 0, 9_999_999),
    })) : [],

    rumors: Array.isArray(src.rumors)
      ? src.rumors.map(r => safeString(r)).filter(Boolean).slice(0, 10)
      : [],

    income: {
      balance:    safeNum(income.balance,  50, -9_999_999, 9_999_999),
      dailyAvg:   safeNum(income.dailyAvg,  6, 0, 9_999_999),
      lastTick:   safeNum(income.lastTick,  0, 0, Number.MAX_SAFE_INTEGER),
      daysInDebt: safeNum(income.daysInDebt, 0, 0, 99999),
    },

    ai: {
      prompt:      safeString(formData.description, safeString(src.ai?.prompt, '')),
      generatedAt: safeNum(src.ai?.generatedAt, Date.now(), 0, Number.MAX_SAFE_INTEGER),
    },
  };
}

/**
 * Price after the wealth-tier multiplier, rounded to a whole coin (minimum 1
 * for anything that wasn't free to begin with).
 *
 * @param {number} price
 * @param {string} priceTier
 * @returns {number}
 */
export function tieredPrice(price, priceTier) {
  const base = Number(price) || 0;
  const mul  = PRICE_TIER_MULS[priceTier] ?? 1;
  const out  = Math.round(base * mul);
  return base > 0 ? Math.max(1, out) : 0;
}

/**
 * Total gp value of everything on the shelves, at tier prices.
 * @param {object} store
 * @returns {number}
 */
export function inventoryValue(store) {
  return (store?.inventory || []).reduce(
    (sum, i) => sum + tieredPrice(i.price, store.priceTier) * (Number(i.stock) || 0),
    0,
  );
}
