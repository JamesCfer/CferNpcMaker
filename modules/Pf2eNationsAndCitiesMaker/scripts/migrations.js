export const CURRENT_SCHEMA_VERSION = 7;

/** @type {Array<{ from: number, to: number, fn: (s: object) => object }>} */
const _migrations = [
  {
    from: 1, to: 2,
    fn(s) {
      const ranks = Array.isArray(s.military?.ranks) ? s.military.ranks.map(r => ({ dailyWage: 0, ...r })) : [];
      const stores = Array.isArray(s.stores) ? s.stores.map(st => ({
        closed: false,
        ...st,
        income: { daysInDebt: 0, ...st.income },
      })) : [];
      return { ...s, _schemaVersion: 2, military: { ...s.military, ranks }, stores };
    },
  },
  {
    from: 2, to: 3,
    fn(s) {
      const stores = Array.isArray(s.stores) ? s.stores.map(st => ({
        marketWeekday: null,
        isBlackMarket: false,
        ...st,
      })) : [];
      return { ...s, _schemaVersion: 3, stores, priceMultiplier: s.priceMultiplier ?? 1.0 };
    },
  },
  {
    from: 3, to: 4,
    fn(s) {
      const stores = Array.isArray(s.stores) ? s.stores.map(st => ({
        priceTier: 'standard',
        ...st,
      })) : [];
      return { ...s, _schemaVersion: 4, stores, religions: s.religions ?? [] };
    },
  },
  {
    from: 4, to: 5,
    fn(s) {
      return { ...s, _schemaVersion: 5, bannerImage: s.bannerImage ?? null };
    },
  },
  {
    from: 5, to: 6,
    fn(s) {
      const stores = Array.isArray(s.stores) ? s.stores.map(st => ({ jobs: [], ...st })) : [];
      return { ...s, _schemaVersion: 6, stores };
    },
  },
  {
    from: 6, to: 7,
    fn(s) {
      return {
        ...s, _schemaVersion: 7,
        treaties:         s.treaties ?? [],
        vassalNationIds:  s.vassalNationIds ?? [],
        suzerainNationId: s.suzerainNationId ?? null,
      };
    },
  },
];

export function migrateSettlement(s) {
  if (!s || typeof s !== 'object') return s;
  let data = s;
  let v = Number(data._schemaVersion) || 0;
  while (v < CURRENT_SCHEMA_VERSION) {
    const m = _migrations.find(m => m.from === v);
    if (!m) break;
    data = m.fn(data);
    v = m.to;
  }
  if ((Number(data._schemaVersion) || 0) < CURRENT_SCHEMA_VERSION) {
    data = { ...data, _schemaVersion: CURRENT_SCHEMA_VERSION };
  }
  return data;
}
