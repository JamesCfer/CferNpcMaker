export const CURRENT_SCHEMA_VERSION = 2;

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
