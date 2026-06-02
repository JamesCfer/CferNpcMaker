export const CURRENT_SCHEMA_VERSION = 1;

/** @type {Array<{ from: number, to: number, fn: (s: object) => object }>} */
const _migrations = [];

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
