/**
 * Import plans: the one format both Roll20 writers understand.
 *
 * A mapper turns a generator's JSON into a plan. The extension then applies it
 * either in the page (no Pro needed) or by sending it to the companion Mod
 * script in chat (Pro games). Both writers follow the same steps:
 *
 *   1. create the character (or handouts / map objects)
 *   2. write `attributes` (always)
 *   3. write `drop.attributes`, then set `drop.trigger` so the sheet's own
 *      compendium-drop worker builds the statblock the way a compendium drag does
 *   4. after `drop.timeoutMs`, if `drop.verify` is still empty the drop did not
 *      run, so write `fallback.attributes` instead
 *
 * @typedef {object} ImportPlan
 * @property {1} v
 * @property {'character'|'handouts'|'map'} kind
 * @property {string} generator           Catalog id that made it.
 * @property {string} summary             One line for chat.
 * @property {{name:string,bio?:string,gmnotes?:string,avatar?:string}} [character]
 * @property {Record<string, string|number|{current:any,max?:any}>} [attributes]
 * @property {{attributes:Record<string,any>, trigger:{name:string,value:string}, verify:string, timeoutMs:number}} [drop]
 * @property {{attributes:Record<string,any>}} [fallback]
 * @property {{name:string,notes:string,gmnotes?:string,avatar?:string}[]} [handouts]
 * @property {object} [map]
 */

export const PLAN_VERSION = 1;
export const CHAT_COMMAND = '!cfer-import';
export const CHUNK_SIZE = 3000;

export function makePlan(kind, generator, fields) {
  return { v: PLAN_VERSION, kind, generator, summary: '', ...fields };
}

/* ── chat transport: base64url(UTF-8 JSON) split into numbered chunks ── */

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';

export function toBase64Url(str) {
  const bytes = new TextEncoder().encode(str);
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] ?? 0) << 8) | (bytes[i + 2] ?? 0);
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63];
    out += i + 1 < bytes.length ? B64[(n >> 6) & 63] : '';
    out += i + 2 < bytes.length ? B64[n & 63] : '';
  }
  return out;
}

export function fromBase64Url(s) {
  const bytes = [];
  let buf = 0, bits = 0;
  for (const ch of s) {
    const v = B64.indexOf(ch);
    if (v < 0) continue;
    buf = (buf << 6) | v; bits += 6;
    if (bits >= 8) { bits -= 8; bytes.push((buf >> bits) & 255); }
  }
  return new TextDecoder().decode(new Uint8Array(bytes));
}

/** Chat lines that carry a plan to the Mod script. */
export function planToChatCommands(plan, id = Math.random().toString(36).slice(2, 10)) {
  const data = toBase64Url(JSON.stringify(plan));
  const parts = [];
  for (let i = 0; i < data.length; i += CHUNK_SIZE) parts.push(data.slice(i, i + CHUNK_SIZE));
  return parts.map((p, i) => `${CHAT_COMMAND} ${id} ${i + 1}/${parts.length} ${p}`);
}

/** Every attribute the plan writes, in write order (for in-page writers and tests). */
export function planAttributeNames(plan) {
  return [
    ...Object.keys(plan.attributes || {}),
    ...Object.keys(plan.drop?.attributes || {}),
    ...(plan.drop ? [plan.drop.trigger.name] : []),
  ];
}
