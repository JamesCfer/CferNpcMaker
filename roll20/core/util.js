/**
 * Small helpers shared by the Roll20 mappers. No Roll20 or Foundry globals here,
 * so everything in core/ runs the same in the extension, the tests and Node.
 */

const ROW_CHARS = '-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz';
let lastTime = 0;
const lastRand = [];

/** Roll20's own row-id algorithm (the one every sheet worker uses), with '_' swapped out. */
export function rowId() {
  let now = Date.now();
  const same = now === lastTime;
  lastTime = now;
  const head = new Array(8);
  for (let i = 7; i >= 0; i--) { head[i] = ROW_CHARS.charAt(now % 64); now = Math.floor(now / 64); }
  let id = head.join('');
  if (same) {
    let i = 11;
    for (; i >= 0 && lastRand[i] === 63; i--) lastRand[i] = 0;
    lastRand[i]++;
  } else {
    for (let i = 0; i < 12; i++) lastRand[i] = Math.floor(64 * Math.random());
  }
  for (let i = 0; i < 12; i++) id += ROW_CHARS.charAt(lastRand[i]);
  return id.replace(/_/g, 'Z');
}

export function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/**
 * Foundry description HTML to plain text: enrichers like @Damage[2d6[fire]],
 * @UUID[...]{Label}, [[/r 1d6]] and @Check[dc:15|type:reflex] become words.
 */
export function plainText(html) {
  let s = String(html ?? '');
  s = s.replace(/@Damage\[([^\]]*\])*[^\]]*\]/g, m => {
    const inner = m.slice(8, -1);
    return inner.replace(/\(([^)]*)\)/g, '$1').replace(/(\d+d\d+(?:\s*[+-]\s*\d+)?)\[([a-z, ]+)\]/gi, '$1 $2').replace(/\[|\]/g, ' ').replace(/\|.*$/, '').replace(/\s+/g, ' ').trim();
  });
  s = s.replace(/@Check\[([^\]]*)\]/g, (_, body) => {
    const parts = Object.fromEntries(body.split('|').map(p => p.includes(':') ? p.split(':') : ['type', p]));
    const type = parts.type ? parts.type.replace(/^\w/, c => c.toUpperCase()) : 'Check';
    return parts.dc ? `DC ${parts.dc} ${type}` : type;
  });
  s = s.replace(/@Template\[([^\]]*)\]/g, (_, body) => {
    const parts = Object.fromEntries(body.split('|').map(p => p.split(':')));
    return [parts.distance && `${parts.distance}-foot`, parts.type].filter(Boolean).join(' ');
  });
  s = s.replace(/@\w+\[[^\]]*\]\{([^}]*)\}/g, '$1');
  s = s.replace(/@UUID\[[^\]]*?\.([^.\]]+)\]/g, (_, last) => last.replace(/-/g, ' '));
  s = s.replace(/\[\[\/\w+\s+([^\]]*?)\]\](\{([^}]*)\})?/g, (_, roll, __, label) => label || roll.replace(/#.*$/, '').trim());
  s = s.replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li|h\d)>/gi, '\n').replace(/<li[^>]*>/gi, '- ');
  s = s.replace(/<hr[^>]*>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  return s.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

export const signed = n => (Number(n) >= 0 ? `+${Number(n)}` : `${Number(n)}`);
export const titleCase = s => String(s ?? '').replace(/[-_]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
export const num = (v, d = 0) => (Number.isFinite(Number(v)) && v !== '' && v !== null ? Number(v) : d);
export const list = v => (Array.isArray(v) ? v : v instanceof Set ? [...v] : v && typeof v === 'object' ? Object.values(v) : []);
