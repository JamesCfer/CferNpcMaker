/**
 * PF2e nations, cities and towns → a folder of handouts: the settlement itself
 * and, for a nation, one handout per child city.
 *
 * Foundry shows these on custom sheets with live economy and army tools; Roll20
 * has no equivalent, so we lay the generated content out as readable handouts.
 */
import { makePlan } from '../plan.js';
import { escapeHtml, titleCase, num } from '../util.js';

const SKIP = /^(_|id$|.*Id$|.*Ids$|.*ActorId$|sceneId$|bannerImage$|borderColor$|borderDrawingId$|ai$|gmOnly$|treasuryHistory$|schemaVersion)/;
const ORDER = ['description', 'summary', 'population', 'government', 'leadership', 'districts', 'demographics', 'religions', 'factions',
  'economy', 'production', 'tradeRoutes', 'stores', 'military', 'stats', 'relations', 'treaties', 'claims', 'hooks', 'rumors', 'notes'];

const empty = v => v === null || v === undefined || v === '' || (Array.isArray(v) && !v.length) || (typeof v === 'object' && !Array.isArray(v) && !Object.keys(v).filter(k => !SKIP.test(k) && !empty(v[k])).length);
const label = k => titleCase(k.replace(/([a-z])([A-Z])/g, '$1 $2'));

function inline(o) {
  const name = o.name || o.title || o.label;
  const rest = Object.entries(o).filter(([k, v]) => !SKIP.test(k) && !['name', 'title', 'label'].includes(k) && !empty(v));
  const text = rest.map(([k, v]) => (typeof v === 'object' ? `${label(k)}: ${Array.isArray(v) ? v.map(x => (typeof x === 'object' ? x.name || JSON.stringify(x) : x)).join(', ') : Object.entries(v).filter(([kk, vv]) => !SKIP.test(kk) && !empty(vv)).map(([kk, vv]) => `${label(kk)} ${vv}`).join(', ')}` : ['description', 'descriptor', 'notes', 'summary'].includes(k) ? String(v) : `${label(k)}: ${v}`)).join('. ');
  return `${name ? `<strong>${escapeHtml(name)}</strong>${text ? '. ' : ''}` : ''}${escapeHtml(text)}`;
}

function render(value) {
  if (Array.isArray(value)) return `<ul>${value.filter(v => !empty(v)).map(v => `<li>${typeof v === 'object' ? inline(v) : escapeHtml(v)}</li>`).join('')}</ul>`;
  if (value && typeof value === 'object') return `<ul>${Object.entries(value).filter(([k, v]) => !SKIP.test(k) && !empty(v)).map(([k, v]) => `<li><strong>${escapeHtml(label(k))}:</strong> ${typeof v === 'object' ? (Array.isArray(v) ? v.map(x => (typeof x === 'object' ? inline(x) : escapeHtml(x))).join('; ') : inline(v)) : escapeHtml(typeof v === 'number' ? v.toLocaleString('en-US') : v)}</li>`).join('')}</ul>`;
  if (typeof value === 'number') return `<p>${value.toLocaleString('en-US')}</p>`;
  return `<p>${escapeHtml(value).replace(/\n+/g, '</p><p>')}</p>`;
}

export function settlementHtml(s, form = {}) {
  const kind = s.kind || form.kind || 'settlement';
  const head = [titleCase(kind), s.size || form.size ? titleCase(s.size || form.size) : '', s.biome || form.biome || '', s.population ? `population ${num(s.population).toLocaleString('en-US')}` : '']
    .filter(Boolean).map(escapeHtml).join(' · ');
  const keys = [...ORDER, ...Object.keys(s).filter(k => !ORDER.includes(k))]
    .filter(k => k in s && !SKIP.test(k) && !['name', 'kind', 'size', 'biome', 'population', 'childCities'].includes(k) && !empty(s[k]));
  return `<p><strong>${head}</strong></p>` + keys.map(k => (k === 'description' || k === 'summary' ? render(s[k]) : `<h3>${escapeHtml(label(k))}</h3>${render(s[k])}`)).join('');
}

export function mapSettlement(response, form = {}) {
  const nation = response.nation || null;
  const main = nation || response.settlement || response.foundrySettlement || response;
  const name = main.name || form.name || 'Settlement';
  const cities = Array.isArray(response.childCities) ? response.childCities : [];
  const handouts = [{ name, notes: settlementHtml({ ...main, kind: nation ? 'nation' : main.kind }, form) + (cities.length ? `<h3>Cities</h3><ul>${cities.map(c => `<li>${escapeHtml(c.name || 'City')}</li>`).join('')}</ul>` : '') }];
  for (const c of cities) handouts.push({ name: c.name || 'City', notes: settlementHtml(c, { biome: form.biome }) });
  return makePlan('handouts', 'pf2e-settlement', { summary: cities.length ? `${name} and ${cities.length} cities` : name, folder: name, handouts });
}
