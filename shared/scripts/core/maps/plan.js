/** Text-only layout planning. This prompt is inlined into the workflow, never the module. */
export const PLAN_PROMPT = `Expand the GM's map description into a drawable room-by-room or area-by-area battle map layout. Return JSON only, with exactly this shape: {"areas":[{"name":"...","where":"north-west","size":"about W by H squares","features":"short comma list"}],"connections":["gatehouse to courtyard through a broken portcullis"],"notes":"one sentence on overall shape"}. Use 4 to 10 areas that fit the requested square dimensions, perspective, genre, and setting. Fill the whole map with distinct areas; their positions must not overlap. Each where must be one of north-west, north, north-east, west, centre, east, south-west, south, south-east. Keep each feature list short and drawable from above. Include no feature that needs text or labels. Respect every explicit GM request: for example, fire in the centre means an area in the centre with fire. Describe practical connections between areas. Do not include commentary or markdown.`;

export const PLAN_POSITIONS = ['north-west', 'north', 'north-east', 'west', 'centre', 'east', 'south-west', 'south', 'south-east'];

/** Drop malformed areas and connections from model output. */
export function validatePlan(value) {
  if (!value || typeof value !== 'object' || !Array.isArray(value.areas)) throw new Error('Missing areas.');
  const used = new Set();
  const areas = value.areas.filter(area => {
    if (!area || typeof area !== 'object' || !['name', 'where', 'size', 'features'].every(key => typeof area[key] === 'string' && area[key].trim())) return false;
    if (!PLAN_POSITIONS.includes(area.where) || used.has(area.where)) return false;
    used.add(area.where);
    return true;
  }).slice(0, 10).map(area => Object.fromEntries(['name', 'where', 'size', 'features'].map(key => [key, area[key].trim()])));
  if (areas.length < 4) throw new Error('A plan needs at least four valid areas.');
  return { areas, connections: Array.isArray(value.connections) ? value.connections.filter(v => typeof v === 'string' && v.trim()).map(v => v.trim()).slice(0, 16) : [], notes: typeof value.notes === 'string' ? value.notes.trim() : '' };
}

export function parsePlan(content) {
  const raw = String(content ?? '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  return validatePlan(JSON.parse(raw));
}

/** Compact prose, preserving areas and connections before optional features. */
export function planText(plan) {
  const areas = plan.areas.map(a => ({ ...a, featureParts: a.features.split(',').map(x => x.trim()).filter(Boolean) }));
  const make = () => `Layout: ${areas.map(a => `the ${a.name} (${a.size}) sits at the ${a.where}${a.featureParts.length ? `, with ${a.featureParts.join(', ')}` : ''}`).join('; ')}. Connections: ${plan.connections.join('; ')}.${plan.notes ? ` ${plan.notes}` : ''}`.replace(/\.\./g, '.');
  let result = make();
  while (result.length > 900 && areas.some(a => a.featureParts.length)) {
    const area = [...areas].reverse().find(a => a.featureParts.length);
    area.featureParts.pop();
    result = make();
  }
  if (result.length > 900) result = result.slice(0, 897).trimEnd().replace(/[\s,;]+$/, '') + '...';
  return result;
}
