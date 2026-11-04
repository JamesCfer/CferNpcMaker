/** Text-only layout planning. This prompt is inlined into the workflow, never the module. */
export const PLAN_PROMPT = `Expand the GM's map description into a drawable battle map layout. Return JSON only, with exactly this shape: {"areas":[{"what":"a kitchen with a hearth, barrels and a work table","where":"north-west","size":"about 5 by 4 squares"},{"what":"a common room with tables and a bar","where":"centre","size":"about 8 by 6 squares"}],"connections":["a doorway between the kitchen and common room"],"notes":"one sentence on overall shape"}. Areas have NO names. Describe each area only by what is physically in it. Respect any number of rooms or areas the GM states exactly: a two-room tavern has exactly two rooms. List every room or area as its own entry, including the main or largest space (for example, the common room of a tavern); never return fewer than 2 areas. For small or simple places use 2 to 4 areas; larger places have at most 8. Areas may be irregular or natural (caves, shorelines, forests, rubble) and do not have to be rectangles or fill a 3x3 grid. Use approximate positions such as north-west, centre, or south; positions may repeat when areas share a region. Keep the requested square dimensions, perspective, genre, and setting. Always include every explicit GM request verbatim in the relevant area's what description, including location requests such as fire in the centre. Describe practical connections between areas. Include nothing that needs text or labels. Do not include commentary or markdown.`;

export const PLAN_POSITIONS = ['north-west', 'north', 'north-east', 'west', 'centre', 'east', 'south-west', 'south', 'south-east'];

/** Keep described areas while making common position wording consistent. */
export function validatePlan(value) {
  if (!value || typeof value !== 'object' || !Array.isArray(value.areas)) throw new Error('Missing areas.');
  const normaliseWhere = input => {
    const original = typeof input === 'string' ? input.trim().toLowerCase().slice(0, 30) : '';
    const words = original.replace(/\b(center|middle)\b/g, 'centre')
      .replace(/\b(front|forward)\b/g, 'south')
      .replace(/\b(rear|back)\b/g, 'north')
      .replace(/\bleft\b/g, 'west').replace(/\bright\b/g, 'east');
    if (PLAN_POSITIONS.includes(words)) return words;
    const compass = words.match(/\b(north|south|east|west)\b/g);
    if (compass?.length > 1) return compass[0];
    if (compass?.length) return compass[0];
    if (/\bcentre\b/.test(words)) return 'centre';
    return original;
  };
  const areas = value.areas.filter(area => area && typeof area === 'object' &&
    typeof area.what === 'string' && area.what.trim())
    .slice(0, 8).map(area => ({ what: area.what.trim(),
      ...(normaliseWhere(area.where) ? { where: normaliseWhere(area.where) } : {}),
      ...(typeof area.size === 'string' && area.size.trim() ? { size: area.size.trim() } : {}) }));
  if (!areas.length) throw new Error('A plan needs at least one valid area.');
  return { areas, rawAreas: value.areas.length, connections: Array.isArray(value.connections) ? value.connections.filter(v => typeof v === 'string' && v.trim()).map(v => v.trim()).slice(0, 16) : [], notes: typeof value.notes === 'string' ? value.notes.trim() : '', ...(areas.length < 2 ? { warning: 'The layout plan was thin; check it before generating.' } : {}) };
}

export function parsePlan(content) {
  const raw = String(content ?? '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  return validatePlan(JSON.parse(raw));
}

/** Keep the GM's words and requested areas while compacting optional detail. */
export function planText(plan, gmDescription = '') {
  const request = String(gmDescription).trim();
  const areas = plan.areas.map(a => ({ ...a, details: a.what.split(',').map(x => x.trim()).filter(Boolean) }));
  const connections = [...plan.connections];
  let notes = plan.notes;
  const ending = 'Do not write any room names or labels on the map.';
  const make = () => [request && `GM request: ${request}.`,
    `Layout: ${areas.map(a => `${a.details.join(', ')}${a.size ? ` (${a.size})` : ''}${a.where ? ` at ${a.where}` : ''}`).join('; ')}.`,
    connections.length && `Connections: ${connections.join('; ')}.`, notes, ending].filter(Boolean).join(' ');
  const words = () => make().split(/\s+/).length;
  while (words() > 120 && connections.length) connections.pop();
  if (words() > 120) notes = '';
  const requestWords = new Set(request.toLowerCase().match(/[a-z]{4,}/g) || []);
  const protectedArea = a => a.details.some(d => (d.toLowerCase().match(/[a-z]{4,}/g) || []).some(w => requestWords.has(w)));
  for (const area of [...areas].reverse()) {
    while (words() > 120 && area.details.length > 1 && !protectedArea(area)) area.details.pop();
  }
  return make();
}
