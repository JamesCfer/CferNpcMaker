/** Text-only layout planning. This prompt is inlined into the workflow, never the module. */
export const PLAN_PROMPT = `Expand the GM's map description into a drawable battle map layout for an image model. Image models cannot follow numbers or measurements, but they follow vivid physical detail, so describe what is there instead of how big it is. Return JSON only, with exactly this shape: {"areas":[{"what":"a cramped kitchen with a soot-blackened stone hearth, stacked ale barrels, hanging pots and a scarred work table","where":"north-west","size":"small, tucked into the corner"},{"what":"a busy common room with long trestle tables, benches, a polished bar along one wall and a warm fireplace","where":"centre","size":"the largest space, filling most of the map"}],"connections":["a narrow doorway from the kitchen into the common room behind the bar"],"notes":"one sentence on overall shape"}. Never use numbers, digits, measurements, square counts, feet, metres, or dimensions anywhere. Give size only in plain words compared to the rest of the map, such as tiny, small, narrow, long, wide, large, or the largest space. Areas have NO names. Describe each area only by what is physically in it: floor material, furniture, objects, clutter, light sources, damage, and how it looks; aim for 4 to 8 concrete details per area. Respect any number of rooms or areas the GM states exactly: a two-room tavern has exactly two rooms. List every room or area as its own entry, including the main or largest space (for example, the common room of a tavern); never return fewer than 2 areas. For small or simple places use 2 to 4 areas; larger places have at most 8. Areas may be irregular or natural (caves, shorelines, forests, rubble) and do not have to be rectangles or fill a 3x3 grid. Use approximate positions such as north-west, centre, or south; positions may repeat when areas share a region. Keep the perspective, genre, and setting. Keep the character of the place: a cave keeps rough, irregular natural rock walls and uneven floor, never straight built walls or square rooms, even when the GM says the walls are stone. When the perspective is top-down, describe every feature as it looks from straight above, because height does not show from overhead: a spiral staircase is a large ring of wedge-shaped stone steps curling around a central stone column; an obelisk is a square dark stone block with a pointed tip and a long shadow across the floor; a tower is a round roof; a statue is its head and shoulders seen from above. Do not invent pits, chasms, drops, lower tiers, or extra stairs the GM did not ask for, and give each GM-named feature one clear spot: one requested staircase means exactly one staircase. Always include every explicit GM request verbatim in the relevant area's what description, including location requests such as fire in the centre. Describe practical connections between areas by what they look like (an iron-banded door, a rope bridge, a gap in the rubble). Include nothing that needs text or labels. Do not include commentary or markdown.`;

export const PLAN_POSITIONS = ['north-west', 'north', 'north-east', 'west', 'centre', 'east', 'south-west', 'south', 'south-east'];

/** Remove measurements: image models ignore them, and they can turn into drawn labels. */
export function stripMeasurements(text) {
  return String(text ?? '')
    .replace(/\(?\s*(?:about|roughly|approximately|around|~)?\s*\d+(?:\.\d+)?\s*(?:x|by|×)\s*\d+(?:\.\d+)?\s*(?:squares?|ft|feet|foot|m|metres?|meters?|tiles?|cells?)?\s*\)?/gi, ' ')
    .replace(/\(?\s*(?:about|roughly|approximately|around|~)?\s*\d+(?:\.\d+)?\s*-?\s*(?:squares?|ft|feet|foot|m|metres?|meters?|tiles?|cells?)\b(?:\s*(?:wide|long|across|deep|tall|high))?\s*\)?/gi, ' ')
    .replace(/\s{2,}/g, ' ').replace(/\s+([,.;])/g, '$1').replace(/^[,;\s]+|[,;\s]+$/g, '').trim();
}

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
    typeof area.what === 'string' && stripMeasurements(area.what))
    .slice(0, 8).map(area => ({ what: stripMeasurements(area.what),
      ...(normaliseWhere(area.where) ? { where: normaliseWhere(area.where) } : {}),
      ...(typeof area.size === 'string' && stripMeasurements(area.size) ? { size: stripMeasurements(area.size) } : {}) }));
  if (!areas.length) throw new Error('A plan needs at least one valid area.');
  return { areas, rawAreas: value.areas.length, connections: Array.isArray(value.connections) ? value.connections.filter(v => typeof v === 'string').map(stripMeasurements).filter(Boolean).slice(0, 16) : [], notes: typeof value.notes === 'string' ? stripMeasurements(value.notes) : '', ...(areas.length < 2 ? { warning: 'The layout plan was thin; check it before generating.' } : {}) };
}

export function parsePlan(content) {
  const raw = String(content ?? '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  return validatePlan(JSON.parse(raw));
}

/** Keep the GM's words and requested areas while compacting optional detail. */
export function planText(plan, gmDescription = '') {
  const request = String(gmDescription).trim().replace(/[.\s]+$/, '');
  const areas = plan.areas.map(a => ({ ...a, details: a.what.split(',').map(x => x.trim()).filter(Boolean) }));
  const connections = [...plan.connections];
  let notes = plan.notes;
  const ending = 'Do not write any room names or labels on the map.';
  const make = () => [request && `GM request: ${request}.`,
    `Layout: ${areas.map(a => `${a.details.join(', ')}${a.size ? ` (${a.size})` : ''}${a.where ? ` at ${a.where}` : ''}`).join('; ')}.`,
    connections.length && `Connections: ${connections.join('; ')}.`, notes, ending].filter(Boolean).join(' ');
  const words = () => make().split(/\s+/).length;
  while (words() > 180 && connections.length) connections.pop();
  if (words() > 180) notes = '';
  const requestWords = new Set(request.toLowerCase().match(/[a-z]{4,}/g) || []);
  const protectedArea = a => a.details.some(d => (d.toLowerCase().match(/[a-z]{4,}/g) || []).some(w => requestWords.has(w)));
  for (const area of [...areas].reverse()) {
    while (words() > 180 && area.details.length > 1 && !protectedArea(area)) area.details.pop();
  }
  return make();
}
