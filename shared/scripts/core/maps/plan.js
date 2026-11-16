/** Text-only layout planning. This prompt is inlined into the workflow, never the module. */
export const PLAN_PROMPT = `Expand the GM's map request into a drawable top-down battle map plan. Rewrite every request as polished visual description with spelling corrected, meaning and placement preserved; never paste GM wording into an area's what. Return JSON only with this shape: {"areas":[{"what":"a cramped kitchen with a soot-blackened hearth and stacked barrels","where":"north-west","size":"small beside the main room"}],"connections":[],"notes":"overall visible shape"}. Use no numbers, digits, measurements, dimensions, text, or labels. Give relative sizes in words. Areas have no names; describe visible floor, objects, clutter, light, and damage with concrete detail. Open natural places (cave, cavern, grotto, forest, clearing, beach, shore, swamp, field, desert, open ruins) are ONE continuous space with exactly one area unless the GM asks for rooms. Do not split a space for a door or staircase. Built interiors include every requested room, including the main or largest space; respect the requested room count. Use approximate positions, repeating them when needed, and describe connections only between separate areas. Preserve genre, setting, and perspective. In caves describe the perimeter as jagged natural rock edges and boulders around an uneven continuous floor; never say walls, brick, masonry, or blocks unless the GM specifies built, carved, masonry, or brick construction. Describe every feature as seen from straight above: a spiral stair is a ring of wedge-shaped steps around a central column; an obelisk is a small square top with four triangular faces meeting at a central point, like a pyramid seen from above, on a wider square base with a long shadow on the floor. Describe tall things only by their top outline and shadow on the floor, never their height or sides. Ban ceilings, stalactites, hanging things, wall faces, front-view arches, "descending into shadow", and anything visible only from the side or below. Invent no pits, chasms, tiers, levels, split floors, or extra stairs; one requested staircase means exactly one. Treat huge, giant, massive, grand, centrepiece, or "in the centre" features as dominant: mention each once and early, with relative size words. A double door is ONE doorway with two leaves: a huge closed double door shown from above as two thick wooden leaves meeting in the middle, lying in a wide gap in the north rock edge. Doors and gates are thin strips in the wall line seen from above; never describe a doorway's arch, frame front, or wall around it. Include every GM feature once at its requested place. No commentary or markdown.`;

export const PLAN_POSITIONS = ['north-west', 'north', 'north-east', 'west', 'centre', 'east', 'south-west', 'south', 'south-east'];

/** Remove measurements: image models ignore them, and they can turn into drawn labels. */
export function stripMeasurements(text) {
  return String(text ?? '')
    .replace(/\(?\s*(?:about|roughly|approximately|around|~)?\s*\d+(?:\.\d+)?\s*(?:x|by|×)\s*\d+(?:\.\d+)?\s*(?:squares?|ft|feet|foot|m|metres?|meters?|tiles?|cells?)?\s*\)?/gi, ' ')
    .replace(/\(?\s*(?:about|roughly|approximately|around|~)?\s*\d+(?:\.\d+)?\s*-?\s*(?:squares?|ft|feet|foot|m|metres?|meters?|tiles?|cells?)\b(?:\s*(?:wide|long|across|deep|tall|high))?\s*\)?/gi, ' ')
    .replace(/\b\d+(?:\.\d+)?\b/g, ' ')
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

/** Turn a planned layout into compact visual prose without repeating the raw GM request. */
export function planText(plan, gmDescription = '') {
  const request = stripMeasurements(gmDescription).replace(/[.\s]+$/, '');
  const openPlace = /\b(cave|cavern|grotto|forest|clearing|beach|shore|swamp|field|desert|open ruins|ruins in the open)s?\b/i.test(request);
  const roomsRequested = /\b(?:rooms?|chambers?)\b/i.test(request) && !/\bno (?:rooms?|chambers?)\b/i.test(request);
  const continuous = openPlace && !roomsRequested;
  const clean = value => stripMeasurements(value).replace(/\boblisk\b/gi, 'obelisk')
    .replace(/\b(?:a |the )?(?:sunken cavern tier|vaulted stone threshold)\s*(?:where\s*)?/gi, '')
    .replace(/\bto the (north|south|east|west) there is a huge double door that is shut\b/gi, 'a single huge closed double door, far wider than a normal door, against the $1 wall')
    .replace(/\bto the (north|south|east|west) there is a spiral stone staircase that is also huge\b/gi, 'a huge spiral stone staircase, a ring of wedge-shaped stone steps around a central column at the $1 edge')
    .replace(/\bwhere to the (north|south|east|west) there is\b/gi, 'at the $1 edge,')
    .replace(/\ba obelisk\b/gi, 'an obelisk')
    .replace(/\s{2,}/g, ' ').replace(/^[,;\s]+|[,;\s]+$/g, '').trim();
  const sourceAreas = continuous ? [{ what: plan.areas.map(a => a.what).join(', '), where: 'centre', size: 'one continuous space with irregular natural edges' }] : plan.areas;
  const areas = sourceAreas.map(a => ({ ...a, details: clean(a.what).split(',').map(x => x.trim()).filter(Boolean) }));
  const connections = continuous ? [] : [...plan.connections].map(clean).filter(Boolean);
  let notes = continuous ? 'A single continuous space with natural irregular edges and an unbroken floor.' : clean(plan.notes);
  const ending = 'Do not write any room names or labels on the map.';
  const lead = request ? `GM request: ${request}. ` : '';
  const make = () => lead + [
    `Layout: ${areas.map(a => `${a.details.join(', ')}${a.size ? ` (${clean(a.size)})` : ''}${a.where ? ` at ${a.where}` : ''}`).join('; ')}.`,
    connections.length && `Connections: ${connections.join('; ')}.`, notes, ending].filter(Boolean).join(' ');
  const words = () => make().split(/\s+/).length;
  while (words() > 180 && connections.length) connections.pop();
  if (words() > 180) notes = '';
  for (const area of [...areas].reverse()) {
    while (words() > 180 && area.details.length > 1) area.details.pop();
  }
  return make();
}
