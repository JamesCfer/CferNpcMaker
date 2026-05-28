/**
 * Starter templates and randomization tables for the Settlement Builder.
 * Pure data + small helpers — no Foundry deps so this file is easy to unit test.
 */

export const SETTLEMENT_TEMPLATES = [
  {
    id: 'custom',
    label: '— Custom —',
    fields: {}, // selecting "custom" leaves the form untouched
  },
  {
    id: 'farming-village',
    label: 'Sleepy farming village',
    fields: {
      kind: 'village', size: 'village', biome: 'temperate plains',
      governmentHint: 'Elected village elder + farmer council',
      populationHint: 220,
      description: 'A quiet farming village surrounded by fields of wheat and barley. A single dirt road runs through the centre, lined by a chapel, a smithy, and the harvest granary. Folk are friendly but wary of outsiders, and the local militia is a handful of farmhands with hunting bows.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'river-trade-town',
    label: 'Bustling river-trade town',
    fields: {
      kind: 'town', size: 'town', biome: 'river valley',
      governmentHint: 'Mayor + merchant guild',
      populationHint: 3200,
      description: 'A busy trade town straddling a wide river. Barges unload spices, wool and iron at the docks daily. A merchant guild dominates politics, and a small town watch keeps the wharves safe at night.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'frontier-outpost',
    label: 'Fortified frontier outpost',
    fields: {
      kind: 'town', size: 'town', biome: 'borderlands',
      governmentHint: 'Garrison commander',
      populationHint: 800,
      description: 'A walled outpost guarding the edge of civilised lands. Most buildings are inside the palisade. A garrison of regular soldiers and a small group of scouts keep watch against monstrous incursions. Stores cater to soldiers and travellers.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'coastal-port',
    label: 'Coastal fishing port',
    fields: {
      kind: 'town', size: 'town', biome: 'coastal',
      governmentHint: 'Harbourmaster + fishers council',
      populationHint: 1800,
      description: 'A salty fishing port clustered around a stone quay. Boats are pulled up onto the shingle at low tide. Fish-mongers, salters, ropewalkers and a single tavern make up the bulk of trade. A small militia of dock workers handles trouble.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'desert-waystation',
    label: 'Desert caravan waystation',
    fields: {
      kind: 'village', size: 'village', biome: 'desert',
      governmentHint: 'Caravan-master council',
      populationHint: 350,
      description: 'A sun-bleached cluster of mudbrick buildings around a deep well. Caravans stop here to water their beasts, trade goods, and hire local guides. The town has a few veteran scouts that double as guards.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'dwarven-hold',
    label: 'Mountain dwarven hold',
    fields: {
      kind: 'city', size: 'city', biome: 'mountain',
      governmentHint: 'Clan thane + council of elders',
      populationHint: 6800,
      description: 'A dwarven hold carved into the side of a mountain, lit by mushroom-glow and forge-fire. Halls of crafters, weapon-smiths, and rune-scribes line the great central avenue. The hold guard is heavily armoured and disciplined.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'elven-enclave',
    label: 'Elven forest enclave',
    fields: {
      kind: 'town', size: 'town', biome: 'ancient forest',
      governmentHint: 'Council of speakers',
      populationHint: 1400,
      description: 'An elven settlement built into the canopy of an ancient forest, connected by rope bridges and elegant wooden walkways. The folk are quiet, observant, and protective of the woods. Rangers patrol the perimeter.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'capital-city',
    label: 'Walled capital city',
    fields: {
      kind: 'city', size: 'metropolis', biome: 'temperate hills',
      governmentHint: 'Monarchy + royal court',
      populationHint: 24000,
      description: 'A great walled capital with stone bastions, a citadel at its centre, and bustling markets in every quarter. Noble houses, churches, academies, and a sizable city watch all jockey for power. Trade flows in from every direction.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'mining-boomtown',
    label: 'Mining boomtown',
    fields: {
      kind: 'town', size: 'town', biome: 'mountain foothills',
      governmentHint: 'Mining company foreman',
      populationHint: 1600,
      description: 'A rough, half-built boomtown grown around a productive mine. Most buildings are wooden and recent. Saloons, prospectors, smiths and a company store dominate. The "law" is whatever the mining company says it is.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
  {
    id: 'temple-town',
    label: 'Pilgrim temple-town',
    fields: {
      kind: 'town', size: 'town', biome: 'sacred valley',
      governmentHint: 'High priest + temple council',
      populationHint: 1100,
      description: 'A town built around a great temple complex. Pilgrims arrive daily and locals make their living tending inns, blessing-tokens, and pilgrim supplies. Temple guards keep order, and the high priest holds the real authority.',
      includeStores: true, includeMilitary: true, includeLeadership: true,
    },
  },
];

export function getTemplate(id) {
  return SETTLEMENT_TEMPLATES.find(t => t.id === id) || SETTLEMENT_TEMPLATES[0];
}

const NAME_TABLES = {
  generic: {
    prefix: ['Ash', 'Old', 'New', 'Black', 'Stone', 'Iron', 'Silver', 'Green', 'Long', 'Red', 'White', 'High', 'Low', 'East', 'West', 'North', 'South'],
    root:   ['ford', 'wood', 'hollow', 'shire', 'haven', 'reach', 'crest', 'gate', 'march', 'fall', 'bridge', 'cross', 'mire', 'glen', 'hold'],
  },
  desert:  { prefix: ['Sun', 'Sand', 'Dune', 'Ash', 'Bone', 'Mirage', 'Salt'],    root: ['reach', 'crossing', 'well', 'oasis', 'rest', 'spire', 'gate'] },
  coastal: { prefix: ['Tide', 'Salt', 'Gull', 'Storm', 'Spray', 'Net', 'Wave'],    root: ['port', 'harbor', 'wharf', 'haven', 'bay', 'point', 'cove'] },
  mountain:{ prefix: ['Stone', 'Iron', 'Frost', 'Mount', 'Cliff', 'Hammer', 'Anvil'], root: ['hold', 'forge', 'peak', 'fall', 'gate', 'deep', 'reach'] },
  forest:  { prefix: ['Green', 'Moss', 'Fern', 'Oak', 'Elder', 'Vine', 'Thorn'], root: ['wood', 'glen', 'hollow', 'thorpe', 'grove', 'shade', 'ring'] },
  river:   { prefix: ['Riven', 'Bridge', 'Ford', 'Reed', 'Otter', 'Mill'],         root: ['ford', 'bridge', 'mill', 'crossing', 'bend', 'reach'] },
};

function pick(arr, rng) { return arr[Math.floor(rng() * arr.length)]; }

function rngFromSeed(seed) {
  let s = 0;
  for (let i = 0; i < seed.length; i++) s = (s * 31 + seed.charCodeAt(i)) >>> 0;
  return function() { s = (s * 1664525 + 1013904223) >>> 0; return s / 0xFFFFFFFF; };
}

function tableForBiome(biome) {
  const b = (biome || '').toLowerCase();
  if (b.includes('desert')) return NAME_TABLES.desert;
  if (b.includes('coast') || b.includes('sea') || b.includes('port')) return NAME_TABLES.coastal;
  if (b.includes('mountain')) return NAME_TABLES.mountain;
  if (b.includes('forest') || b.includes('wood')) return NAME_TABLES.forest;
  if (b.includes('river') || b.includes('valley')) return NAME_TABLES.river;
  return NAME_TABLES.generic;
}

export function randomName({ kind = 'town', biome = '', seed = '' } = {}) {
  const rng = rngFromSeed(`${kind}|${biome}|${seed}|${Date.now()}|${Math.random()}`);
  const tbl = tableForBiome(biome);
  return `${pick(tbl.prefix, rng)}${pick(tbl.root, rng)}`;
}

export function randomSettlement() {
  // Pick a non-custom template at random, then return its fields merged with
  // a freshly rolled name.
  const candidates = SETTLEMENT_TEMPLATES.filter(t => t.id !== 'custom');
  const t = candidates[Math.floor(Math.random() * candidates.length)];
  return {
    templateId: t.id,
    fields: {
      ...t.fields,
      name: randomName({ kind: t.fields.kind, biome: t.fields.biome }),
    },
  };
}
