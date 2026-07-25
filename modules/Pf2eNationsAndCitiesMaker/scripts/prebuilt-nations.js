/**
 * Compendium of 5 prebuilt nations (#112). Ready-to-drop nation journals
 * for users without AI credits — partial data, sanitizeSettlement()
 * backfills the rest. childCityIds starts empty; link cities from the
 * Nation sheet after import.
 */

export const PREBUILT_NATIONS = [
  {
    id: 'kingdom-of-duvane',
    label: 'Kingdom of Duvane (feudal monarchy)',
    data: {
      kind: 'nation', size: 'metropolis', population: 420000, biome: 'temperate hills',
      government: { type: 'Hereditary monarchy', leaderName: 'Queen Meris Duvane' },
      leadership: [
        { title: 'Queen', name: 'Meris Duvane', role: 'ruler', actorId: null },
        { title: 'Lord Chancellor', name: 'Aldous Ferrin', role: 'administration', actorId: null },
      ],
      treasury: { cp: 0, sp: 0, gp: 92000, pp: 400 },
      childCityIds: [],
      notes: 'A stable feudal kingdom built on grain, textiles and a strong royal army. Its capital is Kingsreach.',
    },
  },
  {
    id: 'ironvein-clanhold',
    label: 'Ironvein Clanhold (dwarven mountain confederation)',
    data: {
      kind: 'nation', size: 'city', population: 61000, biome: 'mountain',
      government: { type: 'Council of clan thanes', leaderName: 'Thane Borrik Ironvein' },
      leadership: [{ title: 'High Thane', name: 'Borrik Ironvein', role: 'ruler', actorId: null }],
      treasury: { cp: 0, sp: 0, gp: 148000, pp: 900 },
      childCityIds: [],
      notes: 'A confederation of dwarven holds bound by trade and shared defence, rich in ore, steel and gemstones.',
    },
  },
  {
    id: 'free-cities-of-the-saltmere',
    label: 'Free Cities of the Saltmere (mercantile league)',
    data: {
      kind: 'nation', size: 'city', population: 88000, biome: 'coastal',
      government: { type: 'Elected merchant assembly', leaderName: 'First Speaker Teysa Norrow' },
      leadership: [{ title: 'First Speaker', name: 'Teysa Norrow', role: 'trade', actorId: null }],
      treasury: { cp: 0, sp: 0, gp: 76000, pp: 150 },
      childCityIds: [],
      notes: 'A loose league of port towns bound by shared trade charters rather than a single crown.',
    },
  },
  {
    id: 'sunwrack-caliphate',
    label: 'Sunwrack Caliphate (desert trade sultanate)',
    data: {
      kind: 'nation', size: 'city', population: 54000, biome: 'desert',
      government: { type: 'Hereditary sultanate + caravan-master council', leaderName: 'Sultana Zahra Reeth' },
      leadership: [{ title: 'Sultana', name: 'Zahra Reeth', role: 'ruler', actorId: null }],
      treasury: { cp: 0, sp: 0, gp: 61000, pp: 220 },
      childCityIds: [],
      notes: 'Desert oasis cities linked by caravan routes; wealth flows from salt, dates and toll roads.',
    },
  },
  {
    id: 'greenwardens-accord',
    label: "Greenwarden's Accord (elven forest nation)",
    data: {
      kind: 'nation', size: 'town', population: 22000, biome: 'ancient forest',
      government: { type: 'Council of speakers', leaderName: 'Ilyenne Duskwhisper' },
      leadership: [{ title: 'First Speaker', name: 'Ilyenne Duskwhisper', role: 'council', actorId: null }],
      treasury: { cp: 0, sp: 0, gp: 34000, pp: 60 },
      childCityIds: [],
      notes: 'Forest enclaves bound by shared stewardship of the ancient woods rather than fixed borders.',
    },
  },
];

export function getPrebuiltNation(id) {
  return PREBUILT_NATIONS.find(n => n.id === id) || null;
}
