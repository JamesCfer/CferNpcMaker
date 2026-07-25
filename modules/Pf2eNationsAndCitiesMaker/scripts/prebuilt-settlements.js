/**
 * Compendium of 10 prebuilt settlements (#111). Ready-to-drop towns for
 * users without AI credits — partial data, sanitizeSettlement() backfills
 * the rest (treasury, stats, hours, etc.) exactly like an AI response would.
 */

function shopId() { return `shop-${Math.random().toString(36).slice(2, 10)}`; }
function staffId() { return `staff-${Math.random().toString(36).slice(2, 10)}`; }

function store(name, type, ownerName, { balance = 60, dailyAvg = 6 } = {}) {
  return {
    id: shopId(), name, type,
    owner: { name: ownerName, actorId: null },
    staff: [{ id: staffId(), name: 'Apprentice', role: 'apprentice', shift: 'day', actorId: null }],
    hours: { open: '08:00', close: '20:00', daysClosed: [] },
    inventory: [],
    income: { balance, dailyAvg, lastTick: 0, daysInDebt: 0 },
  };
}

export const PREBUILT_SETTLEMENTS = [
  {
    id: 'thornwick-crossing',
    label: 'Thornwick Crossing (farming village)',
    data: {
      kind: 'village', size: 'village', biome: 'temperate plains', population: 240,
      government: { type: 'Elected village elder', leaderName: 'Alden Marrow' },
      leadership: [{ title: 'Village Elder', name: 'Alden Marrow', role: 'civic leader', actorId: null }],
      military: {
        ranks: [{ rank: 'Militia', count: 6, leaderName: 'Bess Yorrick', dailyWage: 1 }],
        totalGuards: 6, commanderActorId: null, commanderName: 'Bess Yorrick',
      },
      stores: [
        store('Greenleaf Provisions', 'general', 'Mirra Oat'),
        store('The Tipsy Owl', 'tavern', 'Dob Fenwick'),
      ],
      treasury: { cp: 0, sp: 40, gp: 90, pp: 0 },
      production: ['grain', 'wool'],
      notes: 'A quiet farming village at the crossing of two dirt roads, built around a stone granary.',
    },
  },
  {
    id: 'saltmere-quay',
    label: 'Saltmere Quay (fishing port town)',
    data: {
      kind: 'town', size: 'town', biome: 'coastal', population: 1900,
      government: { type: 'Harbourmaster + fishers council', leaderName: 'Ysolde Pell' },
      leadership: [
        { title: 'Harbourmaster', name: 'Ysolde Pell', role: 'civic leader', actorId: null },
        { title: 'Fishers Council Head', name: 'Otto Brinn', role: 'trade', actorId: null },
      ],
      military: {
        ranks: [
          { rank: 'Sergeant', count: 1, leaderName: 'Rask Coldwater', dailyWage: 3 },
          { rank: 'Guard', count: 10, leaderName: '', dailyWage: 1 },
        ],
        totalGuards: 11, commanderActorId: null, commanderName: 'Rask Coldwater',
      },
      stores: [
        store('Salt & Sail Provisions', 'general', 'Nan Kettle'),
        store('The Drowned Bell', 'tavern', 'Ferro Quill'),
        store('Brinewright Tackle', 'other', 'Odval Brinn'),
        store('Quayside Smithy', 'blacksmith', 'Ules Marnoch'),
      ],
      treasury: { cp: 0, sp: 0, gp: 640, pp: 0 },
      production: ['fish', 'salt', 'rope'],
      notes: 'A stone quay lined with drying racks and net-sheds; boats go out with the tide.',
    },
  },
  {
    id: 'ashenford-outpost',
    label: 'Ashenford Outpost (frontier garrison town)',
    data: {
      kind: 'town', size: 'town', biome: 'borderlands', population: 760,
      government: { type: 'Garrison commander', leaderName: 'Captain Wren Ashcombe' },
      leadership: [{ title: 'Garrison Commander', name: 'Wren Ashcombe', role: 'military', actorId: null }],
      military: {
        ranks: [
          { rank: 'Captain', count: 1, leaderName: 'Wren Ashcombe', dailyWage: 5 },
          { rank: 'Sergeant', count: 3, leaderName: '', dailyWage: 3 },
          { rank: 'Guard', count: 24, leaderName: '', dailyWage: 1 },
        ],
        totalGuards: 28, commanderActorId: null, commanderName: 'Wren Ashcombe',
      },
      stores: [
        store('Palisade Armoury', 'armorer', 'Dagny Holt'),
        store('The Last Watch', 'tavern', 'Iven Sable'),
        store('Frontier Trading Post', 'general', 'Costa Rell'),
      ],
      treasury: { cp: 0, sp: 0, gp: 310, pp: 0 },
      production: ['timber', 'game'],
      crimeLevel: 3,
      notes: 'A palisaded outpost watching the wild border country; most trade caters to soldiers and scouts.',
    },
  },
  {
    id: 'copperveil',
    label: 'Copperveil (mountain dwarven hold)',
    data: {
      kind: 'city', size: 'city', biome: 'mountain', population: 7100,
      government: { type: 'Clan thane + council of elders', leaderName: 'Thane Borrik Ironvein' },
      leadership: [
        { title: 'Clan Thane', name: 'Borrik Ironvein', role: 'ruler', actorId: null },
        { title: 'Council Elder', name: 'Hemma Stonebrow', role: 'council', actorId: null },
      ],
      military: {
        ranks: [
          { rank: 'Hold Captain', count: 1, leaderName: 'Grega Deepforge', dailyWage: 6 },
          { rank: 'Hold Guard', count: 60, leaderName: '', dailyWage: 2 },
        ],
        totalGuards: 61, commanderActorId: null, commanderName: 'Grega Deepforge',
      },
      stores: [
        store('Ironvein Forgeworks', 'blacksmith', 'Dromli Hearthhand', { balance: 400, dailyAvg: 20 }),
        store('Deep Cut Jewelers', 'jeweler', 'Sella Runeglint'),
        store('The Molten Tankard', 'tavern', 'Karn Ashbeard'),
        store('Rune-Scribe Hall', 'magic', 'Vess Coldchisel'),
      ],
      treasury: { cp: 0, sp: 0, gp: 3800, pp: 12 },
      production: ['ore', 'steel', 'gems'],
      notes: 'Halls carved into the mountainside, lit by forge-fire and mushroom-glow along the great avenue.',
    },
  },
  {
    id: 'sunwrack-oasis',
    label: 'Sunwrack Oasis (desert caravan waystation)',
    data: {
      kind: 'village', size: 'village', biome: 'desert', population: 380,
      government: { type: 'Caravan-master council', leaderName: 'Basim Reeth' },
      leadership: [{ title: 'Caravan-master', name: 'Basim Reeth', role: 'civic leader', actorId: null }],
      military: {
        ranks: [{ rank: 'Scout', count: 5, leaderName: 'Yara Sund', dailyWage: 2 }],
        totalGuards: 5, commanderActorId: null, commanderName: 'Yara Sund',
      },
      stores: [
        store('The Deep Well', 'inn', 'Halima Dost'),
        store('Sandglass Traders', 'general', 'Reko Vantal'),
      ],
      treasury: { cp: 0, sp: 0, gp: 150, pp: 0 },
      production: ['dates', 'salt'],
      notes: 'Mudbrick buildings ring a deep well; caravans stop here to water beasts and hire guides.',
    },
  },
  {
    id: 'greywillow-hollow',
    label: 'Greywillow Hollow (elven forest enclave)',
    data: {
      kind: 'town', size: 'town', biome: 'ancient forest', population: 1500,
      government: { type: 'Council of speakers', leaderName: 'Ilyenne Duskwhisper' },
      leadership: [{ title: 'Speaker', name: 'Ilyenne Duskwhisper', role: 'council', actorId: null }],
      military: {
        ranks: [{ rank: 'Ranger', count: 14, leaderName: 'Faelar Greenshade', dailyWage: 2 }],
        totalGuards: 14, commanderActorId: null, commanderName: 'Faelar Greenshade',
      },
      stores: [
        store('Canopy Herbalist', 'apothecary', 'Selwyn Farrow'),
        store('Moonleaf Bindery', 'bookbinder', 'Thess Amberwood'),
        store('The Rope-Bridge Inn', 'inn', 'Auberon Vale'),
      ],
      treasury: { cp: 0, sp: 0, gp: 520, pp: 2 },
      production: ['herbs', 'timber'],
      notes: 'Rope bridges and wooden walkways connect homes built into the canopy of an ancient forest.',
    },
  },
  {
    id: 'kingsreach',
    label: 'Kingsreach (walled capital city)',
    data: {
      kind: 'city', size: 'metropolis', biome: 'temperate hills', population: 26000,
      government: { type: 'Monarchy + royal court', leaderName: 'Queen Meris Duvane' },
      leadership: [
        { title: 'Queen', name: 'Meris Duvane', role: 'ruler', actorId: null },
        { title: 'Lord Chancellor', name: 'Aldous Ferrin', role: 'administration', actorId: null },
      ],
      military: {
        ranks: [
          { rank: 'City Guard Captain', count: 4, leaderName: '', dailyWage: 5 },
          { rank: 'City Guard', count: 220, leaderName: '', dailyWage: 2 },
        ],
        totalGuards: 224, commanderActorId: null, commanderName: 'Captain Rowe Halden',
      },
      stores: [
        store('The Gilded Ledger', 'magic', 'Vantessa Orr', { balance: 900, dailyAvg: 40 }),
        store('Highgate Armory', 'armorer', 'Bram Stell'),
        store('Cathedral Row Temple Shop', 'temple', 'Sister Odalys'),
        store('The Crown and Anchor', 'tavern', 'Pell Duskin'),
      ],
      treasury: { cp: 0, sp: 0, gp: 18500, pp: 40 },
      production: ['grain', 'textiles', 'coin'],
      notes: 'Stone bastions ring a citadel at the city’s heart; noble houses and guilds jockey for influence.',
    },
  },
  {
    id: 'gravel-hollow',
    label: 'Gravel Hollow (mining boomtown)',
    data: {
      kind: 'town', size: 'town', biome: 'mountain foothills', population: 1700,
      government: { type: 'Mining company foreman', leaderName: 'Foreman Address Kray' },
      leadership: [{ title: 'Foreman', name: 'Address Kray', role: 'civic leader', actorId: null }],
      military: {
        ranks: [{ rank: 'Company Man', count: 8, leaderName: '', dailyWage: 2 }],
        totalGuards: 8, commanderActorId: null, commanderName: '',
      },
      stores: [
        store('Kray Company Store', 'general', 'Address Kray'),
        store('The Broken Pick', 'tavern', 'Dallia Marsh'),
        store('Foothill Forge', 'blacksmith', 'Umber Coalson'),
      ],
      treasury: { cp: 0, sp: 0, gp: 210, pp: 0 },
      production: ['ore', 'coal'],
      crimeLevel: 4,
      stats: { unrest: 20 },
      notes: 'Half-built and wooden, grown fast around a productive mine. The "law" is whatever the company says.',
    },
  },
  {
    id: 'pilgrims-rest',
    label: "Pilgrim's Rest (temple pilgrimage town)",
    data: {
      kind: 'town', size: 'town', biome: 'sacred valley', population: 1150,
      government: { type: 'High priest + temple council', leaderName: 'High Priest Amaru Cael' },
      leadership: [{ title: 'High Priest', name: 'Amaru Cael', role: 'religious', actorId: null }],
      military: {
        ranks: [{ rank: 'Temple Guard', count: 12, leaderName: '', dailyWage: 2 }],
        totalGuards: 12, commanderActorId: null, commanderName: '',
      },
      stores: [
        store('Blessing-Token Stall', 'temple', 'Sister Wenna'),
        store('The Pilgrim’s Bowl', 'inn', 'Corwan Delft'),
        store('Waystaff Outfitters', 'general', 'Nessa Ohr'),
      ],
      treasury: { cp: 0, sp: 0, gp: 480, pp: 1 },
      production: ['candles', 'incense'],
      religions: [{ id: 'rel-pilgrims-rest', name: 'The Radiant Path', followers: 900, templeStoreId: null, influence: 70 }],
      notes: 'Pilgrims arrive daily to a great temple complex; the high priest holds the real authority.',
    },
  },
  {
    id: 'riverbend-landing',
    label: 'Riverbend Landing (river-trade town)',
    data: {
      kind: 'town', size: 'town', biome: 'river valley', population: 3400,
      government: { type: 'Mayor + merchant guild', leaderName: 'Mayor Farris Coombe' },
      leadership: [
        { title: 'Mayor', name: 'Farris Coombe', role: 'civic leader', actorId: null },
        { title: 'Guildmaster', name: 'Teysa Norrow', role: 'trade', actorId: null },
      ],
      military: {
        ranks: [
          { rank: 'Sergeant', count: 2, leaderName: '', dailyWage: 3 },
          { rank: 'Watch', count: 18, leaderName: '', dailyWage: 1 },
        ],
        totalGuards: 20, commanderActorId: null, commanderName: '',
      },
      stores: [
        store('Wharfside Spice & Wool', 'general', 'Teysa Norrow', { balance: 300, dailyAvg: 15 }),
        store('The Barge Bell', 'tavern', 'Onnick Farr'),
        store('Ironmonger’s Dock', 'blacksmith', 'Coll Vantry'),
      ],
      treasury: { cp: 0, sp: 0, gp: 1200, pp: 3 },
      production: ['spice', 'wool', 'iron'],
      tradeRoutes: [{ partnerId: '', goods: ['wool', 'iron'], gpPerWeek: 40 }],
      notes: 'Barges unload spice, wool and iron at the docks daily; the merchant guild dominates local politics.',
    },
  },
];

export function getPrebuiltSettlement(id) {
  return PREBUILT_SETTLEMENTS.find(s => s.id === id) || null;
}
