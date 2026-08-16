/**
 * JSDoc typedefs for the Pf2eNationsAndCitiesMaker data model.
 * This file has no runtime exports — it exists only for IDE type-checking.
 */

/**
 * @typedef {object} Staff
 * @property {string}      id
 * @property {string}      name
 * @property {string}      role
 * @property {'morning'|'day'|'evening'|'night'|'graveyard'} shift
 * @property {string|null} actorId
 */

/**
 * @typedef {object} InventoryItem
 * @property {string}      id
 * @property {string}      name
 * @property {string|null} itemId
 * @property {number}      price
 * @property {number}      stock
 */

/**
 * @typedef {object} StoreIncome
 * @property {number} balance
 * @property {number} dailyAvg
 * @property {number} lastTick
 * @property {number} daysInDebt   game-days balance has been continuously negative
 */

/**
 * Adventurer guild job posting (#115) — lives on a store of type 'guild'.
 * @typedef {object} Job
 * @property {string}      id
 * @property {string}      title
 * @property {string}      description
 * @property {number}      reward           gp offered on completion
 * @property {'open'|'claimed'|'completed'|'failed'} status
 * @property {string|null} assignedActorId  NPC rolled to take the job
 */

/**
 * @typedef {object} Store
 * @property {string}        id
 * @property {string}        name
 * @property {string}        type
 * @property {boolean}       closed         true when bankrupt 30+ game-days
 * @property {number|null}   marketWeekday  0–6 weekday index; income doubles on this day (#59)
 * @property {boolean}       isBlackMarket  hidden from players; attracts +1 unrest/day per store (#60)
 * @property {'low'|'standard'|'high'|'luxury'} priceTier  wealth tier; multiplies displayed prices (#42)
 * @property {{ name: string, actorId: string|null }} owner
 * @property {Staff[]}       staff
 * @property {{ open: string, close: string, daysClosed: string[] }} hours
 * @property {InventoryItem[]} inventory
 * @property {StoreIncome}   income
 * @property {Job[]}         jobs           adventurer guild job board (#115); type:'guild' stores only
 */

/**
 * @typedef {object} Religion
 * @property {string}      id
 * @property {string}      name
 * @property {number}      followers
 * @property {string|null} templeStoreId  links to a store of type 'temple' (#66)
 * @property {number}      influence      0–100
 */

/**
 * Faction influence within a city (#90); the GM's own bookkeeping — the sheet
 * shows a running total but does not force it to 100.
 * @typedef {object} Faction
 * @property {string} id
 * @property {string} name
 * @property {'church'|'guild'|'nobles'|'criminal'|'crown'} type
 * @property {number} influence  0–100
 */

/**
 * @typedef {object} Rank
 * @property {string} rank
 * @property {number} count
 * @property {string} leaderName
 * @property {number} dailyWage   gp drained from treasury per guard per day
 */

/**
 * @typedef {object} SettlementStats
 * @property {number} hp
 * @property {number} maxHp
 * @property {number} damageThreshold
 * @property {number} hardness
 * @property {number} fortitude
 * @property {number} reflex
 * @property {number} will
 * @property {number} morale
 * @property {number} unrest
 * @property {boolean} occupied  true once hp has been sieged to 0 (#77)
 * @property {string|null} occupiedBy  occupying nation's journal ID; drained daily for gp (#80)
 */

/**
 * @typedef {object} Treasury
 * @property {number} cp
 * @property {number} sp
 * @property {number} gp
 * @property {number} pp
 */

/**
 * @typedef {object} LeadershipEntry
 * @property {string}      title
 * @property {string}      name
 * @property {string}      role
 * @property {string|null} actorId
 */

/**
 * @typedef {object} Military
 * @property {Rank[]}      ranks
 * @property {number}      totalGuards
 * @property {string|null} commanderActorId
 * @property {string}      commanderName
 */

/**
 * A nation's standing with one other nation (#83). Lives on the Nation
 * document's own `relations` list — each nation tracks its side of the
 * relationship independently, so a pair can (deliberately) disagree.
 * @typedef {object} NationRelation
 * @property {string} nationId
 * @property {'ally'|'friendly'|'neutral'|'cold'|'hostile'} relation
 * @property {number} score  -100..100
 */

/**
 * @typedef {object} Settlement
 * @property {number}           _schemaVersion
 * @property {string}           kind           city|town|village|nation
 * @property {string}           size
 * @property {number}           population
 * @property {string}           biome
 * @property {SettlementStats}  stats
 * @property {Treasury}         treasury
 * @property {string[]}         production
 * @property {{ type: string, leaderActorId: string|null, leaderName: string }} government
 * @property {LeadershipEntry[]} leadership
 * @property {Military}         military
 * @property {Store[]}          stores
 * @property {Religion[]}       religions
 * @property {Faction[]}        factions         city-tier; influence breakdown of local power groups (#90)
 * @property {number}           priceMultiplier  multiplied into all displayed prices (1.0 = normal) (#62)
 * @property {string[]}         childCityIds
 * @property {NationRelation[]} relations        nation-tier only; standing with every other Nation (#83)
 * @property {Treaty[]}         treaties         nation-tier only; treaties this nation has signed (#84, #85)
 * @property {string[]}         vassalNationIds  nation-tier only; nations that owe this nation fealty (#86)
 * @property {string|null}      suzerainNationId nation-tier only; nation this nation owes fealty to (#86)
 * @property {Claim[]}          claims           nation-tier only; casus-belli claims on other settlements (#87)
 * @property {{ actorId: string|null, name: string }} heir  nation-tier only; ascends on a `rulerDied` calendar event (#91)
 * @property {string|null}      bannerImage      header banner image path; AI-generated or user-picked (#102)
 * @property {string}           notes
 * @property {{ endpoint: string, model: string, prompt: string }} ai
 */

/**
 * @typedef {object} CalendarDate
 * @property {number} year
 * @property {number} month
 * @property {number} day
 */

/**
 * @typedef {object} ScheduledEvent
 * @property {string}       id
 * @property {string}       kind     tax|payday|festival|custom
 * @property {string}       label
 * @property {CalendarDate} nextFire
 * @property {{ every: number, unit: string }} recurrence
 * @property {object}       payload
 */

/**
 * Mercenary contract (#81) — set when an army was hired as a preset company
 * rather than recruited. Once `expiresDate` passes, the army journal is
 * deleted automatically (no refund) on the next `dayAdvanced` tick.
 * @typedef {object} Contract
 * @property {boolean} active
 * @property {string|null} companyId      key into MERCENARY_COMPANIES
 * @property {string} companyLabel
 * @property {CalendarDate|null} expiresDate
 */

/**
 * Army document shape (#70). Lives on its own JournalEntry, same flag path
 * as Settlement/Nation. `commanderActorId` gives a flat power bonus in
 * battles and sieges equal to the linked Actor's level (#78).
 * @typedef {object} Army
 * @property {'army'} kind
 * @property {string|null} stationedAt    journal ID of the home settlement
 * @property {'garrison'|'field'} mode    garrisons defend the settlement, field armies can march (#74)
 * @property {string|null} destination    journal ID being marched to (#75); null when not moving
 * @property {CalendarDate|null} arrivalDate  date the army reaches `destination` (#75)
 * @property {Array<{ id: string, type: string, count: number, level: number, equipment: string, morale: number }>} units
 * @property {string|null} commanderActorId
 * @property {string|null} ownerNationId  nation this army belongs to; occupier of record on a successful siege (#80)
 * @property {string|null} supplySource   settlement journal ID the army draws food from each day (#79)
 * @property {Contract} contract          set when hired as a mercenary company (#81)
 * @property {string} notes
 */

/**
 * Battle resolution result (#76) — deterministic, computed from two Army
 * unit rosters and a terrain. Casualties are written back onto both armies.
 * @typedef {object} BattleResult
 * @property {string} terrain
 * @property {number} attackerPower
 * @property {number} defenderPower
 * @property {'attacker'|'defender'|'draw'} winner
 * @property {number} attackerCasualtyPct
 * @property {number} defenderCasualtyPct
 * @property {Array<{ id: string, type: string, lost: number, remaining: number }>} attackerCasualties
 * @property {Array<{ id: string, type: string, lost: number, remaining: number }>} defenderCasualties
 */

/**
 * Siege resolution result (#77) — an attacking Army versus a Settlement's
 * HP, hardness and damageThreshold. HP reaching 0 flips `occupied`.
 * @typedef {object} SiegeResult
 * @property {string} terrain
 * @property {number} attackerPower
 * @property {number} damage
 * @property {number} hpBefore
 * @property {number} hpAfter
 * @property {boolean} occupied
 */

/**
 * Treaty between two nations (#84). Lives on the signing nation's own
 * `treaties` list, same one-sided-tracking philosophy as {@link NationRelation}
 * — a treaty isn't mirrored onto the partner's document. Expiry is checked
 * on every `dayAdvanced` tick and posts a chat reminder (#85).
 * @typedef {object} Treaty
 * @property {string}           id
 * @property {string}           partnerNationId
 * @property {string}           kind     non-aggression|defensive|trade|vassalage
 * @property {CalendarDate|null} signedOn
 * @property {CalendarDate|null} expiresOn  null means the treaty never expires
 * @property {string}           terms
 */

/**
 * A nation's claim on a settlement (#87) — casus belli material cited by the
 * declare-war flow (#88) and given up or kept during peace negotiation (#89).
 * Lives on the claiming nation's own `claims` list.
 * @typedef {object} Claim
 * @property {string} id
 * @property {string} targetSettlementId
 * @property {'historical'|'dynastic'|'religious'} kind
 * @property {string} notes
 */
