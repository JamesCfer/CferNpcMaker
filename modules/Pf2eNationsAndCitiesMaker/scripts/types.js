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
 * @property {number}           priceMultiplier  multiplied into all displayed prices (1.0 = normal) (#62)
 * @property {string[]}         childCityIds
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
 * Army document shape (#70). Lives on its own JournalEntry, same flag path
 * as Settlement/Nation. `commanderActorId` bonuses (#78) are still
 * unimplemented — the field exists so later cycles don't need another
 * schema migration.
 * @typedef {object} Army
 * @property {'army'} kind
 * @property {string|null} stationedAt    journal ID of the home settlement
 * @property {'garrison'|'field'} mode    garrisons defend the settlement, field armies can march (#74)
 * @property {string|null} destination    journal ID being marched to (#75); null when not moving
 * @property {CalendarDate|null} arrivalDate  date the army reaches `destination` (#75)
 * @property {Array<{ id: string, type: string, count: number, level: number, equipment: string, morale: number }>} units
 * @property {string|null} commanderActorId
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
 * Stub — Treaty shape (Section F diplomacy, not yet implemented).
 * @typedef {object} Treaty
 * @property {string}           id
 * @property {string}           partnerNationId
 * @property {string}           kind     non-aggression|defensive|trade|vassalage
 * @property {CalendarDate}     signedOn
 * @property {CalendarDate|null} expiresOn
 * @property {string}           terms
 */
