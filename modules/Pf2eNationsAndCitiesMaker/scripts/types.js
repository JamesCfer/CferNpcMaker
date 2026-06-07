/**
 * JSDoc typedefs for the Pf2eNationsAndCitiesMaker data model.
 * This file has no runtime exports — it exists only for IDE type-checking.
 */

/**
 * @typedef {object} Staff
 * @property {string}      id
 * @property {string}      name
 * @property {string}      role
 * @property {string}      shift   morning|day|evening|night|graveyard
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
 * @typedef {object} Store
 * @property {string}        id
 * @property {string}        name
 * @property {string}        type
 * @property {boolean}       closed         true when bankrupt 30+ game-days
 * @property {number|null}   marketWeekday  0–6 weekday index; income doubles on this day (#59)
 * @property {boolean}       isBlackMarket  hidden from players; attracts +1 unrest/day per store (#60)
 * @property {{ name: string, actorId: string|null }} owner
 * @property {Staff[]}       staff
 * @property {{ open: string, close: string, daysClosed: string[] }} hours
 * @property {InventoryItem[]} inventory
 * @property {StoreIncome}   income
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
 * @property {number}           priceMultiplier  multiplied into all displayed prices (1.0 = normal) (#62)
 * @property {string[]}         childCityIds
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
 * Stub — Army document shape (Section E military, not yet implemented).
 * @typedef {object} Army
 * @property {string} id
 * @property {string} name
 * @property {string} stationedAt         journal ID of the home settlement
 * @property {'garrison'|'field'} mode
 * @property {Array<{ type: string, count: number, level: number, equipment: string, morale: number }>} units
 * @property {string|null} commanderActorId
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
