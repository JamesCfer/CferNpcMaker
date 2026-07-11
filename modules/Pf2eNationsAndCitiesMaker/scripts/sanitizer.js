/**
 * Sanitizer for AI-returned settlement payloads. Coerces missing/garbage fields
 * into safe defaults so the custom sheet never sees holes.
 */

import { MODULE_ID } from './constants.js';
import { CURRENT_SCHEMA_VERSION } from './migrations.js';

const KINDS  = new Set(['city', 'town', 'village', 'nation']);
const SIZES  = new Set(['thorp', 'hamlet', 'village', 'town', 'city', 'metropolis']);

function safeNum(n, def, min = -Infinity, max = Infinity) {
  const v = Number(n);
  if (!Number.isFinite(v)) return def;
  return Math.min(max, Math.max(min, v));
}

function safeString(s, def = '') {
  return (typeof s === 'string' && s.trim()) ? s.trim() : def;
}

function shortId() { return Math.random().toString(36).slice(2, 10); }

export function sanitizeSettlement(raw, formData = {}) {
  const s = (raw && typeof raw === 'object') ? raw : {};
  const kind = KINDS.has(s.kind) ? s.kind : (formData.kind || 'town');
  const size = SIZES.has(s.size) ? s.size : (formData.size || 'town');

  const stats = s.stats && typeof s.stats === 'object' ? s.stats : {};
  const treasury = s.treasury && typeof s.treasury === 'object' ? s.treasury : {};

  const maxHp = safeNum(stats.maxHp, safeNum(stats.hp, 100, 1, 9999), 1, 9999);
  const hp    = safeNum(stats.hp, maxHp, 0, maxHp);

  const out = {
    _schemaVersion: CURRENT_SCHEMA_VERSION,
    kind,
    size,
    population: safeNum(s.population, formData.populationHint || 1000, 1, 9_999_999),
    biome:      safeString(s.biome, formData.biome || ''),
    stats: {
      hp, maxHp,
      damageThreshold: safeNum(stats.damageThreshold, 10, 0, 999),
      hardness:        safeNum(stats.hardness,         5, 0, 999),
      fortitude: safeNum(stats.fortitude, 10, -10, 60),
      reflex:    safeNum(stats.reflex,    10, -10, 60),
      will:      safeNum(stats.will,      10, -10, 60),
      morale:    safeNum(stats.morale,    60, 0, 100),
      unrest:    safeNum(stats.unrest,     5, 0, 100),
    },
    treasury: {
      cp: safeNum(treasury.cp, 0, 0, 9_999_999),
      sp: safeNum(treasury.sp, 0, 0, 9_999_999),
      gp: safeNum(treasury.gp, 0, 0, 9_999_999),
      pp: safeNum(treasury.pp, 0, 0, 9_999_999),
    },
    production: Array.isArray(s.production) ? s.production.map(p => safeString(p)).filter(Boolean) : [],
    government: {
      type:           safeString(s.government?.type, formData.governmentHint || 'Council'),
      leaderActorId:  s.government?.leaderActorId || null,
      leaderName:     safeString(s.government?.leaderName, ''),
    },
    leadership: Array.isArray(s.leadership) ? s.leadership.map(l => ({
      title:    safeString(l?.title, 'Council member'),
      name:     safeString(l?.name, 'Unnamed'),
      role:     safeString(l?.role, ''),
      actorId:  l?.actorId || null,
    })) : [],
    military: {
      ranks: Array.isArray(s.military?.ranks) ? s.military.ranks.map(r => ({
        rank:       safeString(r?.rank, 'Guard'),
        count:      safeNum(r?.count, 1, 0, 99999),
        leaderName: safeString(r?.leaderName, ''),
        dailyWage:  safeNum(r?.dailyWage, 0, 0, 9_999_999),
      })) : [],
      totalGuards:       safeNum(s.military?.totalGuards, 0, 0, 999999),
      commanderActorId:  s.military?.commanderActorId || null,
      commanderName:     safeString(s.military?.commanderName, ''),
    },
    districts: Array.isArray(s.districts) ? s.districts.map(d => ({
      id:            safeString(d?.id, `dist-${shortId()}`),
      name:          safeString(d?.name, 'New District'),
      descriptor:    safeString(d?.descriptor, ''),
      leaderActorId: d?.leaderActorId || null,
    })) : [],
    gmOnly: {
      treasury: !!s.gmOnly?.treasury,
      military: !!s.gmOnly?.military,
      income:   !!s.gmOnly?.income,
    },
    stores: Array.isArray(s.stores) ? s.stores.map(st => ({
      id:         safeString(st?.id, `shop-${shortId()}`),
      name:       safeString(st?.name, 'Unnamed Shop'),
      type:       safeString(st?.type, 'general'),
      closed:     !!st?.closed,
      districtId: (typeof st?.districtId === 'string' && st.districtId) ? st.districtId : null,
      owner: {
        name:    safeString(st?.owner?.name, 'Unknown'),
        actorId: st?.owner?.actorId || null,
      },
      staff: Array.isArray(st?.staff) ? st.staff.map(p => ({
        id:      safeString(p?.id, `staff-${shortId()}`),
        name:    safeString(p?.name, 'Worker'),
        role:    safeString(p?.role, 'employee'),
        shift:   safeString(p?.shift, 'day'),
        actorId: p?.actorId || null,
      })) : [],
      hours: {
        open:       safeString(st?.hours?.open,  '08:00'),
        close:      safeString(st?.hours?.close, '20:00'),
        daysClosed: Array.isArray(st?.hours?.daysClosed) ? st.hours.daysClosed.map(d => safeString(d)).filter(Boolean) : [],
        schedule: (() => {
          const dfOpen  = safeString(st?.hours?.open,  '08:00');
          const dfClose = safeString(st?.hours?.close, '20:00');
          const raw     = Array.isArray(st?.hours?.schedule) ? st.hours.schedule : [];
          return Array.from({ length: 7 }, (_, i) => {
            const entry = raw.find(e => Number(e?.day) === i);
            return {
              day:    i,
              open:   safeString(entry?.open,  dfOpen),
              close:  safeString(entry?.close, dfClose),
              closed: !!entry?.closed,
            };
          });
        })(),
      },
      inventory: Array.isArray(st?.inventory) ? st.inventory.map(i => ({
        id:     safeString(i?.id, `inv-${shortId()}`),
        name:   safeString(i?.name, 'Item'),
        itemId: i?.itemId || null,
        price:  safeNum(i?.price, 1, 0, 9_999_999),
        stock:  safeNum(i?.stock, 1, 0, 9_999_999),
      })) : [],
      income: {
        balance:    safeNum(st?.income?.balance,    0, -9_999_999, 9_999_999),
        dailyAvg:   safeNum(st?.income?.dailyAvg,   5,  0, 9_999_999),
        lastTick:   safeNum(st?.income?.lastTick,   0,  0, 9_999_999_999_999),
        daysInDebt: safeNum(st?.income?.daysInDebt, 0,  0, 9_999_999),
      },
      marketWeekday: (st?.marketWeekday != null) ? safeNum(st.marketWeekday, 0, 0, 6) : null,
      isBlackMarket: !!st?.isBlackMarket,
      priceTier: ['low', 'standard', 'high', 'luxury'].includes(st?.priceTier) ? st.priceTier : 'standard',
    })) : [],
    priceMultiplier: safeNum(s.priceMultiplier, 1.0, 0.1, 10.0),
    growthRate: safeNum(s.growthRate, 0.001, 0, 1),
    famineDaysLeft: safeNum(s.famineDaysLeft, 0, 0, 9999),
    crimeLevel: safeNum(s.crimeLevel, 0, 0, 10),
    treasuryHistory: Array.isArray(s.treasuryHistory)
      ? s.treasuryHistory.slice(-30).map(h => ({ gp: safeNum(h?.gp, 0, -9_999_999, 9_999_999) }))
      : [],
    tradeRoutes: Array.isArray(s.tradeRoutes)
      ? s.tradeRoutes.map(r => ({
          partnerId: safeString(r?.partnerId, ''),
          goods: Array.isArray(r?.goods) ? r.goods.join(', ') : safeString(r?.goods, ''),
          gpPerWeek: safeNum(r?.gpPerWeek, 0, 0, 9_999_999),
        }))
      : [],
    religions: Array.isArray(s.religions) ? s.religions.map(r => ({
      id:            safeString(r?.id, `rel-${shortId()}`),
      name:          safeString(r?.name, 'Unknown Religion'),
      followers:     safeNum(r?.followers, 0, 0, 9_999_999),
      templeStoreId: r?.templeStoreId || null,
      influence:     safeNum(r?.influence, 0, 0, 100),
    })) : [],
    demographics: Array.isArray(s.demographics) ? s.demographics.map(d => ({
      ancestry: safeString(d?.ancestry, 'Human'),
      pct:      safeNum(d?.pct, 0, 0, 100),
    })) : [],
    childCityIds: Array.isArray(s.childCityIds) ? s.childCityIds.filter(x => typeof x === 'string') : [],
    sceneId: (typeof s.sceneId === 'string' && s.sceneId) ? s.sceneId : null,
    notes: safeString(s.notes, ''),
    ai: {
      endpoint: safeString(s.ai?.endpoint, 'city-builder'),
      model:    safeString(s.ai?.model, ''),
      prompt:   safeString(s.ai?.prompt, formData.description || ''),
    },
  };

  const KIND_STORE_CAPS = { city: 12, nation: 12, town: 6, village: 2 };
  const storeCap = KIND_STORE_CAPS[kind] ?? 6;
  out.stores = out.stores.slice(0, storeCap);

  return out;
}

/** Returns the journal flag path that holds the settlement payload. */
export const SETTLEMENT_FLAG_PATH = ['flags', MODULE_ID, 'settlement'];
