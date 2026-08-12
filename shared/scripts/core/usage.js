/**
 * Monthly-use tracking.
 *
 * The n8n relay owns the real quota — the client never decrements anything
 * authoritative. This module extracts whatever usage figures a response
 * happens to carry (headers or body), caches them per module in
 * localStorage, and hands the shared BuilderApp a number to display.
 *
 * Nothing here throws: when the server says nothing about usage, `getUsage`
 * returns null and the UI simply hides the indicator.
 */

import { Storage } from './storage.js';

/**
 * @typedef {object} Usage
 * @property {number|null} used       Uses consumed this period, when known.
 * @property {number|null} limit      Monthly allowance for the user's tier.
 * @property {number|null} remaining  Uses left this period.
 * @property {number|null} resetAt    Unix ms when the allowance resets.
 * @property {number}      fetchedAt  Unix ms when this snapshot was taken.
 */

const HEADER_KEYS = {
  remaining: ['x-uses-remaining', 'x-ratelimit-remaining'],
  limit:     ['x-uses-limit',     'x-ratelimit-limit'],
  used:      ['x-uses-used',      'x-ratelimit-used'],
  reset:     ['x-uses-reset',     'x-ratelimit-reset'],
};

/** @returns {number|null} */
function num(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

/**
 * Normalises a reset stamp. Accepts Unix seconds, Unix milliseconds, or an
 * ISO date string; returns Unix milliseconds.
 * @param {unknown} value
 * @returns {number|null}
 */
function resetToMs(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  if (Number.isFinite(n)) return n > 1e11 ? n : n * 1000;
  const parsed = new Date(String(value)).getTime();
  return Number.isNaN(parsed) ? null : parsed;
}

/**
 * Pulls usage figures out of a fetch Response and/or its parsed body.
 *
 * Body shapes understood:
 *   { usage: { used, limit, remaining, reset } }
 *   { used, limit, remaining, reset|resetAt }   ← also the 429 payload shape
 *
 * @param {Response|null} response
 * @param {object|null}   [bodyData]
 * @returns {Usage|null} null when the payload carries no usage information.
 */
export function parseUsage(response, bodyData) {
  const headers = response?.headers;
  const fromHeader = (names) => {
    if (!headers?.get) return null;
    for (const name of names) {
      const v = headers.get(name);
      if (v !== null && v !== undefined && v !== '') return v;
    }
    return null;
  };

  const body = (bodyData && typeof bodyData === 'object')
    ? (bodyData.usage && typeof bodyData.usage === 'object' ? bodyData.usage : bodyData)
    : {};

  let used      = num(body.used      ?? fromHeader(HEADER_KEYS.used));
  let limit     = num(body.limit     ?? fromHeader(HEADER_KEYS.limit));
  let remaining = num(body.remaining ?? fromHeader(HEADER_KEYS.remaining));
  const resetAt = resetToMs(body.reset ?? body.resetAt ?? fromHeader(HEADER_KEYS.reset));

  // A 429 says "limit reached" without spelling out that remaining is 0.
  if (remaining === null && response?.status === 429 && limit !== null) remaining = 0;

  if (remaining === null && limit !== null && used !== null) remaining = Math.max(0, limit - used);
  if (used      === null && limit !== null && remaining !== null) used = Math.max(0, limit - remaining);

  if (used === null && limit === null && remaining === null && resetAt === null) return null;

  return { used, limit, remaining, resetAt, fetchedAt: Date.now() };
}

/**
 * Reads the cached snapshot for a module, discarding it once its reset
 * stamp has passed (the allowance has rolled over, so the numbers are lies).
 *
 * @param {string} moduleFolder
 * @returns {Usage|null}
 */
export function getUsage(moduleFolder) {
  if (!moduleFolder) return null;
  const usage = new Storage(moduleFolder).getUsage();
  if (!usage) return null;
  if (usage.resetAt && usage.resetAt < Date.now()) return null;
  return usage;
}

/**
 * Merges a fresh snapshot over the cached one — a response that only carries
 * `remaining` shouldn't wipe a previously-known `limit`.
 *
 * @param {string}     moduleFolder
 * @param {Usage|null} usage  A parseUsage() result; null is ignored.
 * @returns {Usage|null} The merged snapshot that was stored.
 */
export function recordUsage(moduleFolder, usage) {
  if (!moduleFolder) return null;
  if (!usage) return getUsage(moduleFolder);

  const prev = getUsage(moduleFolder) || {};
  const merged = {
    used:      usage.used      ?? prev.used      ?? null,
    limit:     usage.limit     ?? prev.limit     ?? null,
    remaining: usage.remaining ?? prev.remaining ?? null,
    resetAt:   usage.resetAt   ?? prev.resetAt   ?? null,
    fetchedAt: usage.fetchedAt || Date.now(),
  };
  new Storage(moduleFolder).setUsage(merged);
  return merged;
}

/**
 * Optimistically subtracts a generation's cost so the indicator updates
 * without waiting for the next server round-trip. No-op when usage is unknown.
 *
 * @param {string} moduleFolder
 * @param {number} cost
 * @returns {Usage|null}
 */
export function spendLocally(moduleFolder, cost) {
  if (!moduleFolder) return null;
  const usage = getUsage(moduleFolder);
  const spend = Number(cost) || 0;
  if (!usage || spend <= 0) return usage;
  if (usage.remaining === null && usage.used === null) return usage;

  const next = {
    ...usage,
    remaining: usage.remaining === null ? null : Math.max(0, usage.remaining - spend),
    used:      usage.used      === null ? null : usage.used + spend,
  };
  new Storage(moduleFolder).setUsage(next);
  return next;
}

/**
 * Localises a string when Foundry is around, falling back to the English
 * default outside it (tests, or before i18n is ready).
 *
 * @param {string} key
 * @param {object} data
 * @param {string} fallback
 * @returns {string}
 */
function t(key, data, fallback) {
  try {
    const out = game?.i18n?.format?.(key, data);
    if (out && out !== key) return out;
  } catch (_) {}
  return fallback;
}

/**
 * @param {Usage|null} usage
 * @returns {string} Display text ('' when there is nothing worth showing).
 */
export function formatUsage(usage) {
  if (!usage) return '';
  if (usage.remaining === null) return '';
  const { remaining, limit } = usage;
  return limit !== null
    ? t('NpcBuilder.Usage.Left',        { remaining, limit }, `${remaining} / ${limit} uses left`)
    : t('NpcBuilder.Usage.LeftNoLimit', { remaining },        `${remaining} uses left`);
}

/**
 * @param {Usage|null} usage
 * @returns {string} Tooltip text, e.g. 'Resets in 12 days.' ('' when unknown).
 */
export function formatUsageTooltip(usage) {
  if (!usage?.resetAt) return '';
  const days = Math.ceil((usage.resetAt - Date.now()) / 86400000);
  if (days <= 0)  return t('NpcBuilder.Usage.ResetToday',    {},      'Your monthly allowance resets today.');
  if (days === 1) return t('NpcBuilder.Usage.ResetTomorrow', {},      'Your monthly allowance resets tomorrow.');
  return             t('NpcBuilder.Usage.ResetDays',         { days }, `Your monthly allowance resets in ${days} days.`);
}
