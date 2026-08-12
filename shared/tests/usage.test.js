import { describe, it, expect, beforeEach } from 'vitest';
import { parseUsage, recordUsage, getUsage, spendLocally,
         formatUsage, formatUsageTooltip }  from '../scripts/core/usage.js';
import { Storage }                          from '../scripts/core/storage.js';

const FOLDER = 'TestStoreModule';

/** Minimal stand-in for a fetch Response's header bag. */
function headers(map) {
  const lower = Object.fromEntries(Object.entries(map).map(([k, v]) => [k.toLowerCase(), String(v)]));
  return { get: (name) => (name.toLowerCase() in lower ? lower[name.toLowerCase()] : null) };
}

function response({ status = 200, headers: h = {} } = {}) {
  return { status, headers: headers(h) };
}

beforeEach(() => {
  localStorage.clear();
});

describe('parseUsage', () => {
  it('returns null when neither headers nor body mention usage', () => {
    expect(parseUsage(response(), { ok: true, store: {} })).toBeNull();
    expect(parseUsage(null, null)).toBeNull();
  });

  it('reads the X-Uses-* headers', () => {
    const u = parseUsage(response({ headers: { 'X-Uses-Remaining': '38', 'X-Uses-Limit': '50' } }), null);
    expect(u.remaining).toBe(38);
    expect(u.limit).toBe(50);
    expect(u.used).toBe(12);
  });

  it('reads a nested usage object from the body', () => {
    const u = parseUsage(response(), { usage: { used: 7, limit: 50 } });
    expect(u.used).toBe(7);
    expect(u.remaining).toBe(43);
  });

  it('reads the flat 429 payload shape and infers remaining 0', () => {
    const u = parseUsage(response({ status: 429 }), { message: 'Monthly limit reached.', limit: 15 });
    expect(u.limit).toBe(15);
    expect(u.remaining).toBe(0);
    expect(u.used).toBe(15);
  });

  it('prefers body values over headers', () => {
    const u = parseUsage(response({ headers: { 'X-Uses-Remaining': '5' } }), { usage: { remaining: 9 } });
    expect(u.remaining).toBe(9);
  });

  it('normalises Unix seconds, Unix ms, and ISO reset stamps', () => {
    const secs = parseUsage(response(), { remaining: 1, reset: 1_700_000_000 });
    expect(secs.resetAt).toBe(1_700_000_000_000);

    const ms = parseUsage(response(), { remaining: 1, reset: 1_700_000_000_000 });
    expect(ms.resetAt).toBe(1_700_000_000_000);

    const iso = parseUsage(response(), { remaining: 1, resetAt: '2030-01-01T00:00:00Z' });
    expect(iso.resetAt).toBe(new Date('2030-01-01T00:00:00Z').getTime());
  });

  it('ignores unparseable numbers rather than storing NaN', () => {
    const u = parseUsage(response({ headers: { 'X-Uses-Remaining': 'lots' } }), null);
    expect(u).toBeNull();
  });
});

describe('recordUsage / getUsage', () => {
  it('round-trips a snapshot through localStorage', () => {
    recordUsage(FOLDER, parseUsage(response(), { usage: { remaining: 20, limit: 50 } }));
    expect(getUsage(FOLDER).remaining).toBe(20);
    expect(new Storage(FOLDER).getUsage().limit).toBe(50);
  });

  it('merges over the cached snapshot instead of wiping known fields', () => {
    recordUsage(FOLDER, parseUsage(response(), { usage: { remaining: 20, limit: 50 } }));
    recordUsage(FOLDER, parseUsage(response({ headers: { 'X-Uses-Remaining': '13' } }), null));

    const u = getUsage(FOLDER);
    expect(u.remaining).toBe(13);
    expect(u.limit).toBe(50);
  });

  it('is a no-op without a module folder', () => {
    expect(recordUsage('', parseUsage(response(), { remaining: 3 }))).toBeNull();
    expect(getUsage('')).toBeNull();
  });

  it('discards a snapshot whose reset stamp has passed', () => {
    recordUsage(FOLDER, { used: 50, limit: 50, remaining: 0, resetAt: Date.now() - 1000, fetchedAt: Date.now() });
    expect(getUsage(FOLDER)).toBeNull();
  });

  it('keeps a snapshot whose reset stamp is still ahead', () => {
    recordUsage(FOLDER, { used: 1, limit: 50, remaining: 49, resetAt: Date.now() + 86400000, fetchedAt: Date.now() });
    expect(getUsage(FOLDER).remaining).toBe(49);
  });
});

describe('spendLocally', () => {
  beforeEach(() => {
    recordUsage(FOLDER, parseUsage(response(), { usage: { used: 10, limit: 50, remaining: 40 } }));
  });

  it('subtracts a store generation (7 uses)', () => {
    const u = spendLocally(FOLDER, 7);
    expect(u.remaining).toBe(33);
    expect(u.used).toBe(17);
  });

  it('clamps remaining at zero', () => {
    expect(spendLocally(FOLDER, 999).remaining).toBe(0);
  });

  it('ignores non-positive costs', () => {
    expect(spendLocally(FOLDER, 0).remaining).toBe(40);
    expect(spendLocally(FOLDER, -5).remaining).toBe(40);
  });

  it('does nothing when usage is unknown', () => {
    localStorage.clear();
    expect(spendLocally(FOLDER, 7)).toBeNull();
  });
});

describe('formatUsage', () => {
  it('renders remaining and limit', () => {
    expect(formatUsage({ remaining: 38, limit: 50 })).toBe('38 / 50 uses left');
  });

  it('renders remaining alone when the limit is unknown', () => {
    expect(formatUsage({ remaining: 38, limit: null })).toBe('38 uses left');
  });

  it('renders nothing when remaining is unknown', () => {
    expect(formatUsage(null)).toBe('');
    expect(formatUsage({ remaining: null, limit: 50 })).toBe('');
  });
});

describe('formatUsageTooltip', () => {
  it('is empty without a reset stamp', () => {
    expect(formatUsageTooltip({ remaining: 1 })).toBe('');
  });

  it('counts whole days to the reset', () => {
    expect(formatUsageTooltip({ resetAt: Date.now() + 12 * 86400000 }))
      .toBe('Your monthly allowance resets in 12 days.');
  });

  it('handles today and tomorrow', () => {
    expect(formatUsageTooltip({ resetAt: Date.now() - 1000 })).toBe('Your monthly allowance resets today.');
    expect(formatUsageTooltip({ resetAt: Date.now() + 3600000 })).toBe('Your monthly allowance resets tomorrow.');
  });
});
