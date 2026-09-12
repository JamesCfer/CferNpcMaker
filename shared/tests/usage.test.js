import { describe, it, expect, beforeEach } from 'vitest';
import {
  getUsage,
  setUsage,
  clearUsage,
  onUsageChange,
  extractUsageFromData,
  updateUsageFromResponse,
} from '../scripts/core/usage.js';

function fakeResponse({ headers = {}, status = 200 } = {}) {
  return {
    status,
    headers: { get: (name) => headers[name] ?? null },
  };
}

describe('usage tracker', () => {
  beforeEach(() => clearUsage());

  it('starts with unknown usage', () => {
    expect(getUsage()).toMatchObject({ remaining: null, limit: null });
  });

  it('setUsage stores remaining and limit', () => {
    setUsage(12, 50);
    expect(getUsage()).toMatchObject({ remaining: 12, limit: 50 });
  });

  it('setUsage ignores null / invalid values', () => {
    setUsage(12, 50);
    setUsage(null, null);
    setUsage('not-a-number', -3);
    expect(getUsage()).toMatchObject({ remaining: 12, limit: 50 });
  });

  it('clearUsage forgets stored numbers', () => {
    setUsage(12, 50);
    clearUsage();
    expect(getUsage()).toMatchObject({ remaining: null, limit: null });
  });

  it('notifies listeners on change and supports unsubscribe', () => {
    const seen = [];
    const off = onUsageChange(u => seen.push(u.remaining));
    setUsage(9);
    setUsage(9);     // no change → no extra notification
    off();
    setUsage(3);
    expect(seen).toEqual([9]);
  });

  describe('extractUsageFromData', () => {
    it('reads usesRemaining / limit at the top level', () => {
      expect(extractUsageFromData({ usesRemaining: 7, limit: 50 }))
        .toMatchObject({ remaining: 7, limit: 50 });
    });

    it('reads remaining / usesLeft variants', () => {
      expect(extractUsageFromData({ remaining: 4 }).remaining).toBe(4);
      expect(extractUsageFromData({ usesLeft: 2 }).remaining).toBe(2);
    });

    it('reads a nested usage object', () => {
      expect(extractUsageFromData({ usage: { remaining: 11, limit: 80 } }))
        .toMatchObject({ remaining: 11, limit: 80 });
    });

    it('derives remaining from limit - used', () => {
      expect(extractUsageFromData({ limit: 50, used: 30 }).remaining).toBe(20);
    });

    it('returns nulls for empty or invalid input', () => {
      expect(extractUsageFromData(null)).toEqual({ remaining: null, limit: null, tier: null, resetAt: null });
      expect(extractUsageFromData({})).toEqual({ remaining: null, limit: null, tier: null, resetAt: null });
    });
  });

  describe('updateUsageFromResponse', () => {
    it('prefers response headers', () => {
      updateUsageFromResponse(
        fakeResponse({ headers: { 'X-Uses-Remaining': '6', 'X-Uses-Limit': '15' } }),
        { usesRemaining: 99 }
      );
      expect(getUsage()).toMatchObject({ remaining: 6, limit: 15 });
    });

    it('falls back to the JSON body', () => {
      updateUsageFromResponse(fakeResponse(), { usesRemaining: 42, limit: 50 });
      expect(getUsage()).toMatchObject({ remaining: 42, limit: 50 });
    });

    it('treats a 429 with no usage info as zero remaining', () => {
      updateUsageFromResponse(fakeResponse({ status: 429 }), {});
      expect(getUsage().remaining).toBe(0);
    });

    it('does nothing when neither headers nor body carry usage', () => {
      setUsage(8, 15);
      updateUsageFromResponse(fakeResponse(), { message: 'ok' });
      expect(getUsage()).toMatchObject({ remaining: 8, limit: 15 });
    });
  });
});

describe('tier and reset tracking', () => {
  beforeEach(() => clearUsage());

  it('records a lowercased tier and an ISO reset date', () => {
    const resetDate = '2026-10-01T00:00:00.000Z';
    updateUsageFromResponse(fakeResponse(), {
      ok: true, used: 8, limit: 25, remaining: 17, tier: 'Supporter', resetDate,
    });
    expect(getUsage()).toMatchObject({
      remaining: 17,
      limit: 25,
      tier: 'supporter',
      resetAt: Date.parse(resetDate),
    });
  });

  it('ignores an unparseable reset date', () => {
    updateUsageFromResponse(fakeResponse(), { remaining: 5, limit: 25, resetDate: 'soon' });
    expect(getUsage().resetAt).toBeNull();
  });

  it('accepts a tier nested under usage', () => {
    expect(extractUsageFromData({ usage: { remaining: 2, limit: 3, tier: 'free' } }))
      .toMatchObject({ remaining: 2, limit: 3, tier: 'free' });
  });

  it('clearUsage forgets the tier and reset date', () => {
    setUsage(4, 25, 'supporter', Date.now() + 86400000);
    clearUsage();
    expect(getUsage()).toMatchObject({ tier: null, resetAt: null });
  });

  it('a 429 with a tier zeroes the allowance but keeps the tier', () => {
    setUsage(2, 25, 'supporter');
    updateUsageFromResponse(fakeResponse({ status: 429 }), { tier: 'supporter' });
    expect(getUsage()).toMatchObject({ remaining: 0, tier: 'supporter' });
  });
});
