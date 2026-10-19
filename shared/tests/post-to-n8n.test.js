import { describe, it, expect, afterEach, vi } from 'vitest';
import { postToN8n, AuthError, RateLimitError } from '../scripts/core/adapter.js';

function mockFetch(status, body) {
  globalThis.fetch = vi.fn(async () => new Response(JSON.stringify(body), { status }));
}

describe('postToN8n error mapping', () => {
  afterEach(() => { vi.restoreAllMocks(); });

  it('treats a 401 that reports rateLimitExceeded as a rate limit, not a bad session', async () => {
    mockFetch(401, {
      ok: false, rateLimitExceeded: true, error: 'RATE_LIMIT_EXCEEDED',
      message: 'A store generation costs 7 uses.', limit: 25, tier: 'supporter',
      resetDate: '2026-11-01T00:00:00.000Z',
    });
    const err = await postToN8n('https://x/webhook/a', {}, 'k').catch(e => e);
    expect(err).toBeInstanceOf(RateLimitError);
    expect(err).not.toBeInstanceOf(AuthError);
    expect(err.tier).toBe('Supporter');
    expect(err.resetAt).toBe(Date.parse('2026-11-01T00:00:00.000Z'));
  });

  it('still treats an expired session as an auth error', async () => {
    mockFetch(401, { ok: false, error: 'SESSION_EXPIRED', requiresAuth: true, message: 'Your session has expired.' });
    const err = await postToN8n('https://x/webhook/a', {}, 'k').catch(e => e);
    expect(err).toBeInstanceOf(AuthError);
  });

  it('maps 429 to a rate limit', async () => {
    mockFetch(429, { limit: 3, message: 'Monthly limit reached.' });
    const err = await postToN8n('https://x/webhook/a', {}, 'k').catch(e => e);
    expect(err).toBeInstanceOf(RateLimitError);
    expect(err.tier).toBe('Free');
  });
});
