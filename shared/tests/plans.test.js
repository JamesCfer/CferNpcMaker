import { describe, it, expect } from 'vitest';
import {
  TIER_PLANS,
  ACTION_COSTS,
  getPlan,
  publicPlans,
  nextPlanUp,
  describeAllowance,
} from '../scripts/core/plans.js';

describe('tier plans', () => {
  it('mirrors the relay TIER_LIMITS allowances', () => {
    const byId = Object.fromEntries(TIER_PLANS.map(p => [p.id, p.uses]));
    expect(byId).toMatchObject({
      free: 3,
      supporter: 25,
      professional: 100,
      premium: 200,
    });
  });

  it('treats every relay free-alias as the free plan', () => {
    for (const alias of ['', 'free', 'none', 'no tier', 'FREE', '  Free  ']) {
      expect(getPlan(alias).id).toBe('free');
    }
    expect(getPlan(null).id).toBe('free');
    expect(getPlan(undefined).id).toBe('free');
  });

  it('matches a known tier case-insensitively and falls back to free', () => {
    expect(getPlan('Supporter').uses).toBe(25);
    expect(getPlan('PREMIUM').uses).toBe(200);
    expect(getPlan('mystery-tier').id).toBe('free');
  });

  it('hides the internal superuser tier from the comparison table', () => {
    expect(publicPlans().map(p => p.id)).toEqual(['free', 'supporter', 'professional', 'premium']);
  });

  it('pitches the cheapest tier that beats the current one', () => {
    expect(nextPlanUp('free').id).toBe('supporter');
    expect(nextPlanUp('supporter').id).toBe('professional');
    expect(nextPlanUp('professional').id).toBe('premium');
    expect(nextPlanUp('premium')).toBeNull();
  });

  it('describes an allowance using the real action costs', () => {
    const npc   = ACTION_COSTS.find(a => a.id === 'generation').cost;
    const store = ACTION_COSTS.find(a => a.id === 'store').cost;
    expect(npc).toBe(1);
    expect(store).toBe(7);
    expect(describeAllowance(25)).toBe('25 NPCs or items, or 3 stores');
    expect(describeAllowance(3)).toBe('3 NPCs or items');
  });
});
