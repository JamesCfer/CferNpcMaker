import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';

describe('Foundry map script boundary', () => {
  it('keeps look.js independent of prompt.js', () => {
    const look = readFileSync('shared/scripts/core/maps/look.js', 'utf8');
    expect(look).not.toMatch(/from\s*['"]\.\/prompt\.js['"]/);
    expect(look).toMatch(/from\s*['"]\.\/dials\.js['"]/);
  });
});
