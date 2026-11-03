import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { imageSize, extFor, sessionCost, buildPrompt, toFoundryWalls } from './lib.mjs';
import { buildPrompt as sharedBuildPrompt, toFoundryWalls as sharedToFoundryWalls } from '../../shared/scripts/core/maps/index.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const prompts = JSON.parse(fs.readFileSync(path.join(HERE, 'prompts.json'), 'utf8'));
const edits = JSON.parse(fs.readFileSync(path.join(HERE, 'edits.json'), 'utf8'));

describe('map bake-off helpers', () => {
  it('reads PNG and JPEG dimensions', () => {
    const png = new Uint8Array(32);
    png.set([0x89, 0x50, 0x4e, 0x47]);
    new DataView(png.buffer).setUint32(16, 2048);
    new DataView(png.buffer).setUint32(20, 1365);
    expect(imageSize(png)).toEqual({ width: 2048, height: 1365 });

    const jpg = new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x05, 0x55, 0x08, 0x00, 0, 0, 0]);
    expect(imageSize(jpg)).toEqual({ width: 2048, height: 1365 });
  });

  it('maps media types to extensions', () => {
    expect(extFor('image/jpeg')).toBe('jpg');
    expect(extFor('image/webp')).toBe('webp');
  });

  it('prices a session as 8 maps and 4 edits', () => {
    expect(sessionCost(0.045, 0.025)).toBeCloseTo(0.46);
  });

  it('has 20 prompts with unique ids, valid perspectives, and edits that point at real prompts', () => {
    expect(prompts).toHaveLength(20);
    expect(new Set(prompts.map(p => p.id)).size).toBe(20);
    for (const p of prompts) expect(() => buildPrompt(p)).not.toThrow();
    for (const e of edits) expect(prompts.some(p => p.id === e.source)).toBe(true);
  });
  it('re-exports the shared prompt and wall helpers by identity', () => {
    expect(buildPrompt).toBe(sharedBuildPrompt);
    expect(toFoundryWalls).toBe(sharedToFoundryWalls);
  });
});
