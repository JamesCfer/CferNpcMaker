import { describe, it, expect, beforeEach } from 'vitest';
import { isNewerVersion, escapeHtml }        from '../scripts/core/utils.js';
import { devUrl }                            from '../scripts/core/n8n.js';
import { Storage }                           from '../scripts/core/storage.js';

describe('isNewerVersion', () => {
  it('returns true when major is higher', () => {
    expect(isNewerVersion('2.0.0', '1.9.9')).toBe(true);
  });

  it('returns true when minor is higher', () => {
    expect(isNewerVersion('1.1.0', '1.0.9')).toBe(true);
  });

  it('returns true when patch is higher', () => {
    expect(isNewerVersion('1.0.1', '1.0.0')).toBe(true);
  });

  it('returns false when equal', () => {
    expect(isNewerVersion('1.0.0', '1.0.0')).toBe(false);
  });

  it('returns false when a is older', () => {
    expect(isNewerVersion('0.9.9', '1.0.0')).toBe(false);
    expect(isNewerVersion('1.0.0', '2.0.0')).toBe(false);
  });

  it('handles missing segments (pads with 0)', () => {
    expect(isNewerVersion('2', '1.9.9')).toBe(true);
    expect(isNewerVersion('1.1', '1.0.9')).toBe(true);
  });
});

describe('escapeHtml', () => {
  it('escapes < and >', () => {
    expect(escapeHtml('<script>')).toBe('&lt;script&gt;');
  });

  it('escapes ampersands', () => {
    expect(escapeHtml('a & b')).toBe('a &amp; b');
  });

  it('passes through plain text unchanged', () => {
    expect(escapeHtml('hello world')).toBe('hello world');
  });

  it('coerces non-strings to string', () => {
    expect(escapeHtml(42)).toBe('42');
    expect(typeof escapeHtml(null)).toBe('string');
  });
});

describe('devUrl', () => {
  const base = 'https://example.com/webhook/npc-builder';

  it('returns base url unchanged when devMode is false', () => {
    expect(devUrl(base, false)).toBe(base);
  });

  it('appends -dev suffix when devMode is true', () => {
    expect(devUrl(base, true)).toBe(`${base}-dev`);
  });
});

describe('Storage', () => {
  let storage;

  beforeEach(() => {
    localStorage.clear();
    storage = new Storage('test-module');
  });

  it('getKey returns empty string when nothing stored', () => {
    expect(storage.getKey()).toBe('');
  });

  it('setKey / getKey round-trip', () => {
    storage.setKey('abc123key');
    expect(storage.getKey()).toBe('abc123key');
  });

  it('setKey with empty string clears the key', () => {
    storage.setKey('some-key');
    storage.setKey('');
    expect(storage.getKey()).toBe('');
  });

  it('loadHistory returns [] when nothing stored', () => {
    expect(storage.loadHistory()).toEqual([]);
  });

  it('saveHistory / loadHistory round-trip', () => {
    const entry = { id: 'abc', name: 'Bob', status: 'success', createdAt: 1000, error: null };
    storage.saveHistory([entry]);
    expect(storage.loadHistory()).toEqual([entry]);
  });

  it('loadHistory honours maxEntries limit', () => {
    const entries = Array.from({ length: 5 }, (_, i) => ({
      id: String(i), name: `NPC ${i}`, status: 'success', createdAt: i, error: null,
    }));
    storage.saveHistory(entries);
    expect(storage.loadHistory(3)).toHaveLength(3);
  });

  it('getArtStyle returns empty string when unset', () => {
    expect(storage.getArtStyle()).toBe('');
  });

  it('setArtStyle / getArtStyle round-trip', () => {
    storage.setArtStyle('oil painting');
    expect(storage.getArtStyle()).toBe('oil painting');
  });

  it('clear removes all module keys', () => {
    storage.setKey('k');
    storage.setArtStyle('s');
    storage.saveHistory([{ id: 'x', name: 'X', status: 'success', createdAt: 0, error: null }]);
    storage.clear();
    expect(storage.getKey()).toBe('');
    expect(storage.getArtStyle()).toBe('');
    expect(storage.loadHistory()).toEqual([]);
  });
});
