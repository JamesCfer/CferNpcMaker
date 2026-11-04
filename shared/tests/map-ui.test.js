import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { renderMapThumbnails } from '../../modules/ai-battle-map-maker/scripts/adapter.js';

const markup = readFileSync(resolve('modules/ai-battle-map-maker/templates/form.html'), 'utf8');
const stylesheet = readFileSync(resolve('modules/ai-battle-map-maker/styles/theme.css'), 'utf8');

describe('battle map form flow', () => {
  for (const count of [1, 2, 4]) {
    it(`renders ${count} mockups after the actions in a ${count === 1 ? 1 : 2} column grid`, () => {
      const container = document.createElement('div');
      container.innerHTML = markup;
      const form = container.querySelector('form');
      const box = form.querySelector('.map-thumbnails');
      renderMapThumbnails(box, Array.from({ length: count }, () => ({ b64: 'AAAA', mediaType: 'image/png' })), '4:3', () => {});

      expect([...form.querySelectorAll('*')].some(el => el.style.position)).toBe(false);
      expect(box.style.getPropertyValue('--cols')).toBe(count === 1 ? '1' : '2');
      expect(box.style.getPropertyValue('--map-aspect')).toBe('4 / 3');
      expect(box.querySelectorAll('img')).toHaveLength(count);
      for (const image of box.querySelectorAll('img')) {
        expect(image.style.width).toBe('');
        expect(image.style.height).toBe('');
        expect(image.parentElement.tagName).toBe('BUTTON');
      }
      expect(form.querySelector('.map-actions-section').compareDocumentPosition(form.querySelector('.map-mockup-section')) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
      box.querySelector('button').click();
      expect(box.querySelector('.map-picked').textContent).toBe('Picked');
    });
  }
  it('resets the Foundry button height for mockup tiles', () => {
    const style = document.createElement('style');
    style.textContent = stylesheet;
    const app = document.createElement('div');
    app.id = 'ai-battle-map-maker-app';
    app.className = 'application npc-builder module-ai-battle-map-maker';
    app.innerHTML = markup;
    document.head.append(style);
    document.body.append(app);
    try {
      const box = app.querySelector('.map-thumbnails');
      renderMapThumbnails(box, [{ b64: 'AAAA', mediaType: 'image/png' }], '4:3', () => {});
      expect(getComputedStyle(box.querySelector('.map-thumbnail')).height).toBe('auto');
    } finally {
      app.remove();
      style.remove();
    }
  });
});

