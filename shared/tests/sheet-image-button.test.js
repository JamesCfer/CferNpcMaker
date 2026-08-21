import { describe, it, expect, beforeEach, vi } from 'vitest';
import { hasDefaultImage, injectSheetImageButton } from '../scripts/core/sheet-image-button.js';

/** Minimal stand-in for a Foundry document. */
function makeDoc(img, defaultImg) {
  class Doc {
    static getDefaultArtwork() { return { img: defaultImg }; }
    constructor(i) { this.img = i; }
    toObject() { return { img: this.img }; }
  }
  return new Doc(img);
}

function makeSheet(imgSrc) {
  const root = document.createElement('div');
  root.innerHTML = `<header class="sheet-header">
    <img class="profile" data-edit="img" src="${imgSrc}">
    <div class="details">fields</div>
  </header>`;
  return root;
}

describe('hasDefaultImage', () => {
  it('treats a missing image as default', () => {
    expect(hasDefaultImage(makeDoc('', 'icons/svg/item-bag.svg'))).toBe(true);
  });

  it('treats the document type default artwork as default', () => {
    expect(hasDefaultImage(makeDoc('systems/pf2e/icons/default-icons/feat.svg',
                                   'systems/pf2e/icons/default-icons/feat.svg'))).toBe(true);
  });

  it('treats core placeholder icons as default', () => {
    expect(hasDefaultImage(makeDoc('icons/svg/mystery-man.svg', 'nope'))).toBe(true);
    expect(hasDefaultImage(makeDoc('icons/svg/item-bag.svg', 'nope'))).toBe(true);
  });

  it('treats a real portrait as custom', () => {
    expect(hasDefaultImage(makeDoc('npc-images/goblin-ab12cd34.webp', 'icons/svg/mystery-man.svg'))).toBe(false);
  });
});

describe('injectSheetImageButton', () => {
  let onClick;
  beforeEach(() => { onClick = vi.fn(); });

  it('mounts the button under the portrait when artwork is default', () => {
    const root = makeSheet('icons/svg/mystery-man.svg');
    const doc  = makeDoc('icons/svg/mystery-man.svg', 'icons/svg/mystery-man.svg');

    expect(injectSheetImageButton({ root, doc, title: 't', onClick })).toBe(true);

    const slot = root.querySelector('.npc-builder-image-slot');
    expect(slot).toBeTruthy();
    expect(slot.children[0].tagName).toBe('IMG');
    expect(slot.children[1].classList.contains('npc-builder-sheet-image-btn')).toBe(true);
    // The header keeps its other content untouched.
    expect(root.querySelector('.details')).toBeTruthy();
  });

  it('does nothing when a custom image is already set', () => {
    const root = makeSheet('npc-images/goblin.webp');
    const doc  = makeDoc('npc-images/goblin.webp', 'icons/svg/mystery-man.svg');

    expect(injectSheetImageButton({ root, doc, title: 't', onClick })).toBe(false);
    expect(root.querySelector('.npc-builder-sheet-image-btn')).toBeNull();
  });

  it('does not inject twice', () => {
    const root = makeSheet('icons/svg/mystery-man.svg');
    const doc  = makeDoc('icons/svg/mystery-man.svg', 'icons/svg/mystery-man.svg');

    injectSheetImageButton({ root, doc, title: 't', onClick });
    expect(injectSheetImageButton({ root, doc, title: 't', onClick })).toBe(false);
    expect(root.querySelectorAll('.npc-builder-sheet-image-btn').length).toBe(1);
  });

  it('fires the click handler', () => {
    const root = makeSheet('icons/svg/mystery-man.svg');
    const doc  = makeDoc('icons/svg/mystery-man.svg', 'icons/svg/mystery-man.svg');
    injectSheetImageButton({ root, doc, title: 't', onClick });
    root.querySelector('.npc-builder-sheet-image-btn').click();
    expect(onClick).toHaveBeenCalledOnce();
  });
});
