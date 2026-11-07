#!/usr/bin/env python3
"""Region edit test: the GM boxes part of a finished map and describes a change.

Three ways to do it:
  crop-gemini  crop the box plus a margin, edit the crop with Gemini at 1K, paste it back with a soft edge
  crop-klein   the same with FLUX.2 klein
  full-gemini  send the whole map to Gemini at 2K and describe where to change it

  python3 tools/map-bakeoff/regions.py --methods crop-gemini,crop-klein,full-gemini

Needs OPENROUTER_API_KEY. Python because the test needs crop and paste; the
module does the same with a canvas in the browser. Output: out/regions/.
"""
import argparse, base64, io, json, os, ssl, sys, time, urllib.request
from concurrent.futures import ThreadPoolExecutor
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out', 'regions')
KEY = os.environ.get('OPENROUTER_API_KEY') or sys.exit('Set OPENROUTER_API_KEY.')
CA = '/root/.ccr/ca-bundle.crt'
CTX = ssl.create_default_context(cafile=CA) if os.path.exists(CA) else ssl.create_default_context()

# Box is (left, top, right, bottom) as a share of the map. "where" describes it for full-gemini.
EDITS = [
    {'id': 'tavern-brawl', 'map': 'out/boost/final-gemini/p02-tavern.png', 'box': (0.55, 0.45, 0.95, 0.95), 'where': 'the bottom-right quarter of the map',
     'change': 'A brawl has wrecked this area: tables overturned, broken chairs, and spilled ale across the floor.'},
    {'id': 'crypt-collapse', 'map': 'out/boost/final-gemini/p01-crypt.png', 'box': (0.3, 0.0, 0.7, 0.45), 'where': 'the top centre of the map, under the opening in the ceiling',
     'change': 'More of the ceiling has caved in: a heap of broken stone blocks and rubble now fills the water under the opening.'},
    {'id': 'town-fire', 'map': 'out/boost/final-gemini/t01-town-topdown.png', 'box': (0.38, 0.35, 0.62, 0.68), 'where': 'the centre of the town square',
     'change': 'The well has been destroyed: in its place is a burned-out bonfire of charred beams with ash and scorch marks on the cobbles.'},
]
MODELS = {
    'crop-gemini': ('google/gemini-3.1-flash-image', '1K'),
    'crop-klein': ('black-forest-labs/flux.2-klein-4b', '1K'),
    'full-gemini': ('google/gemini-3.1-flash-image', '2K'),
}
ASPECTS = {'1:1': 1, '4:3': 4 / 3, '3:4': 3 / 4, '16:9': 16 / 9, '9:16': 9 / 16, '3:2': 1.5, '2:3': 2 / 3}
KEEP = 'Do not draw any grid lines, text, labels, border, or frame. Keep the same art style, lighting, and straight-down camera angle.'


def call(model, resolution, prompt, img):
    buf = io.BytesIO(); img.convert('RGB').save(buf, 'PNG')
    ratio = img.width / img.height
    aspect = min(ASPECTS, key=lambda k: abs(ASPECTS[k] - ratio))
    body = {'n': 1, 'model': model, 'prompt': prompt, 'resolution': resolution, 'aspect_ratio': aspect,
            'input_references': [{'type': 'image_url', 'image_url': {'url': 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()}}]}
    req = urllib.request.Request('https://openrouter.ai/api/v1/images', data=json.dumps(body).encode(),
                                 headers={'Authorization': 'Bearer ' + KEY, 'Content-Type': 'application/json', 'X-Title': 'Battle map region edit test'})
    with urllib.request.urlopen(req, context=CTX, timeout=300) as r:
        data = json.load(r)
    return Image.open(io.BytesIO(base64.b64decode(data['data'][0]['b64_json']))).convert('RGB'), data.get('usage', {}).get('cost', 0)


def padded_box(box, w, h, pad=0.15):
    """Pixel box around the GM's box with a margin of pad times the box size, clamped to the map."""
    l, t, r, b = box[0] * w, box[1] * h, box[2] * w, box[3] * h
    px, py = max(32, (r - l) * pad), max(32, (b - t) * pad)
    return tuple(int(v) for v in (max(0, l - px), max(0, t - py), min(w, r + px), min(h, b + py)))


def run(edit, method):
    model, res = MODELS[method]
    base = Image.open(os.path.join(HERE, edit['map'])).convert('RGB')
    w, h = base.size
    t0 = time.time()
    if method == 'full-gemini':
        prompt = f"Edit this battle map. Change only {edit['where']}: {edit['change']} Keep everything else on the map exactly the same, pixel for pixel. {KEEP}"
        out, cost = call(model, res, prompt, base)
        out = out.resize((w, h), Image.LANCZOS)
    else:
        box = padded_box(edit['box'], w, h)
        crop = base.crop(box)
        prompt = f"This is a cut-out piece of a top-down battle map. {edit['change']} Keep the edges of the image, the walls, and the floor layout the same so it fits back into the map. {KEEP}"
        patch, cost = call(model, res, prompt, crop)
        patch = patch.resize(crop.size, Image.LANCZOS)
        # Full strength inside the GM's box, fading to nothing across the margin.
        mask = Image.new('L', crop.size, 0)
        fx, fy = (box[2] - box[0]) * 0.07, (box[3] - box[1]) * 0.07
        ImageDraw.Draw(mask).rectangle((fx, fy, crop.width - fx, crop.height - fy), fill=255)
        mask = mask.filter(ImageFilter.GaussianBlur(max(fx, fy) / 2))
        out = base.copy(); out.paste(patch, box[:2], mask)
    name = f"{edit['id']}--{method}.jpg"
    out.save(os.path.join(OUT, name), quality=90)
    return {'id': edit['id'], 'method': method, 'file': name, 'costUsd': cost, 'seconds': round(time.time() - t0, 1)}


if __name__ == '__main__':
    ap = argparse.ArgumentParser(); ap.add_argument('--methods', default=','.join(MODELS)); a = ap.parse_args()
    os.makedirs(OUT, exist_ok=True)
    jobs = [(e, m) for e in EDITS for m in a.methods.split(',')]
    results = []
    with ThreadPoolExecutor(9) as pool:
        for fut, (e, m) in zip([pool.submit(run, e, m) for e, m in jobs], jobs):
            try:
                r = fut.result(); results.append(r); print(f"  ok   {r['file']}: ${r['costUsd']:.4f}, {r['seconds']}s")
            except Exception as err:
                print(f"  FAIL {e['id']} {m}: {err}")
    for e in EDITS:  # the untouched map with the GM's box drawn on it, for review
        im = Image.open(os.path.join(HERE, e['map'])).convert('RGB'); w, h = im.size
        ImageDraw.Draw(im).rectangle((e['box'][0] * w, e['box'][1] * h, e['box'][2] * w, e['box'][3] * h), outline=(255, 60, 60), width=6)
        im.save(os.path.join(OUT, f"{e['id']}--before.jpg"), quality=90)
    json.dump(results, open(os.path.join(OUT, 'regions.json'), 'w'), indent=2)
    print(f"Done. ${sum(r['costUsd'] or 0 for r in results):.4f} spent.")
