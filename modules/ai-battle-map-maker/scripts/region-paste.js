/** Paste a generated crop with the same soft edge as regions.py. */
export function pasteRegion(base, patch, crop) {
  const result = document.createElement('canvas');
  result.width = base.naturalWidth; result.height = base.naturalHeight;
  const ctx = result.getContext('2d');
  ctx.drawImage(base, 0, 0);
  const layer = document.createElement('canvas');
  layer.width = crop.width; layer.height = crop.height;
  const layerCtx = layer.getContext('2d');
  layerCtx.imageSmoothingQuality = 'high';
  layerCtx.drawImage(patch, 0, 0, crop.width, crop.height);
  const mask = document.createElement('canvas');
  mask.width = crop.width; mask.height = crop.height;
  const maskCtx = mask.getContext('2d');
  const fx = crop.width * 0.07, fy = crop.height * 0.07;
  maskCtx.filter = `blur(${Math.max(fx, fy) / 2}px)`;
  maskCtx.fillStyle = '#fff';
  maskCtx.fillRect(fx, fy, crop.width - 2 * fx, crop.height - 2 * fy);
  layerCtx.globalCompositeOperation = 'destination-in';
  layerCtx.drawImage(mask, 0, 0);
  ctx.drawImage(layer, crop.x, crop.y);
  return result;
}
