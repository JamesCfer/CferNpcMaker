/** @type {Array<{ name: string, priceGp: number, production: string[] }>} */
export const TRADE_GOODS = [
  { name: 'Iron Ingot',       priceGp: 1,    production: ['iron', 'mining', 'metal', 'smelting'] },
  { name: 'Steel Bar',        priceGp: 5,    production: ['iron', 'metal', 'smith', 'forging'] },
  { name: 'Copper Ore',       priceGp: 1,    production: ['copper', 'mining', 'metal'] },
  { name: 'Silver Ore',       priceGp: 5,    production: ['silver', 'mining', 'metal'] },
  { name: 'Gold Dust',        priceGp: 50,   production: ['gold', 'mining', 'metal'] },
  { name: 'Mithral Chunk',    priceGp: 500,  production: ['mithral', 'mining', 'metal', 'dwarven'] },
  { name: 'Coal',             priceGp: 1,    production: ['coal', 'mining', 'fire', 'fuel'] },
  { name: 'Wheat (bushel)',   priceGp: 0.1,  production: ['grain', 'wheat', 'farming', 'food', 'agriculture'] },
  { name: 'Flour (sack)',     priceGp: 0.2,  production: ['grain', 'wheat', 'milling', 'food'] },
  { name: 'Barley (bushel)',  priceGp: 0.1,  production: ['grain', 'barley', 'farming', 'brewing', 'food'] },
  { name: 'Salt (lb)',        priceGp: 0.1,  production: ['salt', 'mining', 'coastal', 'preservation'] },
  { name: 'Spices (lb)',      priceGp: 10,   production: ['spices', 'trade', 'exotic', 'food'] },
  { name: 'Dried Herbs',      priceGp: 0.5,  production: ['herbs', 'alchemy', 'medicine', 'foraging', 'forest'] },
  { name: 'Raw Lumber',       priceGp: 0.2,  production: ['lumber', 'wood', 'forestry', 'timber'] },
  { name: 'Finished Planks',  priceGp: 1,    production: ['lumber', 'wood', 'carpentry', 'timber'] },
  { name: 'Ship Timber',      priceGp: 5,    production: ['lumber', 'timber', 'shipbuilding', 'forestry'] },
  { name: 'Wool (bolt)',      priceGp: 2,    production: ['wool', 'livestock', 'weaving', 'textile'] },
  { name: 'Cloth (bolt)',     priceGp: 5,    production: ['cloth', 'weaving', 'textile', 'trade'] },
  { name: 'Leather (hide)',   priceGp: 1,    production: ['leather', 'livestock', 'tanning', 'hide'] },
  { name: 'Tallow Candles',   priceGp: 0.1,  production: ['livestock', 'tallow', 'candles', 'wax'] },
  { name: 'Ale (barrel)',     priceGp: 2,    production: ['brewing', 'grain', 'tavern', 'trade'] },
  { name: 'Wine (bottle)',    priceGp: 1,    production: ['wine', 'viticulture', 'trade', 'brewing'] },
  { name: 'Dried Fish',       priceGp: 0.5,  production: ['fish', 'fishing', 'food', 'coastal', 'seafaring'] },
  { name: 'Salted Meat',      priceGp: 0.5,  production: ['livestock', 'meat', 'food', 'hunting'] },
  { name: 'Honey (jar)',      priceGp: 2,    production: ['honey', 'beekeeping', 'farming', 'food'] },
  { name: 'Rope (coil)',      priceGp: 1,    production: ['rope', 'hemp', 'seafaring', 'trade'] },
  { name: 'Raw Gems',         priceGp: 25,   production: ['gems', 'mining', 'jewelry', 'crystal'] },
  { name: 'Cut Gemstones',    priceGp: 100,  production: ['gems', 'jewelry', 'lapidary', 'trade'] },
  { name: 'Tanned Hides',     priceGp: 3,    production: ['leather', 'tanning', 'hide', 'livestock'] },
  { name: 'Common Herbs',     priceGp: 0.2,  production: ['herbs', 'foraging', 'farming', 'plants'] },
];

export function goodsForProduction(productionTags) {
  if (!Array.isArray(productionTags) || productionTags.length === 0) return [];
  const tags = productionTags.map(t => t.toLowerCase().trim()).filter(Boolean);
  if (!tags.length) return [];
  return TRADE_GOODS.filter(g =>
    g.production.some(p => tags.some(t => t.includes(p) || p.includes(t)))
  );
}
