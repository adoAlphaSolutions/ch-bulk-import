// ============================================================================
// Toll Product Import Generator — Content Hub External Component (Tile v1)
// ----------------------------------------------------------------------------
// Browser replacement for the Tile macro's Intake -> "Export to Content Hub File".
//
// Input : the vendor INTAKE spreadsheet. Header row (row 1) uses the friendly
//         Tile intake labels (Backsplash Item #, Floor Item #, ..., Manufacturer,
//         Brand, Family/Style Name, Color, Color Family, Look/Style, Shape, Size,
//         Material, Finish, Unit of Measure, Thickness (in), ... reference cols).
//         Optionally include "id" and "identifier" columns for updates.
//
// Transform (mirrors Toll_PCM_Tile_Macros.bas):
//   * The 5 Item # cells concatenate (Floor,Wall,Backsplash,ShowerFloor,Listello)
//     with commas into TB.PCM.E1ItemNumber.
//   * Each filled Item # adds an AreaOfApplication identifier (pipe-joined).
//   * Friendly labels map to CH fields; option-list DISPLAY values are resolved
//     to identifiers using the embedded option-list snapshot (LOOKUPS), which was
//     generated from the same Content Hub package data sources the macro's
//     "Lookups" sheet uses. Reference columns are skipped.
//   * id / identifier pass through when present (updates); blank => new records.
//
// Output: a downloadable workbook with ONE sheet named "M.PCM.Product".
//
// NOTE: LOOKUPS is a snapshot. If the option lists change in Content Hub,
// regenerate it from the package datasources and update this file.
// ============================================================================

const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const SHEET_NAME = 'M.PCM.Product';
const CH_HOST = window.location.origin;

// Option-list snapshot (fallback): source -> [[identifier, displayLabel], ...].
// This can be overridden at runtime by an external file (config.lookupsUrl) or
// by the "Sync from Content Hub" button.
const DEFAULT_LOOKUPS = {"TB.PCM.Manufacturer":[["TB.PCM.Manufacturer.Amerock","Amerock"],["TB.PCM.Manufacturer.Anderson","Anderson"],["TB.PCM.Manufacturer.ArizonaTile","Arizona Tile"],["TB.PCM.Manufacturer.BerensonHardware","Berenson Hardware"],["TB.PCM.Manufacturer.CaliforniaMantelFireplace","California Mantel & Fireplace"],["TB.PCM.Manufacturer.Consentino","Consentino"],["TB.PCM.Manufacturer.Cove","Cove"],["TB.PCM.Manufacturer.DalTile","DalTile"],["TB.PCM.Manufacturer.ElectricMirror","Electric Mirror"],["TB.PCM.Manufacturer.Emser","Emser"],["TB.PCM.Manufacturer.Emtek","Emtek"],["TB.PCM.Manufacturer.HomeSiteServices","Home Site Services"],["TB.PCM.Manufacturer.HunterDouglas","Hunter Douglas"],["TB.PCM.Manufacturer.Infratech","Infratech"],["TB.PCM.Manufacturer.JamisonCollections","Jamison Collections"],["TB.PCM.Manufacturer.JennAir","JennAir"],["TB.PCM.Manufacturer.KitchenAid","KitchenAid"],["TB.PCM.Manufacturer.Kohler","Kohler"],["TB.PCM.Manufacturer.Kwikset","Kwikset"],["TB.PCM.Manufacturer.Marazzi","Marazzi"],["TB.PCM.Manufacturer.Metropolitan","Metropolitan"],["TB.PCM.Manufacturer.Mohawk","Mohawk"],["TB.PCM.Manufacturer.PrecisionCabinetry","Precision Cabinetry"],["TB.PCM.Manufacturer.ProgressLighting","Progress Lighting"],["TB.PCM.Manufacturer.Provenza","Provenza"],["TB.PCM.Manufacturer.Shaw","Shaw"],["TB.PCM.Manufacturer.SherwinWilliams","Sherwin Williams"],["TB.PCM.Manufacturer.Sterling","Sterling"],["TB.PCM.Manufacturer.SubZero","Sub-Zero"],["TB.PCM.Manufacturer.TopKnobs","Top Knobs"],["TB.PCM.Manufacturer.Tuftex","Tuftex"],["TB.PCM.Manufacturer.Unbranded","Unbranded"],["TB.PCM.Manufacturer.WesternWindowSystems","Western Window Systems"],["TB.PCM.Manufacturer.Whirlpool","Whirlpool"],["TB.PCM.Manufacturer.Wolf","Wolf"],["TB.PCM.Manufacturer.CenturyCabinetry","Century Cabinetry"],["TB.PCM.Manufacturer.AmericanOlean","American Olean"],["TB.PCM.Manufacturer.UrbanEffects","Urban Effects"],["TB.PCM.Manufacturer.Coyote","Coyote"],["TB.PCM.Manufacturer.MSI","MSI"],["TB.PCM.Manufacturer.ASKO","ASKO"],["TB.PCM.Manufacturer.Mullican","Mullican"],["TB.PCM.Manufacturer.Mannington","Mannington"]],"TB.PCM.Brand":[["TB.PCM.Brand.AmericanOlean","American Olean"],["TB.PCM.Brand.ArizonaTile","Arizona Tile"],["TB.PCM.Brand.Daltile","Daltile"],["TB.PCM.Brand.EmserTile","Emser Tile"],["TB.PCM.Brand.Marazzi","Marazzi"],["TB.PCM.Brand.Shaw","Shaw"]],"TB.PCM.TileColorFamily":[["TB.PCM.TileColorFamily.Beige","Beige"],["TB.PCM.TileColorFamily.Black","Black"],["TB.PCM.TileColorFamily.Blue","Blue"],["TB.PCM.TileColorFamily.Brown","Brown"],["TB.PCM.TileColorFamily.Cream","Cream"],["TB.PCM.TileColorFamily.Gray","Gray"],["TB.PCM.TileColorFamily.Green","Green"],["TB.PCM.TileColorFamily.Multicolor","Multicolor"],["TB.PCM.TileColorFamily.Orange","Orange"],["TB.PCM.TileColorFamily.Pink","Pink"],["TB.PCM.TileColorFamily.Red","Red"],["TB.PCM.TileColorFamily.Taupe","Taupe"],["TB.PCM.TileColorFamily.TerraCotta","Terra Cotta"],["TB.PCM.TileColorFamily.White","White"],["TB.PCM.TileColorFamily.Yellow","Yellow"]],"TB.PCM.TileStyle":[["TB.PCM.TileStyle.Brick","Brick"],["TB.PCM.TileStyle.ConcreteLook","Concrete Look"],["TB.PCM.TileStyle.Fabric","Fabric"],["TB.PCM.TileStyle.Handmade","Handmade"],["TB.PCM.TileStyle.MarbleLook","Marble Look"],["TB.PCM.TileStyle.Mosaics","Mosaics"],["TB.PCM.TileStyle.NaturalStone","Natural Stone"],["TB.PCM.TileStyle.Patterned","Patterned"],["TB.PCM.TileStyle.Solid","Solid"],["TB.PCM.TileStyle.StoneLook","Stone Look"],["TB.PCM.TileStyle.Textured","Textured"],["TB.PCM.TileStyle.WoodLook","Wood Look"]],"TB.PCM.TileSize":[["TB.PCM.TileSize.HerringboneChevron","Herringbone/Chevron"],["TB.PCM.TileSize.Hexagon","Hexagon"],["TB.PCM.TileSize.LargeFormat","Large Format"],["TB.PCM.TileSize.Pebble","Pebble"],["TB.PCM.TileSize.PennyOval","Penny/Oval"],["TB.PCM.TileSize.Plank","Plank"],["TB.PCM.TileSize.Rectangle","Rectangle"],["TB.PCM.TileSize.Square","Square"],["TB.PCM.TileSize.Subway","Subway"],["TB.PCM.TileSize.UniqueShapes","Unique Shapes"]],"TB.PCM.TileMaterial":[["TB.PCM.TileMaterial.Ceramic","Ceramic"],["TB.PCM.TileMaterial.Porcelain","Porcelain"],["TB.PCM.TileMaterial.Glass","Glass"],["TB.PCM.TileMaterial.NaturalStoneMarble","Natural Stone - Marble"],["TB.PCM.TileMaterial.NaturalStoneTravertine","Natural Stone - Travertine"],["TB.PCM.TileMaterial.NaturalStoneLimestone","Natural Stone - Limestone"],["TB.PCM.TileMaterial.NaturalStoneMixed","Natural Stone - Mixed"]],"TB.PCM.TileFinish":[["TB.PCM.TileFinish.3DSculptural","3D/Sculptural"],["TB.PCM.TileFinish.Glossy","Glossy"],["TB.PCM.TileFinish.Matte","Matte"],["TB.PCM.TileFinish.Satin","Satin"],["TB.PCM.TileFinish.SlipResistant","Slip Resistant"],["TB.PCM.TileFinish.Smooth","Smooth"],["TB.PCM.TileFinish.Textured","Textured"]],"TB.PCM.UnitOfMeasure":[["TB.PCM.UnitOfMeasure.SQF","SQF"],["TB.PCM.UnitOfMeasure.Each","Each"]]};

// Option lists are read LIVE from Content Hub on demand and cached for the
// session here (source -> [[identifier, label], ...]). DEFAULT_LOOKUPS is only
// used as an offline fallback when the live read is unavailable.
let liveCache = {};
const DATA_SOURCES = Object.keys(DEFAULT_LOOKUPS);

// Field -> backing option-list source (mirrors GetLookupSourceForCHField).
const FIELD_SOURCE = {
  'TB.PCM.Product.Manufacturer': 'TB.PCM.Manufacturer',
  'TB.PCM.Product.TileUnitOfMeasure': 'TB.PCM.UnitOfMeasure',
  'TB.PCM.DivisionSelected': 'TB.Division'
};
function sourceFor(field) { return FIELD_SOURCE[field] || field; }

// Default category selection (Tile).
const DEFAULT_CATEGORY_ID = 'TB.PCM.Category.Tiles';
// TB.PCM.Category taxonomy snapshot: [ [categoryId, label, [[subId, subLabel], ...]], ... ]
const CATEGORIES = [["TB.PCM.Category.Appliances","Appliances",[["TB.PCM.Category.CookingPackage","Cooking Package"],["TB.PCM.Category.LaundryAppliances","Laundry Appliances"],["TB.PCM.Category.Other","Other"],["TB.PCM.Category.Refrigerators","Refrigerators"],["TB.PCM.Category.UnderCabinetRefrigeration","Under Cabinet Refrigeration"]]],["TB.PCM.Category.Baseboards","Baseboards",[]],["TB.PCM.Category.Plumbing.BathHardware","Bath Hardware",[]],["TB.PCM.Category.BathTubs","Bath Tubs",[]],["TB.PCM.Category.CabinetEnhancements","Cabinet Enhancements",[]],["TB.PCM.Category.CabinetHardware","Cabinet Hardware",[]],["TB.PCM.Category.Cabinets.Cabinets","Cabinets",[]],["TB.PCM.Category.Cabinets","Cabinets",[]],["TB.PCM.Category.Flooring.Carpet","Carpet",[]],["TB.PCM.Category.Countertops","Countertops",[]],["TB.PCM.Category.Doors","Doors",[]],["TB.PCM.Category.Exteriors","Exteriors",[["TB.PCM.Category.ColorSchemes","Color Schemes"]]],["TB.PCM.Category.Fireplaces","Fireplaces",[]],["TB.PCM.Category.FloorPlanOptions","Floor Plan Options",[["TB.PCM.Category.Interior","Interior"],["TB.PCM.Category.Outdoor","Outdoor"]]],["TB.PCM.Category.Flooring","Flooring",[["TB.PCM.Category.Hardwood","Hardwood"],["TB.PCM.Category.Flooring.Laminate","Laminate"]]],["TB.PCM.Category.InteriorD.GarageEpoxy","Garage Epoxy",[]],["TB.PCM.Category.HomeDetails","Home Details",[["TB.PCM.Category.BaseboardTrim","Baseboard & Trim"],["TB.PCM.Category.CeilingBeam","Ceiling Beam"],["TB.PCM.Category.Drywall","Drywall"],["TB.PCM.Category.Electrical","Electrical"],["TB.PCM.Category.EnergySolar","Energy & Solar"],["TB.PCM.Category.FireSuppression","Fire Suppression"],["TB.PCM.Category.HomeDet.Fireplace","Fireplace"],["TB.PCM.Category.GasSystems","Gas Systems"],["TB.PCM.Category.Gutters","Gutters"],["TB.PCM.Category.HVAC","HVAC"],["TB.PCM.Category.HumidifierSystems","Humidifier Systems"],["TB.PCM.Category.Insulation","Insulation"],["TB.PCM.Category.RoughPlumbing","Rough Plumbing"],["TB.PCM.Category.Tubs","Tubs"],["TB.PCM.Category.Wallfinish","Wall finish"],["TB.PCM.Category.WaterSystems","Water Systems"]]],["TB.PCM.Category.HomeTechnology","Home Technology",[["TB.PCM.Category.AudioVideo","Audio & Video"],["TB.PCM.Category.Network","Network"],["TB.PCM.Category.Security","Security"],["TB.PCM.Category.SmartLighting","Smart Lighting"]]],["TB.PCM.Category.InteriorDetails","Interior Details",[["TB.PCM.Category.Closets","Closets"],["TB.PCM.Category.DoorHardware","Door Hardware"],["TB.PCM.Category.InteriorD.Mirrors","Mirrors"],["TB.PCM.Category.Paint","Paint"],["TB.PCM.Category.Trim","Trim"],["TB.PCM.Category.WindowCoverings","Window Coverings"]]],["TB.PCM.Category.KitchenFaucets","Kitchen Faucets",[]],["TB.PCM.Category.Flooring.LVP","LVP",[]],["TB.PCM.Category.InteriorD.Lighting","Lighting",[]],["TB.PCM.Category.OutdoorOptions","Outdoor Options",[["TB.PCM.Category.DecksBalconies","Decks & Balconies"],["TB.PCM.Category.Fencing","Fencing"],["TB.PCM.Category.Hardscaping","Hardscaping"],["TB.PCM.Category.Landscaping","Landscaping"],["TB.PCM.Category.OutdoorAccessories","Outdoor Accessories"],["TB.PCM.Category.OutdoorFireplace","Outdoor Fireplace"],["TB.PCM.Category.OutdoorKitchen","Outdoor Kitchen"],["TB.PCM.Category.Pool","Pool"]]],["TB.PCM.Category.PerformanceShowering","Performance Showering",[]],["TB.PCM.Category.Plumbing","Plumbing",[["TB.PCM.Category.BathFaucets","Bath Faucets"],["TB.PCM.Category.BathSink","Bath Sink"],["TB.PCM.Category.EntertainmentFaucets","Entertainment Faucets"],["TB.PCM.Category.EntertainmentSink","Entertainment Sink"],["TB.PCM.Category.KitchenSink","Kitchen Sink"],["TB.PCM.Category.LaundryFaucets","Laundry Faucets"],["TB.PCM.Category.LaundrySink","Laundry Sink"],["TB.PCM.Category.ShowerPackage","Shower Package"],["TB.PCM.Category.SoapDispenser","Soap Dispenser"],["TB.PCM.Category.TubFaucets","Tub Faucets"],["TB.PCM.Category.TubShowerPackage","Tub/Shower Package"],["TB.PCM.Category.Plumbing.Tubs","Tubs"]]],["TB.PCM.Category.Plumbing.ShowerEnclosures","Shower Enclosures",[]],["TB.PCM.Category.Sinks","Sinks",[]],["TB.PCM.Category.SolidSurfaces","Solid Surfaces",[]],["TB.PCM.Category.Stairs","Stairs",[]],["TB.PCM.Category.Tiles","Tile",[]],["TB.PCM.Category.Plumbing.Toilets","Toilets",[]]];
// Active category tree (replaced by a live read from Content Hub when available).
let categories = CATEGORIES;
// Content Hub required fields for M.PCM.Product (validated on NEW records).
const REQUIRED_FIELDS = ['TB.PCM.ProductName', 'TB.PCM.Category', 'TB.PCM.Product.Manufacturer', 'Color', 'TB.PCM.Product.SKU'];

// Intake friendly label (normalized) -> Content Hub field.
const FIELD_MAP = {
  'product name': 'TB.PCM.ProductName',
  'sku': 'TB.PCM.Product.SKU',
  'toll sku': 'TB.PCM.TollSKU',
  'manufacturer': 'TB.PCM.Product.Manufacturer',
  'brand': 'TB.PCM.Brand',
  'family/style name': 'TB.PCM.FamilyName',
  'color': 'Color',
  'color family': 'TB.PCM.TileColorFamily',
  'look/style': 'TB.PCM.TileStyle',
  'shape': 'TB.PCM.TileSize',
  'size': 'TB.PCM.Size',
  'material': 'TB.PCM.TileMaterial',
  'finish': 'TB.PCM.TileFinish',
  'unit of measure': 'TB.PCM.Product.TileUnitOfMeasure',
  'thickness (in)': 'TB.PCM.Product.TileThickness'
};

// Item # columns, in the macro's concatenation order, with the AOA each implies.
const ITEM_COLS = [
  { label: 'floor item #',        aoa: 'TB.PCM.AreaOfApplication.FloorTile' },
  { label: 'wall item #',         aoa: 'TB.PCM.AreaOfApplication.WallTile' },
  { label: 'backsplash item #',   aoa: 'TB.PCM.AreaOfApplication.TileBacksplash' },
  { label: 'shower floor item #', aoa: 'TB.PCM.AreaOfApplication.ShowerFloor' },
  { label: 'listello item #',     aoa: 'TB.PCM.AreaOfApplication.Listello' }
];

const ID_LABELS = { 'id': 'id', 'content hub id': 'id', 'identifier': 'identifier', 'content hub identifier': 'identifier' };

const OUT_COLS = [
  'id', 'identifier', 'TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU',
  'TB.PCM.E1ItemNumber', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'Color',
  'TB.PCM.TileColorFamily', 'TB.PCM.TileStyle', 'TB.PCM.TileSize', 'TB.PCM.Size', 'TB.PCM.TileMaterial',
  'TB.PCM.TileFinish', 'TB.PCM.Product.TileUnitOfMeasure', 'TB.PCM.Product.TileThickness', 'TB.PCM.AreaOfApplication'
];
const TEXT_FIELDS = new Set(['id', 'identifier', 'TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU', 'TB.PCM.E1ItemNumber', 'TB.PCM.FamilyName', 'Color', 'TB.PCM.Size', 'TB.PCM.Product.TileThickness', 'TB.PCM.AreaOfApplication']);

const CSS = `
  .g-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 860px; }
  .g-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .g-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .g-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .g-drop.g-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .g-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; }
  .g-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .g-btn:disabled { opacity: .5; cursor: not-allowed; }
  .g-dry   { background: #edf2f7; color: #2d3748; }
  .g-go    { background: #2f855a; color: #fff; }
  .g-sync  { background: #6b46c1; color: #fff; }
  .g-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 360px; overflow: auto; white-space: pre-wrap; display: none; }
  .g-ok { color: #68d391; } .g-skip { color: #cbd5e0; } .g-err { color: #fc8181; } .g-info { color: #90cdf4; }
`;

function loadXLSX() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const s = document.createElement('script');
    s.src = SHEETJS_URL;
    s.onload = () => (window.XLSX ? resolve(window.XLSX) : reject(new Error('SheetJS loaded but window.XLSX missing.')));
    s.onerror = () => reject(new Error('Could not load SheetJS (check CSP / network).'));
    document.head.appendChild(s);
  });
}

function norm(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }

// [[identifier,label],...] -> Map of label/segment/full-id (lowercased) -> identifier
function buildLabelMap(pairs) {
  const m = new Map();
  for (const [identifier, label] of pairs) {
    if (!identifier) continue;
    m.set(identifier.toLowerCase(), identifier);
    const last = identifier.split('.').pop();
    if (last) m.set(last.toLowerCase(), identifier);
    if (label) m.set(String(label).trim().toLowerCase(), identifier);
  }
  return m;
}

// Resolve display value(s) -> identifiers. Returns { value, bad } where `bad`
// lists the tokens that could not be matched.
function resolveField(value, map) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return { value: '', bad: [] };
  const tokens = raw.split(/[|,]/).map(t => t.trim()).filter(Boolean);
  const ids = [], bad = [];
  for (const tok of tokens) {
    let id = map.get(tok.toLowerCase());
    if (!id && tok.indexOf('.') >= 0) id = tok; // already an identifier
    if (id) ids.push(id); else bad.push(tok);
  }
  return { value: ids.join('|'), bad };
}

// Read a culture-keyed label from an SDK data source value (Map or object).
function readProp(item, name) {
  const p = item && item.properties ? item.properties[name] : undefined;
  if (p == null) return '';
  if (typeof p === 'object') { const k = Object.keys(p); return k.length ? String(p[k[0]]) : ''; }
  return String(p);
}
function itemId(it) {
  if (it && it.id != null) return it.id;
  if (it && it.self && it.self.href) return Number(it.self.href.split('/').pop());
  return null;
}

// Read the TB.PCM.Category taxonomy live and build [ [catId, label, [[subId, subLabel]...]], ... ].
async function loadCategoriesLive() {
  const chql = `Definition.Name=='TB.PCM.Category'`;
  const items = []; let skip = 0, safety = 0;
  while (safety++ < 100) {
    const url = `${CH_HOST}/api/entities/query?query=${encodeURIComponent(chql)}&members=TaxonomyLabel&skip=${skip}&take=200`;
    const res = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const arr = ((await res.json()) || {}).items || [];
    if (!arr.length) break;
    items.push(...arr);
    if (arr.length < 200) break;
    skip += 200;
  }
  if (!items.length) throw new Error('no items');
  try { console.log('[category item sample]', JSON.stringify(items[0])); } catch (e) { /* noop */ }

  const byId = new Map(), byIdent = new Map();
  for (const it of items) {
    const id = itemId(it), identifier = it.identifier || '';
    const label = readProp(it, 'TaxonomyLabel') || readProp(it, 'Label') || identifier.split('.').pop();
    const node = { id, identifier, label, parentId: null, children: [] };
    if (id != null) byId.set(id, node);
    if (identifier) byIdent.set(identifier, node);
  }
  for (const it of items) {
    const rels = it.relations || {};
    const rel = rels.TB_PCM_CategoryToSelf || rels['TB.PCM.CategoryToSelf'] || rels.CategoryToSelf;
    let pref = rel && (rel.parent_id || rel.parent || (rel.parents && rel.parents[0]));
    if (pref == null) continue;
    const s = (typeof pref === 'object') ? (pref.href || pref.id || '') : String(pref);
    const tail = String(s).split('/').pop();
    let pid = null;
    if (tail.indexOf('.') >= 0 && byIdent.has(tail)) pid = byIdent.get(tail).id;
    else if (tail && !isNaN(Number(tail))) pid = Number(tail);
    const node = byId.get(itemId(it));
    if (node) node.parentId = pid;
  }
  const tops = [];
  for (const n of byId.values()) {
    if (n.parentId != null && byId.has(n.parentId)) byId.get(n.parentId).children.push(n);
    else tops.push(n);
  }
  const cmp = (a, b) => String(a.label).localeCompare(String(b.label));
  tops.sort(cmp); tops.forEach(t => t.children.sort(cmp));
  return tops.map(t => [t.identifier, t.label, t.children.map(c => [c.identifier, c.label])]);
}

function dsLabel(v, culture) {
  const L = v && (v.labels || v.Labels);
  if (!L) return '';
  if (typeof L.get === 'function') return L.get(culture) || L.get('en-US') || '';
  return L[culture] || L['en-US'] || '';
}

// Get one source's [[identifier, label], ...], reading live from Content Hub
// (cached for the session). Falls back to the embedded snapshot if the live
// read is unavailable or returns nothing.
async function getSourcePairs(source, client, culture, log) {
  if (liveCache[source]) return liveCache[source];
  if (client && client.dataSources && typeof client.dataSources.getAsync === 'function') {
    try {
      const ds = await client.dataSources.getAsync(source);
      const values = (ds && (ds.values || ds.Values)) || [];
      const pairs = [];
      for (const v of values) { const id = v.identifier || v.Identifier; if (id) pairs.push([id, dsLabel(v, culture)]); }
      if (pairs.length) { liveCache[source] = pairs; return pairs; }
      if (log) log(`  ${source}: 0 live values — using fallback snapshot`, 'g-skip');
    } catch (e) {
      if (log) log(`  ${source}: live read failed (${e && e.message ? e.message : e}) — using fallback`, 'g-skip');
    }
  }
  const fb = DEFAULT_LOOKUPS[source] || [];
  liveCache[source] = fb;
  return fb;
}

async function parseFileAOA(file, XLSX) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
}

function ts() { const d = new Date(), p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// Tile match key: concatenated Item # (E1) + Color, case-insensitive.
function matchKey(e1, color) {
  return (String(e1 == null ? '' : e1).trim() + '|' + String(color == null ? '' : color).trim()).toLowerCase();
}

// Index a Content Hub export (headers = CH field names) by match key -> {id, identifier}.
function buildExportIndex(aoa) {
  if (!aoa || aoa.length < 2) return { error: 'Content Hub export has no data rows.' };
  const headers = aoa[0].map(h => String(h == null ? '' : h).trim().toLowerCase());
  const col = name => headers.indexOf(name);
  const idCol = col('id'), identCol = col('identifier'), e1Col = col('tb.pcm.e1itemnumber'), colorCol = col('color');
  if (e1Col < 0 || colorCol < 0) return { error: 'Export must include TB.PCM.E1ItemNumber and Color columns.' };
  if (idCol < 0 && identCol < 0) return { error: 'Export must include an id or identifier column.' };
  const idx = new Map();
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i];
    const e1 = row[e1Col], color = row[colorCol];
    if (!String(e1 == null ? '' : e1).trim() && !String(color == null ? '' : color).trim()) continue;
    const key = matchKey(e1, color);
    if (idx.has(key)) continue;
    const values = {};
    for (let j = 0; j < headers.length; j++) values[headers[j]] = row[j]; // headers are lowercased
    idx.set(key, {
      id: idCol >= 0 ? row[idCol] : '', identifier: identCol >= 0 ? row[identCol] : '',
      e1: String(e1 == null ? '' : e1).trim(), color: String(color == null ? '' : color).trim(),
      values
    });
  }
  return { idx };
}

function buildRecord(rowObj) {
  const rec = {};
  for (const [lbl, chf] of Object.entries(ID_LABELS)) {
    const v = (rowObj[lbl] || '').trim();
    if (v && !rec[chf]) rec[chf] = v;
  }
  const e1 = [], aoa = [];
  for (const it of ITEM_COLS) {
    const v = (rowObj[it.label] || '').trim();
    if (v) { e1.push(v); aoa.push(it.aoa); }
  }
  if (e1.length) rec['TB.PCM.E1ItemNumber'] = e1.join(',');
  if (aoa.length) rec['TB.PCM.AreaOfApplication'] = aoa.join('|');
  for (const [lbl, chf] of Object.entries(FIELD_MAP)) {
    const v = (rowObj[lbl] || '').trim();
    if (v) rec[chf] = v;
  }
  // Potential fallbacks for required fields (applied later, NEW records only).
  const d1 = (rowObj['description 1'] || '').trim();
  const d2 = (rowObj['description 2'] || '').trim();
  rec.__fallbackName = d1 || d2;
  rec.__fallbackSku = ['backsplash item #', 'floor item #', 'shower floor item #', 'wall item #', 'listello item #']
    .map(l => (rowObj[l] || '').trim()).filter(Boolean).join(',');
  return rec;
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const client = context && context.client;
      const cfg = (context && context.config) || {};
      const culture = (context && (context.culture || (context.options && context.options.culture))) || 'en-US';
      const style = document.createElement('style'); style.textContent = CSS;
      const wrap = document.createElement('div'); wrap.className = 'g-wrap';
      wrap.innerHTML = `
        <div class="g-title">🏭 Product Import Generator — Tile</div>
        <div class="g-sub">Upload the vendor <b>intake</b> spreadsheet (friendly Tile headers).
          The tool builds the Item Number + Area of Application, resolves option-list values to
          identifiers, and downloads a ready-to-import <b>${SHEET_NAME}</b> workbook. Add
          <code>id</code> and <code>identifier</code> columns for updates; leave them out for new records.</div>
        <div class="g-row" style="margin-bottom:14px">
          <label style="font-size:13px;color:#555">Category:</label>
          <select id="g-cat" style="padding:6px;border:1px solid #cbd5e0;border-radius:6px;font-size:13px"></select>
          <label style="font-size:13px;color:#555">Sub-category:</label>
          <select id="g-subcat" style="padding:6px;border:1px solid #cbd5e0;border-radius:6px;font-size:13px"></select>
          <span id="g-catstatus" style="font-size:11px;color:#888"></span>
        </div>
        <div class="g-drop" id="g-drop">📎 <b>1. Intake file</b> — drop your vendor .xlsx / .csv here, or click to browse</div>
        <input type="file" id="g-file" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-drop" id="g-drop2">🔁 <b>2. Content Hub export</b> (optional — only for UPDATES) — drop the export with id/identifier</div>
        <input type="file" id="g-file2" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-row">
          <button class="g-btn g-dry" id="g-dry" disabled>🔍 Validate (dry run)</button>
          <button class="g-btn g-go"  id="g-go"  disabled>⬇ Generate import file</button>
          <span id="g-status" style="font-size:13px;color:#555"></span>
        </div>
        <div style="font-size:12px;color:#888;margin-bottom:6px">
          No export = create new records. With export = update: intake rows are matched by Item # + Color to pull id/identifier.
        </div>
        <div class="g-row">
          <button class="g-btn g-sync" id="g-sync">🔄 Reload option lists from Content Hub</button>
          <span id="g-lookups" style="font-size:12px;color:#888"></span>
        </div>
        <div class="g-log" id="g-log"></div>
      `;
      rootElement.innerHTML = ''; rootElement.appendChild(style); rootElement.appendChild(wrap);

      const drop = wrap.querySelector('#g-drop'), input = wrap.querySelector('#g-file');
      const drop2 = wrap.querySelector('#g-drop2'), input2 = wrap.querySelector('#g-file2');
      const catSel = wrap.querySelector('#g-cat'), subSel = wrap.querySelector('#g-subcat');
      const catStatus = wrap.querySelector('#g-catstatus');

      function fillSubcategories() {
        const entry = categories.find(c => c[0] === catSel.value);
        const kids = entry ? entry[2] : [];
        if (!kids.length) {
          subSel.innerHTML = '<option value="">(no sub-categories)</option>';
          subSel.disabled = true;
        } else {
          subSel.disabled = false;
          subSel.innerHTML = '<option value="">— use category —</option>' +
            kids.map(([id, label]) => `<option value="${id}">${label}</option>`).join('');
        }
      }
      function fillCategories() {
        const prev = catSel.value;
        catSel.innerHTML = categories.map(([id, label]) => `<option value="${id}">${label}</option>`).join('');
        catSel.value = categories.some(c => c[0] === prev) ? prev : (categories.some(c => c[0] === DEFAULT_CATEGORY_ID) ? DEFAULT_CATEGORY_ID : (categories[0] && categories[0][0]));
        fillSubcategories();
      }
      fillCategories();
      catSel.addEventListener('change', fillSubcategories);

      // Load the current category tree from Content Hub; fall back to the snapshot.
      catStatus.textContent = 'categories: snapshot';
      (async () => {
        try {
          const live = await loadCategoriesLive();
          if (live && live.length) { categories = live; fillCategories(); catStatus.textContent = `categories: live (${live.length})`; }
        } catch (e) {
          catStatus.textContent = `categories: snapshot — live load failed (${e && e.message ? e.message : e})`;
        }
      })();
      const dryBtn = wrap.querySelector('#g-dry'), goBtn = wrap.querySelector('#g-go');
      const syncBtn = wrap.querySelector('#g-sync'), lookupsStatus = wrap.querySelector('#g-lookups');
      const status = wrap.querySelector('#g-status'), logEl = wrap.querySelector('#g-log');
      let currentFile = null, currentExport = null;

      const lookupCount = () => Object.values(liveCache).reduce((a, v) => a + (v ? v.length : 0), 0);
      function setLookupsStatus(note) { lookupsStatus.textContent = `Option lists: ${lookupCount()} values ${note}`; }
      setLookupsStatus('— read live from Content Hub on first run.');

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div'); if (cls) line.className = cls;
        line.textContent = msg; logEl.appendChild(line); logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }

      function refreshStatus() {
        const parts = [];
        if (currentFile) parts.push(`intake: ${currentFile.name}`);
        if (currentExport) parts.push(`export: ${currentExport.name} (UPDATE)`);
        status.textContent = parts.join('  ·  ');
        dryBtn.disabled = !currentFile; goBtn.disabled = !currentFile;
      }
      function onFile(file) { if (!file) return; clearLog(); currentFile = file; refreshStatus(); }
      function onExport(file) { if (!file) return; clearLog(); currentExport = file; refreshStatus(); }

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('g-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('g-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('g-hover'); onFile(e.dataTransfer.files[0]); });

      drop2.addEventListener('click', () => input2.click());
      input2.addEventListener('change', e => onExport(e.target.files[0]));
      drop2.addEventListener('dragover', e => { e.preventDefault(); drop2.classList.add('g-hover'); });
      drop2.addEventListener('dragleave', () => drop2.classList.remove('g-hover'));
      drop2.addEventListener('drop', e => { e.preventDefault(); drop2.classList.remove('g-hover'); onExport(e.dataTransfer.files[0]); });

      async function run(dryRun) {
        clearLog(); dryBtn.disabled = true; goBtn.disabled = true;
        log(dryRun ? '── VALIDATE (no file written) ──' : '── GENERATE IMPORT FILE ──', 'g-info');
        try {
          const XLSX = await loadXLSX();
          const aoa = await parseFileAOA(currentFile, XLSX);
          if (!aoa || aoa.length < 2) { log('No data rows found.', 'g-err'); return; }

          const rawHeaders = aoa[0].map(h => String(h == null ? '' : h).trim());
          const normHeaders = rawHeaders.map(norm);

          const known = new Set([...Object.keys(FIELD_MAP), ...Object.keys(ID_LABELS), ...ITEM_COLS.map(i => i.label)]);
          const skipped = rawHeaders.filter((h, i) => h && !known.has(normHeaders[i]));

          // Build records, keeping the spreadsheet row number (header = row 1).
          const records = [];
          for (let i = 1; i < aoa.length; i++) {
            const row = aoa[i];
            if (!row.some(c => String(c == null ? '' : c).trim() !== '')) continue;
            const obj = {};
            for (let j = 0; j < normHeaders.length; j++) if (normHeaders[j]) obj[normHeaders[j]] = row[j];
            const rec = buildRecord(obj);
            rec.__row = i + 1;
            records.push(rec);
          }

          // Category from the dropdowns: use the sub-category (leaf) if chosen, else the category.
          const catValue = (subSel.value || catSel.value || '');
          for (const r of records) r['TB.PCM.Category'] = catValue;

          log(`Data rows: ${records.length}. Recognized columns: ${rawHeaders.filter((h, i) => known.has(normHeaders[i])).length}.`, 'g-info');
          log(`Category: ${catValue}`, 'g-info');
          if (skipped.length) log(`Skipped (reference-only): ${skipped.join(', ')}`, 'g-skip');

          // UPDATE mode: match intake rows to a Content Hub export by (E1 Item # + Color).
          const updateMode = !!currentExport;
          let outputRecords = records;
          if (updateMode) {
            const expAoa = await parseFileAOA(currentExport, XLSX);
            const { idx, error } = buildExportIndex(expAoa);
            if (error) { log(`✗ Content Hub export: ${error}`, 'g-err'); return; }
            log(`UPDATE mode — Content Hub export: ${idx.size} record(s) indexed by (Item # + Color).`, 'g-info');
            let matched = 0; const unmatched = [];
            for (const r of records) {
              const ik = matchKey(r['TB.PCM.E1ItemNumber'], r['Color']);
              const m = idx.get(ik);
              if (m) {
                if (m.id != null && String(m.id).trim()) r['id'] = m.id;
                if (m.identifier != null && String(m.identifier).trim()) r['identifier'] = m.identifier;
                r.__matched = true; r.__export = m.values; matched++;
              } else {
                r.__matched = false;
                unmatched.push({ row: r.__row, e1: r['TB.PCM.E1ItemNumber'] || '', color: r['Color'] || '' });
              }
            }
            log(`Matched ${matched} of ${records.length} intake row(s) to existing records.`, matched ? 'g-ok' : 'g-err');
            if (unmatched.length) {
              log(`Unmatched (skipped) — intake values that found no export match:`, 'g-err');
              unmatched.forEach(u => log(`   Row ${u.row}: Item#="${u.e1}"  Color="${u.color}"`, 'g-err'));
              const sample = [...idx.values()].slice(0, 15);
              log(`Content Hub export contains (Item# / Color) — compare against the above:`, 'g-skip');
              sample.forEach(e => log(`   Item#="${e.e1}"  Color="${e.color}"`, 'g-skip'));
              if (idx.size > sample.length) log(`   … and ${idx.size - sample.length} more`, 'g-skip');
            }
            outputRecords = records.filter(r => r.__matched);
            if (outputRecords.length === 0) { log('Nothing to update — no intake rows matched the export.', 'g-err'); return; }
          }

          // NEW records only: fill required ProductName/SKU from fallbacks if absent.
          if (!updateMode) {
            for (const r of outputRecords) {
              if (!String(r['TB.PCM.ProductName'] == null ? '' : r['TB.PCM.ProductName']).trim() && r.__fallbackName) r['TB.PCM.ProductName'] = r.__fallbackName;
              if (!String(r['TB.PCM.Product.SKU'] == null ? '' : r['TB.PCM.Product.SKU']).trim() && r.__fallbackSku) r['TB.PCM.Product.SKU'] = r.__fallbackSku;
            }
          }

          const used = new Set();
          outputRecords.forEach(r => Object.keys(r).forEach(k => { if (!k.startsWith('__')) used.add(k); }));

          // Build option-list maps from the embedded snapshot.
          const meta = {};
          for (const f of OUT_COLS) {
            if (!used.has(f) || TEXT_FIELDS.has(f)) { meta[f] = { isOption: false }; continue; }
            const src = sourceFor(f);
            const pairs = await getSourcePairs(src, client, culture, log);
            if (pairs && pairs.length) {
              meta[f] = { isOption: true, map: buildLabelMap(pairs) };
              log(`  ${f} → ${src}: ${pairs.length} options`, 'g-info');
            } else {
              meta[f] = { isOption: false };
              log(`  ${f}: no option list found (${src}) — passed through as text`, 'g-skip');
            }
          }

          const unresolved = [];
          for (const r of outputRecords) {
            for (const f of Object.keys(r)) {
              if (f.startsWith('__')) continue;
              if (meta[f] && meta[f].isOption) {
                const res = resolveField(r[f], meta[f].map);
                res.bad.forEach(b => unresolved.push({ row: r.__row, field: f, value: b }));
                if (res.bad.length && updateMode) {
                  const orig = r.__export ? r.__export[f.toLowerCase()] : '';   // keep existing CH value
                  r[f] = (orig == null ? '' : orig);
                } else {
                  r[f] = res.value;   // create mode: blank if unresolved
                }
              }
            }
          }

          if (unresolved.length === 0) {
            log('All option-list values matched. ✓', 'g-ok');
          } else if (updateMode) {
            log(`⚠ ${unresolved.length} value(s) not found — the existing Content Hub value was kept for these:`, 'g-err');
            unresolved.slice(0, 100).forEach(u => log(`   Row ${u.row} · ${u.field}: "${u.value}" (kept original)`, 'g-err'));
            if (unresolved.length > 100) log(`   … and ${unresolved.length - 100} more`, 'g-err');
          } else {
            log(`⚠ ${unresolved.length} value(s) not found — these cells will be left BLANK in the output:`, 'g-err');
            unresolved.slice(0, 100).forEach(u => log(`   Row ${u.row} · ${u.field}: "${u.value}"`, 'g-err'));
            if (unresolved.length > 100) log(`   … and ${unresolved.length - 100} more`, 'g-err');
          }

          // Required-field check (new records only; updates inherit existing values).
          if (!updateMode) {
            const missingReport = [];
            for (const r of outputRecords) {
              const miss = REQUIRED_FIELDS.filter(f => !String(r[f] == null ? '' : r[f]).trim());
              if (miss.length) missingReport.push({ row: r.__row, miss });
            }
            if (missingReport.length) {
              log(`⚠ Required field(s) missing on ${missingReport.length} new record(s) — Content Hub will reject these until filled:`, 'g-err');
              missingReport.slice(0, 100).forEach(m => log(`   Row ${m.row}: ${m.miss.join(', ')}`, 'g-err'));
            } else {
              log('All required fields present. ✓', 'g-ok');
            }
          }

          if (dryRun) { log('Dry run complete — review the items above, then Generate.', 'g-info'); return; }

          const outCols = OUT_COLS.filter(f => used.has(f));
          const outRows = [outCols];
          for (const r of outputRecords) outRows.push(outCols.map(f => (r[f] == null ? '' : r[f])));

          const ws = XLSX.utils.aoa_to_sheet(outRows);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
          const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const fname = `ContentHub_TilesImport_${ts()}.xlsx`;
          downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), fname);
          log(`✓ Generated ${fname} — ${outputRecords.length} ${updateMode ? 'update' : 'new'} row(s), sheet "${SHEET_NAME}".`, 'g-ok');
          if (unresolved.length) log(updateMode
            ? `⚠ ${unresolved.length} unmatched value(s) kept their original Content Hub value (listed above).`
            : `⚠ ${unresolved.length} unmatched value(s) were left blank (listed above).`, 'g-err');
        } catch (e) {
          log(`✗ ${e && e.message ? e.message : e}`, 'g-err');
        } finally {
          dryBtn.disabled = false; goBtn.disabled = false;
        }
      }

      dryBtn.addEventListener('click', () => run(true));
      goBtn.addEventListener('click', () => run(false));

      syncBtn.addEventListener('click', async () => {
        clearLog(); syncBtn.disabled = true;
        log('── RELOAD OPTION LISTS FROM CONTENT HUB ──', 'g-info');
        liveCache = {}; // clear the session cache so values are re-fetched
        let total = 0;
        for (const name of DATA_SOURCES) {
          const pairs = await getSourcePairs(name, client, culture, log);
          total += pairs.length;
          log(`  ${name}: ${pairs.length} value(s)`, pairs.length ? 'g-ok' : 'g-err');
        }
        setLookupsStatus('— reloaded live from Content Hub.');
        log(`✓ Loaded ${total} value(s) across ${DATA_SOURCES.length} option list(s).`, 'g-ok');
        syncBtn.disabled = false;
      });
    },
    unmount() { rootElement.innerHTML = ''; }
  };
}
