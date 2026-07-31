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

// Option-list snapshot: source -> [[identifier, displayLabel], ...]
const LOOKUPS = {"TB.PCM.Manufacturer":[["TB.PCM.Manufacturer.Amerock","Amerock"],["TB.PCM.Manufacturer.Anderson","Anderson"],["TB.PCM.Manufacturer.ArizonaTile","Arizona Tile"],["TB.PCM.Manufacturer.BerensonHardware","Berenson Hardware"],["TB.PCM.Manufacturer.CaliforniaMantelFireplace","California Mantel & Fireplace"],["TB.PCM.Manufacturer.Consentino","Consentino"],["TB.PCM.Manufacturer.Cove","Cove"],["TB.PCM.Manufacturer.DalTile","DalTile"],["TB.PCM.Manufacturer.ElectricMirror","Electric Mirror"],["TB.PCM.Manufacturer.Emser","Emser"],["TB.PCM.Manufacturer.Emtek","Emtek"],["TB.PCM.Manufacturer.HomeSiteServices","Home Site Services"],["TB.PCM.Manufacturer.HunterDouglas","Hunter Douglas"],["TB.PCM.Manufacturer.Infratech","Infratech"],["TB.PCM.Manufacturer.JamisonCollections","Jamison Collections"],["TB.PCM.Manufacturer.JennAir","JennAir"],["TB.PCM.Manufacturer.KitchenAid","KitchenAid"],["TB.PCM.Manufacturer.Kohler","Kohler"],["TB.PCM.Manufacturer.Kwikset","Kwikset"],["TB.PCM.Manufacturer.Marazzi","Marazzi"],["TB.PCM.Manufacturer.Metropolitan","Metropolitan"],["TB.PCM.Manufacturer.Mohawk","Mohawk"],["TB.PCM.Manufacturer.PrecisionCabinetry","Precision Cabinetry"],["TB.PCM.Manufacturer.ProgressLighting","Progress Lighting"],["TB.PCM.Manufacturer.Provenza","Provenza"],["TB.PCM.Manufacturer.Shaw","Shaw"],["TB.PCM.Manufacturer.SherwinWilliams","Sherwin Williams"],["TB.PCM.Manufacturer.Sterling","Sterling"],["TB.PCM.Manufacturer.SubZero","Sub-Zero"],["TB.PCM.Manufacturer.TopKnobs","Top Knobs"],["TB.PCM.Manufacturer.Tuftex","Tuftex"],["TB.PCM.Manufacturer.Unbranded","Unbranded"],["TB.PCM.Manufacturer.WesternWindowSystems","Western Window Systems"],["TB.PCM.Manufacturer.Whirlpool","Whirlpool"],["TB.PCM.Manufacturer.Wolf","Wolf"],["TB.PCM.Manufacturer.CenturyCabinetry","Century Cabinetry"],["TB.PCM.Manufacturer.AmericanOlean","American Olean"],["TB.PCM.Manufacturer.UrbanEffects","Urban Effects"],["TB.PCM.Manufacturer.Coyote","Coyote"],["TB.PCM.Manufacturer.MSI","MSI"],["TB.PCM.Manufacturer.ASKO","ASKO"],["TB.PCM.Manufacturer.Mullican","Mullican"],["TB.PCM.Manufacturer.Mannington","Mannington"]],"TB.PCM.Brand":[["TB.PCM.Brand.AmericanOlean","American Olean"],["TB.PCM.Brand.ArizonaTile","Arizona Tile"],["TB.PCM.Brand.Daltile","Daltile"],["TB.PCM.Brand.EmserTile","Emser Tile"],["TB.PCM.Brand.Marazzi","Marazzi"],["TB.PCM.Brand.Shaw","Shaw"]],"TB.PCM.TileColorFamily":[["TB.PCM.TileColorFamily.Beige","Beige"],["TB.PCM.TileColorFamily.Black","Black"],["TB.PCM.TileColorFamily.Blue","Blue"],["TB.PCM.TileColorFamily.Brown","Brown"],["TB.PCM.TileColorFamily.Cream","Cream"],["TB.PCM.TileColorFamily.Gray","Gray"],["TB.PCM.TileColorFamily.Green","Green"],["TB.PCM.TileColorFamily.Multicolor","Multicolor"],["TB.PCM.TileColorFamily.Orange","Orange"],["TB.PCM.TileColorFamily.Pink","Pink"],["TB.PCM.TileColorFamily.Red","Red"],["TB.PCM.TileColorFamily.Taupe","Taupe"],["TB.PCM.TileColorFamily.TerraCotta","Terra Cotta"],["TB.PCM.TileColorFamily.White","White"],["TB.PCM.TileColorFamily.Yellow","Yellow"]],"TB.PCM.TileStyle":[["TB.PCM.TileStyle.Brick","Brick"],["TB.PCM.TileStyle.ConcreteLook","Concrete Look"],["TB.PCM.TileStyle.Fabric","Fabric"],["TB.PCM.TileStyle.Handmade","Handmade"],["TB.PCM.TileStyle.MarbleLook","Marble Look"],["TB.PCM.TileStyle.Mosaics","Mosaics"],["TB.PCM.TileStyle.NaturalStone","Natural Stone"],["TB.PCM.TileStyle.Patterned","Patterned"],["TB.PCM.TileStyle.Solid","Solid"],["TB.PCM.TileStyle.StoneLook","Stone Look"],["TB.PCM.TileStyle.Textured","Textured"],["TB.PCM.TileStyle.WoodLook","Wood Look"]],"TB.PCM.TileSize":[["TB.PCM.TileSize.HerringboneChevron","Herringbone/Chevron"],["TB.PCM.TileSize.Hexagon","Hexagon"],["TB.PCM.TileSize.LargeFormat","Large Format"],["TB.PCM.TileSize.Pebble","Pebble"],["TB.PCM.TileSize.PennyOval","Penny/Oval"],["TB.PCM.TileSize.Plank","Plank"],["TB.PCM.TileSize.Rectangle","Rectangle"],["TB.PCM.TileSize.Square","Square"],["TB.PCM.TileSize.Subway","Subway"],["TB.PCM.TileSize.UniqueShapes","Unique Shapes"]],"TB.PCM.TileMaterial":[["TB.PCM.TileMaterial.Ceramic","Ceramic"],["TB.PCM.TileMaterial.Porcelain","Porcelain"],["TB.PCM.TileMaterial.Glass","Glass"],["TB.PCM.TileMaterial.NaturalStoneMarble","Natural Stone - Marble"],["TB.PCM.TileMaterial.NaturalStoneTravertine","Natural Stone - Travertine"],["TB.PCM.TileMaterial.NaturalStoneLimestone","Natural Stone - Limestone"],["TB.PCM.TileMaterial.NaturalStoneMixed","Natural Stone - Mixed"]],"TB.PCM.TileFinish":[["TB.PCM.TileFinish.3DSculptural","3D/Sculptural"],["TB.PCM.TileFinish.Glossy","Glossy"],["TB.PCM.TileFinish.Matte","Matte"],["TB.PCM.TileFinish.Satin","Satin"],["TB.PCM.TileFinish.SlipResistant","Slip Resistant"],["TB.PCM.TileFinish.Smooth","Smooth"],["TB.PCM.TileFinish.Textured","Textured"]],"TB.PCM.UnitOfMeasure":[["TB.PCM.UnitOfMeasure.SQF","SQF"],["TB.PCM.UnitOfMeasure.Each","Each"]]};

// Field -> backing option-list source (mirrors GetLookupSourceForCHField).
const FIELD_SOURCE = {
  'TB.PCM.Product.Manufacturer': 'TB.PCM.Manufacturer',
  'TB.PCM.Product.TileUnitOfMeasure': 'TB.PCM.UnitOfMeasure',
  'TB.PCM.DivisionSelected': 'TB.Division'
};
function sourceFor(field) { return FIELD_SOURCE[field] || field; }

// Intake friendly label (normalized) -> Content Hub field.
const FIELD_MAP = {
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
  'id', 'identifier', 'TB.PCM.E1ItemNumber', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand',
  'TB.PCM.FamilyName', 'Color', 'TB.PCM.TileColorFamily', 'TB.PCM.TileStyle', 'TB.PCM.TileSize',
  'TB.PCM.Size', 'TB.PCM.TileMaterial', 'TB.PCM.TileFinish', 'TB.PCM.Product.TileUnitOfMeasure',
  'TB.PCM.Product.TileThickness', 'TB.PCM.AreaOfApplication'
];
const TEXT_FIELDS = new Set(['id', 'identifier', 'TB.PCM.E1ItemNumber', 'TB.PCM.FamilyName', 'Color', 'TB.PCM.Size', 'TB.PCM.Product.TileThickness', 'TB.PCM.AreaOfApplication']);

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

function resolveValue(value, map, field, rowNum, unresolved) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const tokens = raw.split(/[|,]/).map(t => t.trim()).filter(Boolean);
  const ids = [];
  for (const tok of tokens) {
    let id = map.get(tok.toLowerCase());
    if (!id && tok.indexOf('.') >= 0) id = tok; // already an identifier
    if (id) ids.push(id);
    else unresolved.push({ row: rowNum, field, value: tok }); // omit -> leave blank
  }
  return ids.join('|'); // blank if nothing resolved
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
  return rec;
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render() {
      const style = document.createElement('style'); style.textContent = CSS;
      const wrap = document.createElement('div'); wrap.className = 'g-wrap';
      wrap.innerHTML = `
        <div class="g-title">🏭 Product Import Generator — Tile</div>
        <div class="g-sub">Upload the vendor <b>intake</b> spreadsheet (friendly Tile headers).
          The tool builds the Item Number + Area of Application, resolves option-list values to
          identifiers, and downloads a ready-to-import <b>${SHEET_NAME}</b> workbook. Add
          <code>id</code> and <code>identifier</code> columns for updates; leave them out for new records.</div>
        <div class="g-drop" id="g-drop">📎 Drop your intake .xlsx / .csv here, or click to browse</div>
        <input type="file" id="g-file" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-row">
          <button class="g-btn g-dry" id="g-dry" disabled>🔍 Validate (dry run)</button>
          <button class="g-btn g-go"  id="g-go"  disabled>⬇ Generate import file</button>
          <span id="g-status" style="font-size:13px;color:#555"></span>
        </div>
        <div class="g-log" id="g-log"></div>
      `;
      rootElement.innerHTML = ''; rootElement.appendChild(style); rootElement.appendChild(wrap);

      const drop = wrap.querySelector('#g-drop'), input = wrap.querySelector('#g-file');
      const dryBtn = wrap.querySelector('#g-dry'), goBtn = wrap.querySelector('#g-go');
      const status = wrap.querySelector('#g-status'), logEl = wrap.querySelector('#g-log');
      let currentFile = null;

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div'); if (cls) line.className = cls;
        line.textContent = msg; logEl.appendChild(line); logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }

      function onFile(file) {
        if (!file) return;
        clearLog(); currentFile = file;
        status.textContent = `${file.name} — ready`;
        dryBtn.disabled = false; goBtn.disabled = false;
      }
      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('g-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('g-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('g-hover'); onFile(e.dataTransfer.files[0]); });

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

          log(`Data rows: ${records.length}. Recognized columns: ${rawHeaders.filter((h, i) => known.has(normHeaders[i])).length}.`, 'g-info');
          if (skipped.length) log(`Skipped (reference-only): ${skipped.join(', ')}`, 'g-skip');

          const used = new Set();
          records.forEach(r => Object.keys(r).forEach(k => { if (k !== '__row') used.add(k); }));

          // Build option-list maps from the embedded snapshot.
          const meta = {};
          for (const f of OUT_COLS) {
            if (!used.has(f) || TEXT_FIELDS.has(f)) { meta[f] = { isOption: false }; continue; }
            const src = sourceFor(f);
            const pairs = LOOKUPS[src];
            if (pairs && pairs.length) {
              meta[f] = { isOption: true, map: buildLabelMap(pairs) };
              log(`  ${f} → ${src}: ${pairs.length} options`, 'g-info');
            } else {
              meta[f] = { isOption: false };
              log(`  ${f}: no option list found (${src}) — passed through as text`, 'g-skip');
            }
          }

          const unresolved = [];
          for (const r of records) {
            for (const f of Object.keys(r)) {
              if (f === '__row') continue;
              if (meta[f] && meta[f].isOption) r[f] = resolveValue(r[f], meta[f].map, f, r.__row, unresolved);
            }
          }

          if (unresolved.length === 0) {
            log('All option-list values matched. ✓', 'g-ok');
          } else {
            log(`⚠ ${unresolved.length} value(s) not found — these cells will be left BLANK in the output:`, 'g-err');
            unresolved.slice(0, 100).forEach(u => log(`   Row ${u.row} · ${u.field}: "${u.value}"`, 'g-err'));
            if (unresolved.length > 100) log(`   … and ${unresolved.length - 100} more`, 'g-err');
          }

          if (dryRun) { log('Dry run complete — fix the values above (or leave them blank), then Generate.', 'g-info'); return; }

          const outCols = OUT_COLS.filter(f => used.has(f));
          const outRows = [outCols];
          for (const r of records) outRows.push(outCols.map(f => (r[f] == null ? '' : r[f])));

          const ws = XLSX.utils.aoa_to_sheet(outRows);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
          const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const fname = `ContentHub_TilesImport_${ts()}.xlsx`;
          downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), fname);
          log(`✓ Generated ${fname} — ${records.length} row(s), sheet "${SHEET_NAME}".`, 'g-ok');
          if (unresolved.length) log(`⚠ ${unresolved.length} unmatched value(s) were left blank (listed above).`, 'g-err');
        } catch (e) {
          log(`✗ ${e && e.message ? e.message : e}`, 'g-err');
        } finally {
          dryBtn.disabled = false; goBtn.disabled = false;
        }
      }

      dryBtn.addEventListener('click', () => run(true));
      goBtn.addEventListener('click', () => run(false));
    },
    unmount() { rootElement.innerHTML = ''; }
  };
}
