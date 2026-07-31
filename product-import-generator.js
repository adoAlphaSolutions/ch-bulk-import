// ============================================================================
// Toll Product Import Generator — Content Hub External Component (Tile v1)
// ----------------------------------------------------------------------------
// Browser replacement for the Tile macro's Intake -> "Export to Content Hub File".
//
// Input : the vendor INTAKE spreadsheet. Header row (row 1) uses the friendly
//         Tile intake labels, e.g.:
//           Backsplash Item # | Floor Item # | Shower Floor Item # | Wall Item # |
//           Listello Item # | Manufacturer | Brand | Family/Style Name | Color |
//           Color Family | Look/Style | Shape | Size | Material | Finish |
//           Unit of Measure | Thickness (in) | Backsplash Tile | Floor Tile |
//           Shower Floor | Wall Tile | Listello | Other Details | ... (reference)
//         Optionally include "id" and "identifier" columns for updates.
//
// Transform (mirrors Toll_PCM_Tile_Macros.bas):
//   * The 5 Item # cells concatenate (Floor,Wall,Backsplash,ShowerFloor,Listello)
//     with commas into TB.PCM.E1ItemNumber.
//   * Each filled Item # adds an AreaOfApplication identifier (pipe-joined):
//       Floor->FloorTile Wall->WallTile Backsplash->TileBacksplash
//       ShowerFloor->ShowerFloor Listello->Listello
//   * Friendly labels map to CH fields; option-list values are resolved to
//     identifiers LIVE from Content Hub; reference columns are skipped.
//   * id / identifier pass through when present (updates); blank => new records.
//
// Output: a downloadable workbook with ONE sheet named "M.PCM.Product".
// This tool only READS taxonomies and WRITES a local file.
// ============================================================================

const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const CH_HOST = window.location.origin;
const SHEET_NAME = 'M.PCM.Product';

// Field -> backing taxonomy (mirrors GetLookupSourceForCHField).
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

// Output column order for the M.PCM.Product sheet.
const OUT_COLS = [
  'id', 'identifier', 'TB.PCM.E1ItemNumber', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand',
  'TB.PCM.FamilyName', 'Color', 'TB.PCM.TileColorFamily', 'TB.PCM.TileStyle', 'TB.PCM.TileSize',
  'TB.PCM.Size', 'TB.PCM.TileMaterial', 'TB.PCM.TileFinish', 'TB.PCM.Product.TileUnitOfMeasure',
  'TB.PCM.Product.TileThickness', 'TB.PCM.AreaOfApplication'
];
// Never resolved as option-lists (plain text / already identifiers).
const TEXT_FIELDS = new Set(['id', 'identifier', 'TB.PCM.E1ItemNumber', 'TB.PCM.FamilyName', 'Color', 'TB.PCM.Product.TileThickness']);

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

function readProp(item, name) {
  const p = item && item.properties ? item.properties[name] : undefined;
  if (p == null) return '';
  if (typeof p === 'object') { const k = Object.keys(p); return k.length ? String(p[k[0]]) : ''; }
  return String(p);
}
function readLabel(it) { return readProp(it, 'Label') || readProp(it, 'TaxonomyLabel') || readProp(it, 'Name') || ''; }

async function queryDefinitionValues(defName) {
  const chql = `Definition.Name=='${defName}'`;
  const out = []; let skip = 0, safety = 0;
  while (safety++ < 500) {
    const url = `${CH_HOST}/api/entities/query?query=${encodeURIComponent(chql)}&skip=${skip}&take=200`;
    const res = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
    if (!res.ok) { const b = await res.text().catch(() => ''); throw new Error(`HTTP ${res.status}${b ? ' — ' + b : ''}`); }
    const data = await res.json();
    const items = (data && data.items) || [];
    if (items.length === 0) break;
    for (const it of items) {
      const identifier = it.identifier || (it.self && it.self.href ? it.self.href.split('/').pop() : null);
      if (identifier) out.push({ identifier, label: readLabel(it) });
    }
    if (items.length < 200) break;
    skip += 200;
  }
  return out;
}

function buildLabelMap(values) {
  const m = new Map();
  for (const v of values) {
    if (!v.identifier) continue;
    m.set(v.identifier.toLowerCase(), v.identifier);
    const last = v.identifier.split('.').pop();
    if (last) m.set(last.toLowerCase(), v.identifier);
    if (v.label) m.set(String(v.label).trim().toLowerCase(), v.identifier);
  }
  return m;
}

function resolveValue(value, map, field, unresolved) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const tokens = raw.split(/[|,]/).map(t => t.trim()).filter(Boolean);
  const ids = [];
  for (const tok of tokens) {
    let id = map.get(tok.toLowerCase());
    if (!id && tok.indexOf('.') >= 0) id = tok;
    if (id) ids.push(id);
    else { ids.push('[' + tok + ']'); unresolved.push({ field, value: tok }); }
  }
  return ids.join('|');
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

// Turn one intake row (normalized-label -> value) into a CH field record.
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
          identifiers (live from Content Hub), and downloads a ready-to-import <b>${SHEET_NAME}</b>
          workbook. Add <code>id</code> and <code>identifier</code> columns for updates; leave them out for new records.
          Nothing is written to Content Hub.</div>
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
          const dataRows = aoa.slice(1).filter(r => r.some(c => String(c == null ? '' : c).trim() !== ''));

          // Report recognized vs skipped columns.
          const known = new Set([...Object.keys(FIELD_MAP), ...Object.keys(ID_LABELS), ...ITEM_COLS.map(i => i.label)]);
          const skipped = rawHeaders.filter((h, i) => h && !known.has(normHeaders[i]));
          log(`Data rows: ${dataRows.length}. Recognized columns: ${rawHeaders.filter((h, i) => known.has(normHeaders[i])).length}.`, 'g-info');
          if (skipped.length) log(`Skipped (reference-only): ${skipped.join(', ')}`, 'g-skip');

          // Build records.
          const records = dataRows.map(row => {
            const obj = {};
            for (let i = 0; i < normHeaders.length; i++) if (normHeaders[i]) obj[normHeaders[i]] = row[i];
            return buildRecord(obj);
          });

          // Which CH fields are actually used?
          const used = new Set();
          records.forEach(r => Object.keys(r).forEach(k => used.add(k)));

          // Probe option-list fields (once each) and resolve.
          const meta = {};
          for (const f of OUT_COLS) {
            if (!used.has(f) || TEXT_FIELDS.has(f)) { meta[f] = { isOption: false }; continue; }
            const src = sourceFor(f);
            let values = [];
            try { values = await queryDefinitionValues(src); }
            catch (e) { log(`  (lookup failed for ${src}: ${e.message}) — ${f} treated as text`, 'g-skip'); }
            meta[f] = { isOption: values.length > 0, map: buildLabelMap(values) };
            if (meta[f].isOption) log(`  ${f} → ${src}: ${values.length} options`, 'g-info');
          }

          const unresolved = [];
          for (const r of records) {
            for (const f of Object.keys(r)) {
              if (meta[f] && meta[f].isOption) r[f] = resolveValue(r[f], meta[f].map, f, unresolved);
            }
          }

          log(`Unresolved values: ${unresolved.length}`, unresolved.length ? 'g-err' : 'g-ok');
          unresolved.slice(0, 50).forEach(u => log(`  ✗ ${u.field}: "${u.value}" not found in option list`, 'g-err'));
          if (unresolved.length > 50) log(`  … and ${unresolved.length - 50} more`, 'g-err');

          if (dryRun) { log('Dry run complete — fix any unresolved values, then Generate.', 'g-info'); return; }

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
          if (unresolved.length) log('⚠ Some values were left bracketed [like this] — review before importing.', 'g-err');
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
