// ============================================================================
// Toll Product Import Generator — Content Hub External Component (Tile v1)
// ----------------------------------------------------------------------------
// Browser-based replacement for the Excel macro "Export to Content Hub File".
//
// Input : an .xlsx/.csv whose HEADER ROW is Content Hub field names
//         (id, identifier, TB.PCM.ProductName, TB.PCM.Brand, TB.PCM.TileColorFamily,
//         TB.PCM.AreaOfApplication, TB.PCM.DivisionSelected, ...). Cells hold
//         DISPLAY values (e.g. "Beige", "Floor Tile", "Arizona, Colorado").
// Output: a downloadable workbook with ONE sheet named "M.PCM.Product" where
//         option-list / taxonomy display values are resolved to their Content
//         Hub identifiers (multi-value fields joined with "|"). Ready for the
//         standard Content Hub Excel import.
//
// Lookups are pulled LIVE from Content Hub: for a field, the backing taxonomy
// is queried (Definition.Name=='<source>') and a label -> identifier map built.
// Field -> source overrides mirror the macro's GetLookupSourceForCHField.
//
// This tool only READS taxonomies and WRITES a local file; it never creates or
// modifies Content Hub entities.
// ============================================================================

const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const CH_HOST = window.location.origin;
const SHEET_NAME = 'M.PCM.Product';

// Field -> backing taxonomy/datasource (mirrors GetLookupSourceForCHField).
const FIELD_SOURCE = {
  'TB.PCM.Product.Manufacturer': 'TB.PCM.Manufacturer',
  'TB.PCM.Product.TileUnitOfMeasure': 'TB.PCM.UnitOfMeasure',
  'TB.PCM.DivisionSelected': 'TB.Division'
};

// Columns that are never option-lists (skip the lookup probe).
const ALWAYS_TEXT = new Set(['id', 'identifier', 'ProductNumber', 'Color']);

function sourceFor(field) { return FIELD_SOURCE[field] || field; }

const CSS = `
  .g-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 820px; }
  .g-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .g-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .g-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 32px; text-align: center;
             cursor: pointer; color: #555; margin-bottom: 12px; }
  .g-drop.g-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .g-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; }
  .g-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .g-btn:disabled { opacity: .5; cursor: not-allowed; }
  .g-dry   { background: #edf2f7; color: #2d3748; }
  .g-go    { background: #2f855a; color: #fff; }
  .g-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px;
             padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 340px; overflow: auto;
             white-space: pre-wrap; display: none; }
  .g-ok   { color: #68d391; }
  .g-skip { color: #cbd5e0; }
  .g-err  { color: #fc8181; }
  .g-info { color: #90cdf4; }
`;

let XLSX;
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

// Read a property off a REST entity resource (plain or culture-keyed).
function readProp(item, name) {
  const p = item && item.properties ? item.properties[name] : undefined;
  if (p == null) return '';
  if (typeof p === 'object') {
    const keys = Object.keys(p);
    return keys.length ? String(p[keys[0]]) : '';
  }
  return String(p);
}

function readLabel(item) {
  return readProp(item, 'Label') || readProp(item, 'TaxonomyLabel') || readProp(item, 'Name') || '';
}

// Query every value of a taxonomy/definition; return [{identifier, label}].
async function queryDefinitionValues(defName) {
  const chql = `Definition.Name=='${defName}'`;
  const out = [];
  let skip = 0, safety = 0;
  while (safety++ < 500) {
    const url = `${CH_HOST}/api/entities/query?query=${encodeURIComponent(chql)}&skip=${skip}&take=200`;
    const res = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
    if (!res.ok) {
      const b = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}${b ? ' — ' + b : ''}`);
    }
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

// label(lower) / last-segment(lower) / full-id(lower) -> identifier
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

// Resolve one cell's display value(s) to pipe-joined identifiers.
function resolveCell(value, map, field, unresolved) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const tokens = raw.split(/[|,]/).map(t => t.trim()).filter(Boolean);
  const ids = [];
  for (const tok of tokens) {
    let id = map.get(tok.toLowerCase());
    if (!id && tok.indexOf('.') >= 0) id = tok; // already an identifier from a newer taxonomy
    if (id) { ids.push(id); }
    else { ids.push('[' + tok + ']'); unresolved.push({ field, value: tok }); }
  }
  return ids.join('|');
}

async function parseFileAOA(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
}

function ts() {
  const d = new Date(), p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render() {
      const style = document.createElement('style');
      style.textContent = CSS;

      const wrap = document.createElement('div');
      wrap.className = 'g-wrap';
      wrap.innerHTML = `
        <div class="g-title">🏭 Product Import Generator — Tile</div>
        <div class="g-sub">Upload a spreadsheet with Content Hub field-name headers and display values.
          The tool resolves option-list values to identifiers (live from Content Hub) and downloads a
          ready-to-import <b>${SHEET_NAME}</b> workbook. Nothing is written to Content Hub.</div>
        <div class="g-drop" id="g-drop">📎 Drop your .xlsx / .csv here, or click to browse</div>
        <input type="file" id="g-file" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-row">
          <button class="g-btn g-dry" id="g-dry" disabled>🔍 Validate (dry run)</button>
          <button class="g-btn g-go"  id="g-go"  disabled>⬇ Generate import file</button>
          <span id="g-status" style="font-size:13px;color:#555"></span>
        </div>
        <div class="g-log" id="g-log"></div>
        <div style="font-size:12px;color:#888;margin-top:10px">
          Headers = Content Hub field names (e.g. id, identifier, TB.PCM.Brand, TB.PCM.TileColorFamily,
          TB.PCM.AreaOfApplication, TB.PCM.DivisionSelected). Multi-value cells may be comma-separated.
        </div>
      `;

      rootElement.innerHTML = '';
      rootElement.appendChild(style);
      rootElement.appendChild(wrap);

      const drop   = wrap.querySelector('#g-drop');
      const input  = wrap.querySelector('#g-file');
      const dryBtn = wrap.querySelector('#g-dry');
      const goBtn  = wrap.querySelector('#g-go');
      const status = wrap.querySelector('#g-status');
      const logEl  = wrap.querySelector('#g-log');

      let currentFile = null;

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div');
        if (cls) line.className = cls;
        line.textContent = msg;
        logEl.appendChild(line);
        logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }

      function onFile(file) {
        if (!file) return;
        clearLog();
        currentFile = file;
        status.textContent = `${file.name} — ready`;
        dryBtn.disabled = false;
        goBtn.disabled = false;
      }

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('g-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('g-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('g-hover'); onFile(e.dataTransfer.files[0]); });

      async function run(dryRun) {
        clearLog();
        dryBtn.disabled = true; goBtn.disabled = true;
        log(dryRun ? '── VALIDATE (no file written) ──' : '── GENERATE IMPORT FILE ──', 'g-info');
        try {
          await loadXLSX();
          const aoa = await parseFileAOA(currentFile);
          if (!aoa || aoa.length < 2) { log('No data rows found.', 'g-err'); return; }

          const headers = aoa[0].map(h => String(h == null ? '' : h).trim());
          const dataRows = aoa.slice(1).filter(r => r.some(c => String(c == null ? '' : c).trim() !== ''));
          log(`Headers: ${headers.length}. Data rows: ${dataRows.length}.`, 'g-info');

          // Probe each column (once per source) to see if it's an option-list.
          const sourceCache = new Map();
          const colInfo = [];
          for (const header of headers) {
            if (!header || ALWAYS_TEXT.has(header)) { colInfo.push({ header, isOption: false }); continue; }
            const source = sourceFor(header);
            let entry = sourceCache.get(source);
            if (!entry) {
              let values = [];
              try { values = await queryDefinitionValues(source); }
              catch (e) { log(`  (lookup failed for ${source}: ${e.message}) — ${header} treated as text`, 'g-skip'); }
              entry = { map: buildLabelMap(values), isOption: values.length > 0, count: values.length };
              sourceCache.set(source, entry);
              if (entry.isOption) log(`  ${header} → ${source}: ${entry.count} options`, 'g-info');
            }
            colInfo.push({ header, isOption: entry.isOption, map: entry.map });
          }

          // Transform rows.
          const unresolved = [];
          const outRows = [headers.slice()];
          for (const row of dataRows) {
            const out = [];
            for (let c = 0; c < headers.length; c++) {
              const ci = colInfo[c];
              const val = row[c];
              if (ci && ci.isOption) out.push(resolveCell(val, ci.map, ci.header, unresolved));
              else out.push(val == null ? '' : val);
            }
            outRows.push(out);
          }

          log(`Unresolved values: ${unresolved.length}`, unresolved.length ? 'g-err' : 'g-ok');
          unresolved.slice(0, 50).forEach(u => log(`  ✗ ${u.field}: "${u.value}" not found in option list`, 'g-err'));
          if (unresolved.length > 50) log(`  … and ${unresolved.length - 50} more`, 'g-err');

          if (dryRun) { log('Dry run complete — fix any unresolved values, then Generate.', 'g-info'); return; }

          const ws = XLSX.utils.aoa_to_sheet(outRows);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
          const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const fname = `ContentHub_TilesImport_${ts()}.xlsx`;
          downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), fname);
          log(`✓ Generated ${fname} — ${dataRows.length} row(s), sheet "${SHEET_NAME}".`, 'g-ok');
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
