// Generic "add + edit + remove" table bound to one Sheets tab. Every list
// page (Products, Suppliers, Customers, Deliveries, Orders, OrderLines,
// Allocations) configures one of these instead of hand-rolling its
// own fetch/append/update/delete logic, so behaviour — including how a row
// gets deleted or overwritten (via its real sheet row index), how the
// form's unsaved-changes guard works, and how sort/search/column-visibility
// behave — stays identical across tables.
//
// Add and Edit share one <dialog>, opened via a toolbar button or a row's
// Edit button rather than sitting inline on the page, with
// js/dialog-guard.js protecting against losing in-progress input on an
// accidental Cancel/Escape/backdrop click.
//
// config = {
//   mount: '#el',            // container to render the panel into
//   title: 'Suppliers',
//   tab: 'Suppliers',        // sheet tab name
//   headers: [...],          // must match Setup's REQUIRED_TABS for this tab
//   autoId: 'SUP' | null,    // prefix for a generated ID in column A, or null
//                            // if column A is a user-entered natural key
//                            // (e.g. Products.CODE) — that key field is
//                            // shown but locked once editing an existing
//                            // row, since other tabs may reference it.
//   fields: [                // one entry per header, in header order
//     { key:'NAME', label:'Name', type:'text', required:true, primary:true }, // primary: this cell (and any 'image' cell) opens the detail dialog
//     { key:'TIER', label:'Tier', type:'text', placeholder:'e.g. Standard' },
//     { key:'ACTIVE', label:'Active', type:'checkbox' },        // stores TRUE/FALSE
//     { key:'ACTIVE', label:'Active', type:'checkbox', trueValue:'Yes', falseValue:'No' }, // stores Yes/No instead
//     { key:'CATEGORY', label:'Category', type:'select', options:[...] },
//     { key:'SUPPLIER', label:'Supplier', type:'ref', refTab:'Suppliers', refValue:'NAME' },
//     { key:'COST', label:'Cost', type:'number', format:'currency' },
//     { key:'PHOTO', label:'Photo', type:'image' },             // plain URL input; renders as a table thumbnail
//   ],
//   sample: [ [...row values in header order, 'auto' for the id col...], ... ]
//   inlineEdit: true,         // opt-in: click a cell to edit it in place instead
//                              // of opening the dialog (see below) — off by
//                              // default so other tables keep today's
//                              // click-Edit-button behaviour unchanged.
// }
//
// fields[i] corresponds to headers[i+1] when autoId is set (headers[0] is
// the generated ID, not user-entered) or headers[i] when autoId is null
// (headers[0] IS a field, e.g. CODE).
//
// `format: 'currency'` on a 'number' field displays it as "£12.74" in the
// table (rounded for display only — the stored value keeps full precision)
// and right-aligns/sorts it numerically like any other number field.
//
// The table itself gets, for free: click-a-header to sort (numeric-aware
// for number/currency fields), a search box that filters across every
// column regardless of visibility, and a Columns picker to hide/show
// individual columns — the hidden-column set persists per table via
// localStorage.
//
// When `inlineEdit` is on:
//   - Any 'image' field is always rendered as the leading column(s), ahead
//     of the sheet's real header order, so a product's photo reads first —
//     this is purely a display reorder; the underlying headers/fields
//     arrays (and therefore the sheet's append-only column order) never
//     change.
//   - Clicking a cell for an editable field type (text/number/select/
//     checkbox/date/ref) turns just that cell into an input; Enter or blur
//     saves it with a single-cell Sheets write, Escape reverts it.
//     'textarea' and 'image' cells aren't inline-editable — edit those from
//     the detail dialog instead.
//   - Clicking an 'image' cell, a field flagged `primary:true`, or (when
//     `autoId` is null) the natural-key column always opens the same
//     add/edit dialog as a "detail card" for the full row — titled with
//     that record's primary value instead of the generic "Edit X" when one
//     is set. A "Details" button in the row's action column does the same,
//     as a guaranteed fallback if no cell in a given row happens to do so.

async function initCrudTable(config) {
  const { mount, title, tab, headers, autoId, fields, sample, inlineEdit } = config;
  const root = document.querySelector(mount);
  const range = `${tab}!A:${colLetter(headers.length)}`;
  const endCol = colLetter(headers.length);
  const singular = /ies$/.test(title) ? title.replace(/ies$/, 'y') : title.replace(/s$/, '');
  let token = null;
  let sheetId = null;
  let rowsData = []; // [{ rowNumber, values }] from the last successful load — array index k always maps to rowNumber k+2, so it's never reordered even when the *displayed* view is sorted/filtered
  let editing = null; // null = Add mode; { rowNumber, id } = editing that existing row
  let sortColIdx = null;
  let sortDir = 1;
  let searchTerm = '';
  const hiddenColsKey = `${tab}_hiddenCols`;
  let hiddenCols = new Set(JSON.parse(localStorage.getItem(hiddenColsKey) || '[]'));
  const refCache = {};
  let inlineEditing = null; // { rowNumber, colIdx, field, original } while a single cell is mid-edit, else null

  root.innerHTML = `
    <div class="panel-h" style="border:none; background:none; padding:0 0 8px;">
      <h3 style="text-transform:none; letter-spacing:normal; font-size:15px; color:var(--ink); font-weight:600;">${escapeHtml(title)}</h3>
      <button type="button" class="btn add-btn sm" id="${tab}_addBtn" style="margin:0;">+ Add ${escapeHtml(singular)}</button>
      <button type="button" class="btn ghost sm" id="${tab}_reload" style="margin:0;">Reload</button>
      ${sample ? `<button type="button" class="btn ghost sm" id="${tab}_seed" style="margin:0;">Load sample data</button>` : ''}
    </div>
    <div class="filter-bar">
      <input type="search" id="${tab}_search" placeholder="Search ${escapeHtml(title.toLowerCase())}...">
      <div class="col-picker-wrap">
        <button type="button" class="btn ghost sm" id="${tab}_colsBtn">Columns</button>
        <div class="col-menu" id="${tab}_colsMenu" style="display:none;">
          ${headers.map(h => `<label><input type="checkbox" data-col="${escapeHtml(h)}" ${hiddenCols.has(h) ? '' : 'checked'}> ${escapeHtml(h)}</label>`).join('')}
        </div>
      </div>
    </div>
    <div class="panel"><div class="scroll" style="overflow-x:auto;"><table id="${tab}_table"><thead><tr></tr></thead><tbody></tbody></table></div></div>
    <div id="${tab}_status" class="mono" style="margin-top:8px; font-size:12.5px; color:var(--steel);"></div>

    <dialog id="${tab}_dialog">
      <form id="${tab}_form"></form>
    </dialog>
  `;

  const statusEl = root.querySelector(`#${tab}_status`);
  const tableEl = root.querySelector(`#${tab}_table`);
  const dialogEl = root.querySelector(`#${tab}_dialog`);
  const form = root.querySelector(`#${tab}_form`);
  const searchInput = root.querySelector(`#${tab}_search`);
  const colsBtn = root.querySelector(`#${tab}_colsBtn`);
  const colsMenu = root.querySelector(`#${tab}_colsMenu`);

  function setStatus(msg, kind) {
    statusEl.textContent = msg;
    statusEl.style.color = kind === 'error' ? 'var(--crit)' : kind === 'ok' ? 'var(--ok)' : 'var(--steel)';
  }

  function setDlgStatus(msg, kind) {
    const el = root.querySelector(`#${tab}_dlgStatus`);
    if (!el) return;
    el.textContent = msg;
    el.style.color = kind === 'error' ? 'var(--crit)' : kind === 'ok' ? 'var(--ok)' : 'var(--steel)';
  }

  async function ensureAuth(statusFn = setStatus) {
    const cfg = getStoredConfig();
    token = tryResumeSession();
    if (!token || !cfg.spreadsheetId) { statusFn('Sign in on the Dashboard first.', 'error'); return null; }
    return cfg;
  }

  async function ensureSheetId(cfg) {
    if (sheetId !== null) return sheetId;
    const tabs = await sheetsGetTabs(cfg.spreadsheetId, token);
    const found = tabs.find(t => t.title === tab);
    sheetId = found ? found.sheetId : null;
    return sheetId;
  }

  async function loadRefOptions(field, cfg) {
    if (refCache[field.refTab]) return refCache[field.refTab];
    const rows = await sheetsGet(cfg.spreadsheetId, `${field.refTab}!A:Z`, token);
    if (!rows.length) { refCache[field.refTab] = []; return []; }
    const [head, ...rest] = rows;
    const valueIdx = head.indexOf(field.refValue);
    const labelIdx = field.refLabel ? head.indexOf(field.refLabel) : -1;
    const opts = rest
      .filter(r => (r[valueIdx] || '').toString().trim() !== '')
      .map(r => ({ value: r[valueIdx], label: labelIdx >= 0 && r[labelIdx] ? `${r[valueIdx]} — ${r[labelIdx]}` : r[valueIdx] }));
    refCache[field.refTab] = opts;
    return opts;
  }

  // Header/field keys can contain spaces or slashes (e.g. "UNIT TYPE",
  // "BATCH/LOT NUMBER") which aren't safe inside a bare CSS id selector —
  // every element id and lookup goes through this sanitized form instead.
  function fid(f) {
    return `${tab}_f_${f.key.replace(/[^A-Za-z0-9]/g, '_')}`;
  }

  // Maps a `headers` column index back to its field definition (undefined
  // for the autoId-owned ID column, which has no field).
  function fieldForHeaderIndex(idx) {
    const fieldIdx = autoId ? idx - 1 : idx;
    return fields[fieldIdx];
  }

  // Column render order: image field(s) first, then everything else in its
  // real header order. Display-only — sorting/hiding/search still key off
  // the real header index carried in each cell's data-col-idx, so this
  // never touches the sheet's actual (append-only) column order.
  const displayOrder = (() => {
    const idxs = headers.map((_, i) => i);
    const imageIdxs = idxs.filter(i => { const f = fieldForHeaderIndex(i); return f && f.type === 'image'; });
    const restIdxs = idxs.filter(i => !imageIdxs.includes(i));
    return [...imageIdxs, ...restIdxs];
  })();

  function isInlineEditable(field) {
    return !!field && ['text', 'number', 'select', 'checkbox', 'date', 'ref'].includes(field.type);
  }

  // Whether clicking this column's cell should open the detail dialog
  // instead of (or as well as, for non-editable types) inline-editing.
  function opensDetail(field, idx) {
    if (!inlineEdit || !field) return false;
    if (field.type === 'image' || field.primary) return true;
    if (!autoId && idx === 0) return true; // the natural key — locked in the dialog too
    return false;
  }

  function fieldInputHtml(f) {
    const id = fid(f);
    if (f.type === 'textarea') return `<textarea id="${id}" placeholder="${escapeHtml(f.placeholder || '')}"></textarea>`;
    if (f.type === 'checkbox') return `<input type="checkbox" id="${id}">`;
    if (f.type === 'number') return `<input type="number" step="any" id="${id}" placeholder="${escapeHtml(f.placeholder || '')}">`;
    if (f.type === 'date') return `<input type="date" id="${id}">`;
    if (f.type === 'select') return `<select id="${id}"><option value="">—</option>${f.options.map(o => {
      const opt = typeof o === 'object' ? o : { value: o, label: o };
      return `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`;
    }).join('')}</select>`;
    if (f.type === 'ref') return `<select id="${id}"><option value="">Loading…</option></select>`;
    // 'image' falls through to the default plain text input below — it's
    // just a URL field; only its table display differs (see displayCell).
    return `<input id="${id}" placeholder="${escapeHtml(f.placeholder || '')}">`;
  }

  // Builds the dialog's form DOM once. Ref-type <select> options are filled
  // in later, on each dialog open, so they're never stale.
  function renderForm() {
    form.innerHTML = `
      <div class="dlg-h" id="${tab}_dlgTitle">Add ${escapeHtml(singular)}</div>
      <div class="dlg-b">
        <div class="row">
          ${fields.map(f => `<div class="col field"><label>${escapeHtml(f.label)}${f.required ? ' *' : ''}</label>${fieldInputHtml(f)}</div>`).join('')}
        </div>
        <div id="${tab}_dlgStatus" class="mono" style="margin-top:10px; font-size:12.5px; color:var(--steel);"></div>
      </div>
      <div class="dlg-f">
        <button type="button" class="btn ghost" data-cancel>Cancel</button>
        <button type="submit" class="btn" id="${tab}_submitBtn">Add ${escapeHtml(singular)}</button>
      </div>
    `;
    resetFieldValues();
  }

  function resetFieldValues() {
    for (const f of fields) {
      const el = form.querySelector(`#${fid(f)}`);
      if (f.type === 'checkbox') el.checked = false;
      else if (f.type === 'date' && f.default === 'today') el.value = new Date().toISOString().slice(0, 10);
      else el.value = '';
    }
    // The natural-key field (Products.CODE etc.) is only ever locked while
    // editing an existing row — Add mode always starts unlocked.
    if (!autoId) form.querySelector(`#${fid(fields[0])}`).disabled = false;
  }

  // Fills the form from an existing row's raw cell values (index-aligned to
  // `headers`, offset by 1 when autoId owns column A). Dispatches a real
  // 'change' event per field afterwards so page-specific logic hooked onto
  // these inputs (e.g. the Allocations FEFO helper) reacts the same way it
  // would to the user actually picking that value.
  function populateFieldsFromRow(rowValues) {
    const offset = autoId ? 1 : 0;
    fields.forEach((f, i) => {
      const el = form.querySelector(`#${fid(f)}`);
      const raw = rowValues[i + offset] ?? '';
      if (f.type === 'checkbox') el.checked = (raw || '').toString().toUpperCase() === (f.trueValue || 'TRUE').toUpperCase();
      else el.value = raw;
    });
    fields.forEach(f => {
      form.querySelector(`#${fid(f)}`).dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  async function refreshRefOptions() {
    const cfg = getStoredConfig();
    if (!cfg.spreadsheetId) return;
    for (const f of fields) {
      if (f.type !== 'ref') continue;
      const sel = form.querySelector(`#${fid(f)}`);
      try {
        delete refCache[f.refTab]; // always fetch fresh when the dialog opens
        const opts = await loadRefOptions(f, cfg);
        sel.innerHTML = `<option value="">—</option>` + opts.map(o => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`).join('');
      } catch (err) {
        sel.innerHTML = `<option value="">(error loading ${escapeHtml(f.refTab)})</option>`;
      }
    }
  }

  function fieldValue(f) {
    const el = form.querySelector(`#${fid(f)}`);
    if (f.type === 'checkbox') return el.checked ? (f.trueValue || 'TRUE') : (f.falseValue || 'FALSE');
    return el.value.trim();
  }

  function isNumericField(f) {
    return !!f && (f.type === 'number' || f.format === 'currency');
  }

  function displayCell(field, value) {
    if (field && field.type === 'checkbox') {
      const on = (value || '').toString().toUpperCase() === (field.trueValue || 'TRUE').toUpperCase();
      const onLabel = field.trueValue || 'Yes';
      const offLabel = field.falseValue || 'No';
      return `<span class="tag ${on ? 't-ok' : 't-amb'}">${on ? onLabel : offLabel}</span>`;
    }
    if (field && field.format === 'currency') {
      const n = parseFloat(value);
      return isNaN(n) ? '' : `£${n.toFixed(2)}`;
    }
    if (field && field.type === 'image') {
      const url = (value || '').toString().trim();
      const box = 'width:38px;height:38px;border-radius:8px;vertical-align:middle;background:#EEF0F2;';
      return url
        ? `<img src="${escapeHtml(url)}" alt="" loading="lazy" style="${box}object-fit:cover;display:block;">`
        : `<span style="${box}display:inline-block;"></span>`;
    }
    return escapeHtml(value);
  }

  // ---- inline cell editing (opt-in via config.inlineEdit) ----

  function inlineInputHtml(field, value) {
    const v = (value ?? '').toString();
    if (field.type === 'checkbox') {
      const checked = v.toUpperCase() === (field.trueValue || 'TRUE').toUpperCase();
      return `<input type="checkbox" ${checked ? 'checked' : ''}>`;
    }
    if (field.type === 'number') return `<input type="number" step="any" value="${escapeHtml(v)}">`;
    if (field.type === 'date') return `<input type="date" value="${escapeHtml(v)}">`;
    if (field.type === 'select') return `<select><option value="">—</option>${field.options.map(o => {
      const opt = typeof o === 'object' ? o : { value: o, label: o };
      return `<option value="${escapeHtml(opt.value)}" ${opt.value === v ? 'selected' : ''}>${escapeHtml(opt.label)}</option>`;
    }).join('')}</select>`;
    if (field.type === 'ref') return `<select><option value="${escapeHtml(v)}">${escapeHtml(v || 'Loading…')}</option></select>`;
    return `<input value="${escapeHtml(v)}">`;
  }

  async function populateInlineRefOptions(field, selectEl) {
    const cfg = getStoredConfig();
    if (!cfg.spreadsheetId) return;
    try {
      delete refCache[field.refTab]; // always fetch fresh, same as the dialog does
      const opts = await loadRefOptions(field, cfg);
      const current = selectEl.value;
      selectEl.innerHTML = `<option value="">—</option>` + opts.map(o => `<option value="${escapeHtml(o.value)}" ${o.value === current ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('');
    } catch (err) {
      // leave the current-value placeholder option in place
    }
  }

  // Turns one <td> into a live input, saving on blur/Enter (or immediately
  // on change, for select/ref/checkbox) with a single-cell Sheets write —
  // no full-row resubmission and no dialog. Escape reverts without saving.
  function beginInlineEdit(td, rowNumber, colIdx, field) {
    if (inlineEditing) return; // one cell at a time; the other one commits on its own blur
    const row = rowsData[rowNumber - 2];
    if (!row) return;
    const original = row.values[colIdx] ?? '';
    inlineEditing = { rowNumber, colIdx, field, original };
    td.innerHTML = inlineInputHtml(field, original);
    const input = td.querySelector('input, select');
    input.focus();
    if (input.select) input.select();
    if (field.type === 'ref') populateInlineRefOptions(field, input);

    const revert = () => { td.innerHTML = displayCell(field, original); };

    const commit = async () => {
      if (!inlineEditing) return; // already committed/cancelled via another event
      inlineEditing = null;
      const value = field.type === 'checkbox'
        ? (input.checked ? (field.trueValue || 'TRUE') : (field.falseValue || 'FALSE'))
        : input.value.trim();

      if (field.required && !value) { setStatus(`${field.label} is required.`, 'error'); revert(); return; }
      if (field.type !== 'checkbox' && value === (original ?? '').toString().trim()) { revert(); return; }

      const cfg = await ensureAuth();
      if (!cfg) { revert(); return; }
      const col = colLetter(colIdx + 1);
      try {
        await sheetsUpdateRange(cfg.spreadsheetId, `${tab}!${col}${rowNumber}:${col}${rowNumber}`, [[value]], token, 'USER_ENTERED');
        row.values[colIdx] = value;
        td.innerHTML = displayCell(field, value);
        setStatus('Saved.', 'ok');
      } catch (err) {
        setStatus('Error saving: ' + err.message, 'error');
        revert();
      }
    };
    const cancel = () => { inlineEditing = null; revert(); };

    input.addEventListener('blur', commit);
    if (field.type === 'select' || field.type === 'ref' || field.type === 'checkbox') {
      input.addEventListener('change', commit);
    }
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
    });
  }

  // Best available title for a record's detail dialog: its `primary`
  // field's value, falling back to the natural key (when there is one),
  // falling back to the generic "Edit X" set by the caller.
  function recordTitle(rowValues) {
    const offset = autoId ? 1 : 0;
    const primaryField = fields.find(f => f.primary);
    if (primaryField) {
      const v = (rowValues[fields.indexOf(primaryField) + offset] || '').toString().trim();
      if (v) return v;
    }
    if (!autoId) {
      const v = (rowValues[0] || '').toString().trim();
      if (v) return v;
    }
    return null;
  }

  async function loadRows() {
    const cfg = getStoredConfig();
    if (!cfg.spreadsheetId) { setStatus('Sign in on the Dashboard first.', 'error'); return; }
    token = tryResumeSession();
    if (!token) { setStatus('Sign in on the Dashboard first.', 'error'); return; }
    setStatus('Loading...');
    try {
      const rows = await sheetsGet(cfg.spreadsheetId, range, token);
      const rest = rows.length ? rows.slice(1) : [];
      rowsData = rest.map((values, i) => ({ rowNumber: i + 2, values }));
      applyViewAndRender();
      setStatus(`Loaded ${rest.length} row(s).`, 'ok');
    } catch (err) {
      setStatus('Error loading: ' + err.message, 'error');
    }
  }

  function compareCell(a, b, colIdx) {
    if (isNumericField(fieldForHeaderIndex(colIdx))) {
      const na = parseFloat(a); const nb = parseFloat(b);
      const va = isNaN(na) ? -Infinity : na;
      const vb = isNaN(nb) ? -Infinity : nb;
      return va - vb;
    }
    return (a ?? '').toString().localeCompare((b ?? '').toString());
  }

  function applyViewAndRender() {
    let view = rowsData;
    if (searchTerm) {
      view = view.filter(d => d.values.some(v => (v ?? '').toString().toLowerCase().includes(searchTerm)));
    }
    if (sortColIdx !== null) {
      view = [...view].sort((a, b) => compareCell(a.values[sortColIdx], b.values[sortColIdx], sortColIdx) * sortDir);
    }
    renderTable(view);
  }

  function renderTable(viewRows) {
    const thead = tableEl.querySelector('thead tr');
    const tbody = tableEl.querySelector('tbody');

    thead.innerHTML = displayOrder.map(idx => {
      const h = headers[idx];
      if (hiddenCols.has(h)) return '';
      const numCls = isNumericField(fieldForHeaderIndex(idx)) ? ' num' : '';
      const arrow = sortColIdx === idx ? (sortDir === 1 ? ' ▲' : ' ▼') : '';
      return `<th class="sortable${numCls}" data-col-idx="${idx}">${escapeHtml(h)}${arrow}</th>`;
    }).join('') + '<th></th>';

    if (!viewRows.length) {
      const visibleCount = headers.filter(h => !hiddenCols.has(h)).length + 1;
      const msg = rowsData.length ? 'No rows match your search.' : 'No rows yet.';
      tbody.innerHTML = `<tr><td colspan="${visibleCount}" class="empty" style="padding:18px;">${msg}</td></tr>`;
      return;
    }

    tbody.innerHTML = viewRows.map(d => {
      const cells = displayOrder.map(idx => {
        const h = headers[idx];
        if (hiddenCols.has(h)) return '';
        const field = fieldForHeaderIndex(idx);
        const classes = [];
        if (isNumericField(field)) classes.push('num');
        let titleAttr = '';
        if (opensDetail(field, idx)) { classes.push('cell-detail'); titleAttr = ' title="Click to view details"'; }
        else if (inlineEdit && isInlineEditable(field)) { classes.push('cell-editable'); titleAttr = ' title="Click to edit"'; }
        const classAttr = classes.length ? ` class="${classes.join(' ')}"` : '';
        return `<td${classAttr}${titleAttr} data-col-idx="${idx}">${displayCell(field, d.values[idx])}</td>`;
      }).join('');
      const editLabel = inlineEdit ? 'Details' : 'Edit';
      return `<tr data-row="${d.rowNumber}">${cells}<td style="white-space:nowrap;"><button type="button" class="btn ghost sm" data-edit="${d.rowNumber}">${editLabel}</button> <button type="button" class="btn danger sm" data-del="${d.rowNumber}">Remove</button></td></tr>`;
    }).join('');
  }

  async function handleDelete(rowNumber) {
    if (!confirm('Remove this row? This can\'t be undone.')) return;
    const cfg = await ensureAuth();
    if (!cfg) return;
    try {
      const sid = await ensureSheetId(cfg);
      if (sid === null) throw new Error(`Tab "${tab}" not found — run Setup first.`);
      await sheetsBatchUpdate(cfg.spreadsheetId, [{
        deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: rowNumber - 1, endIndex: rowNumber } }
      }], token);
      setStatus('Row removed.', 'ok');
      loadRows();
    } catch (err) {
      setStatus('Error removing row: ' + err.message, 'error');
    }
  }

  tableEl.addEventListener('click', (e) => {
    const th = e.target.closest('th[data-col-idx]');
    if (th) {
      const idx = parseInt(th.dataset.colIdx, 10);
      if (sortColIdx === idx) sortDir *= -1; else { sortColIdx = idx; sortDir = 1; }
      applyViewAndRender();
      return;
    }
    const delBtn = e.target.closest('button[data-del]');
    if (delBtn) { handleDelete(parseInt(delBtn.dataset.del, 10)); return; }
    const editBtn = e.target.closest('button[data-edit]');
    if (editBtn) { handleEdit(parseInt(editBtn.dataset.edit, 10)); return; }

    if (!inlineEdit) return;
    const td = e.target.closest('td[data-col-idx]');
    if (!td || inlineEditing) return; // ignore stray clicks while a cell is already mid-edit
    const idx = parseInt(td.dataset.colIdx, 10);
    const rowNumber = parseInt(td.closest('tr').dataset.row, 10);
    const field = fieldForHeaderIndex(idx);
    if (opensDetail(field, idx)) { handleEdit(rowNumber); return; }
    if (isInlineEditable(field)) beginInlineEdit(td, rowNumber, idx, field);
  });

  searchInput.addEventListener('input', () => {
    searchTerm = searchInput.value.trim().toLowerCase();
    applyViewAndRender();
  });

  colsBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    colsMenu.style.display = colsMenu.style.display === 'none' ? 'block' : 'none';
  });
  document.addEventListener('click', (e) => {
    if (colsMenu.style.display !== 'none' && !colsMenu.contains(e.target) && e.target !== colsBtn) {
      colsMenu.style.display = 'none';
    }
  });
  colsMenu.addEventListener('change', (e) => {
    const cb = e.target.closest('input[type=checkbox]');
    if (!cb) return;
    const col = cb.dataset.col;
    if (cb.checked) hiddenCols.delete(col); else hiddenCols.add(col);
    localStorage.setItem(hiddenColsKey, JSON.stringify([...hiddenCols]));
    applyViewAndRender();
  });

  renderForm();
  const guard = attachDiscardGuard(dialogEl, { scope: form });

  async function openAddDialog() {
    editing = null;
    root.querySelector(`#${tab}_dlgTitle`).textContent = `Add ${singular}`;
    root.querySelector(`#${tab}_submitBtn`).textContent = `Add ${singular}`;
    setDlgStatus('');
    resetFieldValues();
    await refreshRefOptions();
    dialogEl.showModal();
    guard.arm();
  }

  async function handleEdit(rowNumber) {
    const row = rowsData[rowNumber - 2];
    if (!row) return;
    const rowValues = row.values;
    editing = { rowNumber, id: rowValues[0] };
    root.querySelector(`#${tab}_dlgTitle`).textContent = recordTitle(rowValues) || `Edit ${singular}`;
    root.querySelector(`#${tab}_submitBtn`).textContent = 'Save changes';
    setDlgStatus('');
    resetFieldValues();
    await refreshRefOptions();
    populateFieldsFromRow(rowValues);
    // The natural key (e.g. Products.CODE) can't be changed once other tabs
    // may already reference it by that value.
    if (!autoId) form.querySelector(`#${fid(fields[0])}`).disabled = true;
    dialogEl.showModal();
    guard.arm();
  }

  root.querySelector(`#${tab}_addBtn`).addEventListener('click', openAddDialog);
  form.querySelector('[data-cancel]').addEventListener('click', () => guard.guardedClose());

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const cfg = await ensureAuth(setDlgStatus);
    if (!cfg) return;

    for (const f of fields) {
      if (f.required && !fieldValue(f)) { setDlgStatus(`${f.label} is required.`, 'error'); return; }
    }

    const isEdit = !!editing;
    setDlgStatus(isEdit ? 'Saving...' : 'Adding...');
    try {
      const values = fields.map(fieldValue);
      if (isEdit) {
        const row = autoId ? [editing.id, ...values] : values;
        await sheetsUpdateRange(cfg.spreadsheetId, `${tab}!A${editing.rowNumber}:${endCol}${editing.rowNumber}`, [row], token, 'USER_ENTERED');
        setStatus(`${singular} updated.`, 'ok');
      } else {
        let row = values;
        if (autoId) {
          const existing = await sheetsGet(cfg.spreadsheetId, range, token);
          const id = nextId(existing.length ? existing : [headers], autoId);
          row = [id, ...values];
        }
        await sheetsAppend(cfg.spreadsheetId, range, row, token);
        setStatus(`${singular} added.`, 'ok');
      }
      dialogEl.close(); // programmatic close — the discard guard only ever intercepts 'cancel'/backdrop, never this
      loadRows();
    } catch (err) {
      setDlgStatus('Error: ' + err.message, 'error');
    }
  });

  if (sample) {
    root.querySelector(`#${tab}_seed`).addEventListener('click', async () => {
      const cfg = await ensureAuth();
      if (!cfg) return;
      setStatus('Loading sample data...');
      try {
        // Column A is often a generated ID (not a stable key sample rows can
        // be matched against), so instead of trying to dedupe row-by-row,
        // sample data only ever loads into a tab that has no data rows yet —
        // simple and impossible to double up.
        const existing = await sheetsGet(cfg.spreadsheetId, range, token);
        if (existing.slice(1).length) {
          setStatus('This table already has data — sample rows weren\'t added, to avoid duplicates.', 'error');
          return;
        }
        let rowsToAdd = sample;
        if (autoId) {
          let base = [headers];
          rowsToAdd = sample.map(r => {
            const id = nextId(base, autoId);
            base = base.concat([[id]]);
            return [id, ...r.slice(1)];
          });
        }
        await sheetsAppendRows(cfg.spreadsheetId, range, rowsToAdd, token);
        setStatus(`Added ${rowsToAdd.length} sample row(s).`, 'ok');
        loadRows();
      } catch (err) {
        setStatus('Error loading sample data: ' + err.message, 'error');
      }
    });
  }

  root.querySelector(`#${tab}_reload`).addEventListener('click', loadRows);

  loadRows();
}
