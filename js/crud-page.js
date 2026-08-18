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
//     { key:'NAME', label:'Name', type:'text', required:true },
//     { key:'TIER', label:'Tier', type:'text', placeholder:'e.g. Standard' },
//     { key:'ACTIVE', label:'Active', type:'checkbox' },
//     { key:'CATEGORY', label:'Category', type:'select', options:[...] },
//     { key:'SUPPLIER', label:'Supplier', type:'ref', refTab:'Suppliers', refValue:'NAME' },
//     { key:'COST', label:'Cost', type:'number', format:'currency' },
//   ],
//   sample: [ [...row values in header order, 'auto' for the id col...], ... ]
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

async function initCrudTable(config) {
  const { mount, title, tab, headers, autoId, fields, sample } = config;
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
      if (f.type === 'checkbox') el.checked = (raw || '').toString().toUpperCase() === 'TRUE';
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
    if (f.type === 'checkbox') return el.checked ? 'TRUE' : 'FALSE';
    return el.value.trim();
  }

  function isNumericField(f) {
    return !!f && (f.type === 'number' || f.format === 'currency');
  }

  function displayCell(field, value) {
    if (field && field.type === 'checkbox') {
      const on = (value || '').toString().toUpperCase() === 'TRUE';
      return `<span class="tag ${on ? 't-ok' : 't-amb'}">${on ? 'Yes' : 'No'}</span>`;
    }
    if (field && field.format === 'currency') {
      const n = parseFloat(value);
      return isNaN(n) ? '' : `£${n.toFixed(2)}`;
    }
    return escapeHtml(value);
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

    thead.innerHTML = headers.map((h, idx) => {
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
      const cells = headers.map((h, idx) => {
        if (hiddenCols.has(h)) return '';
        const field = fieldForHeaderIndex(idx);
        const numCls = isNumericField(field) ? ' class="num"' : '';
        return `<td${numCls}>${displayCell(field, d.values[idx])}</td>`;
      }).join('');
      return `<tr data-row="${d.rowNumber}">${cells}<td style="white-space:nowrap;"><button type="button" class="btn ghost sm" data-edit="${d.rowNumber}">Edit</button> <button type="button" class="btn danger sm" data-del="${d.rowNumber}">Remove</button></td></tr>`;
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
    if (editBtn) handleEdit(parseInt(editBtn.dataset.edit, 10));
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
    root.querySelector(`#${tab}_dlgTitle`).textContent = `Edit ${singular}`;
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
