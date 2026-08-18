// Generic "add + edit + remove" table bound to one Sheets tab. Every list
// page (Products, Suppliers, Customers, Deliveries, Orders, OrderLines,
// Allocations, Prices) configures one of these instead of hand-rolling its
// own fetch/append/update/delete logic, so behaviour — including how a row
// gets deleted or overwritten (via its real sheet row index), and how the
// form's unsaved-changes guard works — stays identical across tables.
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
//   ],
//   sample: [ [...row values in header order, 'auto' for the id col...], ... ]
// }
//
// fields[i] corresponds to headers[i+1] when autoId is set (headers[0] is
// the generated ID, not user-entered) or headers[i] when autoId is null
// (headers[0] IS a field, e.g. CODE).

async function initCrudTable(config) {
  const { mount, title, tab, headers, autoId, fields, sample } = config;
  const root = document.querySelector(mount);
  const range = `${tab}!A:${colLetter(headers.length)}`;
  const endCol = colLetter(headers.length);
  const singular = /ies$/.test(title) ? title.replace(/ies$/, 'y') : title.replace(/s$/, '');
  let token = null;
  let sheetId = null;
  let rowsData = []; // data rows (no header) from the last successful load, for Edit lookups
  let editing = null; // null = Add mode; { rowNumber, id } = editing that existing row
  const refCache = {};

  root.innerHTML = `
    <div class="panel-h" style="border:none; background:none; padding:0 0 8px;">
      <h3 style="text-transform:none; letter-spacing:normal; font-size:15px; color:var(--ink); font-weight:600;">${escapeHtml(title)}</h3>
      <button type="button" class="btn add-btn sm" id="${tab}_addBtn" style="margin:0;">+ Add ${escapeHtml(singular)}</button>
      <button type="button" class="btn ghost sm" id="${tab}_reload" style="margin:0;">Reload</button>
      ${sample ? `<button type="button" class="btn ghost sm" id="${tab}_seed" style="margin:0;">Load sample data</button>` : ''}
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

  function fieldInputHtml(f) {
    const id = fid(f);
    if (f.type === 'textarea') return `<textarea id="${id}" placeholder="${escapeHtml(f.placeholder || '')}"></textarea>`;
    if (f.type === 'checkbox') return `<input type="checkbox" id="${id}">`;
    if (f.type === 'number') return `<input type="number" step="any" id="${id}" placeholder="${escapeHtml(f.placeholder || '')}">`;
    if (f.type === 'date') return `<input type="date" id="${id}">`;
    if (f.type === 'select') return `<select id="${id}"><option value="">—</option>${f.options.map(o => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join('')}</select>`;
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

  function displayCell(header, value) {
    const field = fields.find(f => f.key === header);
    if (field && field.type === 'checkbox') {
      const on = (value || '').toString().toUpperCase() === 'TRUE';
      return `<span class="tag ${on ? 't-ok' : 't-amb'}">${on ? 'Yes' : 'No'}</span>`;
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
      const [head, ...rest] = rows.length ? rows : [headers];
      rowsData = rest;
      renderTable(head, rest);
      setStatus(`Loaded ${rest.length} row(s).`, 'ok');
    } catch (err) {
      setStatus('Error loading: ' + err.message, 'error');
    }
  }

  function renderTable(head, rows) {
    const thead = tableEl.querySelector('thead tr');
    const tbody = tableEl.querySelector('tbody');
    thead.innerHTML = head.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '<th></th>';
    if (!rows.length) {
      tbody.innerHTML = `<tr><td colspan="${head.length + 1}" class="empty" style="padding:18px;">No rows yet.</td></tr>`;
      return;
    }
    tbody.innerHTML = rows.map((r, i) => {
      const cells = head.map(h => `<td>${displayCell(h, r[head.indexOf(h)])}</td>`).join('');
      return `<tr data-row="${i + 2}">${cells}<td style="white-space:nowrap;"><button type="button" class="btn ghost sm" data-edit="${i + 2}">Edit</button> <button type="button" class="btn danger sm" data-del="${i + 2}">Remove</button></td></tr>`;
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
    const delBtn = e.target.closest('button[data-del]');
    if (delBtn) { handleDelete(parseInt(delBtn.dataset.del, 10)); return; }
    const editBtn = e.target.closest('button[data-edit]');
    if (editBtn) handleEdit(parseInt(editBtn.dataset.edit, 10));
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
    const rowValues = rowsData[rowNumber - 2];
    if (!rowValues) return;
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
