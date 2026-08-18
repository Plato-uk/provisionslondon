// Self-contained date-range popover picker (no external calendar library —
// the project has no build step, so this is plain vanilla JS/CSS). Used by
// purchase-orders.html for "Expected delivery" since a plain <input
// type=date> can't express a range.
//
// createDateRangePicker({ mount, onChange, initialStart, initialEnd })
//   mount:      '#el' — an empty container this renders into
//   onChange:   ({ start, end }) => {}  — start/end are 'YYYY-MM-DD' or null,
//               fired on every day click and on Clear
// Returns { getRange(), setRange(start, end), clear(), close() }.
//
// Click a day to start a range; click a second day to complete it (earlier
// date always becomes start, regardless of click order). Clicking again
// after a range is complete starts a new one.

function createDateRangePicker({ mount, onChange, initialStart, initialEnd }) {
  const root = document.querySelector(mount);
  let start = initialStart || null;
  let end = initialEnd || null;
  const today = new Date();
  const seed = start ? new Date(start + 'T00:00:00') : today;
  let viewYear = seed.getFullYear();
  let viewMonth = seed.getMonth();

  root.innerHTML = `
    <button type="button" class="dr-trigger mono"></button>
    <div class="dr-pop" hidden>
      <div class="dr-hd">
        <button type="button" class="dr-nav" data-dir="-1" aria-label="Previous month">‹</button>
        <span class="dr-label"></span>
        <button type="button" class="dr-nav" data-dir="1" aria-label="Next month">›</button>
      </div>
      <div class="dr-grid"></div>
      <div class="dr-ft">
        <button type="button" class="dr-clear">Clear</button>
        <button type="button" class="dr-done">Done</button>
      </div>
    </div>
  `;
  const trigger = root.querySelector('.dr-trigger');
  const pop = root.querySelector('.dr-pop');
  const label = root.querySelector('.dr-label');
  const grid = root.querySelector('.dr-grid');

  const pad = n => String(n).padStart(2, '0');
  const toISO = (y, m, d) => `${y}-${pad(m + 1)}-${pad(d)}`;
  const fmt = iso => new Date(iso + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });

  function updateTrigger() {
    if (start && end && start !== end) trigger.textContent = `${fmt(start)} – ${fmt(end)}`;
    else if (start) trigger.textContent = fmt(start);
    else trigger.textContent = 'Pick delivery dates';
  }

  function renderGrid() {
    label.textContent = new Date(viewYear, viewMonth, 1).toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
    const firstDow = (new Date(viewYear, viewMonth, 1).getDay() + 6) % 7; // Monday = 0
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
    const todayIso = toISO(today.getFullYear(), today.getMonth(), today.getDate());
    let html = `<div class="dr-dow">${['M', 'T', 'W', 'T', 'F', 'S', 'S'].map(d => `<span>${d}</span>`).join('')}</div><div class="dr-days">`;
    for (let i = 0; i < firstDow; i++) html += `<span class="dr-day dr-pad"></span>`;
    for (let day = 1; day <= daysInMonth; day++) {
      const iso = toISO(viewYear, viewMonth, day);
      const classes = ['dr-day'];
      if (iso === start || iso === end) classes.push('dr-sel');
      else if (start && end && iso > start && iso < end) classes.push('dr-in');
      if (iso === todayIso) classes.push('dr-today');
      html += `<button type="button" class="${classes.join(' ')}" data-date="${iso}">${day}</button>`;
    }
    grid.innerHTML = html + `</div>`;
  }

  function open() { renderGrid(); pop.hidden = false; }
  function close() { pop.hidden = true; }

  trigger.addEventListener('click', (e) => { e.stopPropagation(); pop.hidden ? open() : close(); });
  document.addEventListener('click', (e) => { if (!pop.hidden && !root.contains(e.target)) close(); });
  root.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !pop.hidden) { close(); trigger.focus(); } });

  root.querySelector('.dr-hd').addEventListener('click', (e) => {
    const btn = e.target.closest('.dr-nav');
    if (!btn) return;
    viewMonth += parseInt(btn.dataset.dir, 10);
    if (viewMonth < 0) { viewMonth = 11; viewYear--; }
    if (viewMonth > 11) { viewMonth = 0; viewYear++; }
    renderGrid();
  });

  grid.addEventListener('click', (e) => {
    const btn = e.target.closest('.dr-day:not(.dr-pad)');
    if (!btn) return;
    const iso = btn.dataset.date;
    if (!start || (start && end)) { start = iso; end = null; }
    else if (iso < start) { end = start; start = iso; }
    else { end = iso; }
    updateTrigger();
    renderGrid();
    onChange && onChange({ start, end });
  });

  root.querySelector('.dr-clear').addEventListener('click', () => {
    start = null; end = null;
    updateTrigger();
    renderGrid();
    onChange && onChange({ start, end });
  });
  root.querySelector('.dr-done').addEventListener('click', close);

  updateTrigger();

  return {
    getRange: () => ({ start, end }),
    setRange: (s, e) => { start = s || null; end = e || null; updateTrigger(); if (!pop.hidden) renderGrid(); },
    clear: () => { start = null; end = null; updateTrigger(); },
    close
  };
}
