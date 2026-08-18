// Discard-unsaved-changes guard for <dialog> elements. Attaches once per
// dialog; intercepts every close attempt (explicit Cancel button, ESC via
// the native 'cancel' event, and backdrop click) through the same
// dirty-check, so there's exactly one confirm/discard code path, not one per
// dialog. Dirtiness defaults to "do any input/select/textarea inside the
// dialog differ from their values when arm() was last called"; pass
// isDirtyOverride for dialogs whose unsaved state isn't plain form fields.
//
// Programmatic dialogEl.close() calls (e.g. after a successful save) are left
// completely alone — native .close() only fires a 'close' event, never
// 'cancel', so this guard never intercepts them.
function attachDiscardGuard(dialogEl, opts = {}) {
  const scope = opts.scope || dialogEl;
  const isDirtyOverride = opts.isDirtyOverride || null;
  let snapshot = null;

  function fieldEls() {
    return [...scope.querySelectorAll('input, select, textarea')].filter(el => el.id && !el.disabled);
  }
  function snapshotNow() {
    const s = {};
    fieldEls().forEach(el => { s[el.id] = el.type === 'checkbox' ? el.checked : el.value; });
    return s;
  }
  function isDirty() {
    if (isDirtyOverride) return isDirtyOverride();
    if (!snapshot) return false;
    const now = snapshotNow();
    return Object.keys(snapshot).some(id => snapshot[id] !== now[id]);
  }
  function arm() { if (!isDirtyOverride) snapshot = snapshotNow(); }

  function ensurePanel() {
    let panel = dialogEl.querySelector(':scope > .discard-panel');
    if (panel) return panel;
    dialogEl.insertAdjacentHTML('beforeend', `
      <div class="discard-panel">
        <div class="box">
          <h3>Discard unsaved changes?</h3>
          <p>You have edits that haven't been saved. Closing now will lose them.</p>
          <div class="actions">
            <button type="button" class="btn ghost sm" data-role="keep-editing">Keep editing</button>
            <button type="button" class="btn danger sm" data-role="discard">Discard changes</button>
          </div>
        </div>
      </div>`);
    panel = dialogEl.querySelector(':scope > .discard-panel');
    panel.querySelector('[data-role="keep-editing"]').addEventListener('click', () => panel.classList.remove('open'));
    panel.querySelector('[data-role="discard"]').addEventListener('click', () => {
      panel.classList.remove('open');
      dialogEl.close();
    });
    return panel;
  }

  function guardedClose() {
    if (isDirty()) { ensurePanel().classList.add('open'); return; }
    dialogEl.close();
  }

  dialogEl.addEventListener('cancel', (e) => {
    if (isDirty()) { e.preventDefault(); ensurePanel().classList.add('open'); }
  });
  dialogEl.addEventListener('click', (e) => {
    if (e.target === dialogEl) guardedClose(); // backdrop click
  });

  return { arm, guardedClose, isDirty };
}
