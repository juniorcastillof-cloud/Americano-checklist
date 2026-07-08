/**
 * app.js
 * Main application controller.
 */

const state = {
  phaseId: null,
  area: '',
  inspector: '',
  date: '',
  generalNotes: '',
  itemStates: {}, // itemId -> { status, notes, notesHistory, correctiveAction, locked, photoIds }
};

// ---------------------------------------------------------------
// Init
// ---------------------------------------------------------------
async function initApp() {
  document.getElementById('fProject').value =
    `${CHECKLIST_DATA.project.name_en} / ${CHECKLIST_DATA.project.name_es}`;
  document.getElementById('fSubcontractor').value =
    `${CHECKLIST_DATA.project.subcontractor_en}`;
  document.getElementById('disclaimerEn').textContent = CHECKLIST_DATA.disclaimer_en;
  document.getElementById('disclaimerEs').textContent = CHECKLIST_DATA.disclaimer_es;
  document.getElementById('fDate').value = new Date().toISOString().slice(0, 10);

  const phaseSelect = document.getElementById('fPhase');
  phaseSelect.innerHTML = CHECKLIST_DATA.phases
    .map((p) => `<option value="${p.id}">${p.number}. ${p.name_en} / ${p.name_es}</option>`)
    .join('');
  state.phaseId = CHECKLIST_DATA.phases[0].id;

  const savedInspector = await Storage.getMeta('lastInspector');
  if (savedInspector) {
    document.getElementById('fInspector').value = savedInspector;
    state.inspector = savedInspector;
  }

  setLang('en');
  await loadItemStatesForCurrentPhaseArea();
  renderChecklist();
  updateSubmitBar();

  // Microsoft Graph — restore prior session silently if the SDK is present and configured.
  if (typeof msal !== 'undefined') {
    try {
      const signedIn = await GraphSync.init();
      if (signedIn) markSyncConnected();
    } catch (e) {
      console.warn('Graph init skipped:', e.message);
    }
  }
  refreshSyncLabel();
}

// ---------------------------------------------------------------
// Phase / Area handling
// ---------------------------------------------------------------
async function onPhaseOrAreaChange() {
  state.phaseId = document.getElementById('fPhase').value;
  state.area = document.getElementById('fArea').value.trim();
  await loadItemStatesForCurrentPhaseArea();
  renderChecklist();
  updateSubmitBar();
}

async function loadItemStatesForCurrentPhaseArea() {
  state.itemStates = {};
  if (!state.area) return; // no area typed yet -> nothing to load
  const rows = await Storage.getItemsForPhaseArea(state.phaseId, state.area);
  rows.forEach((r) => {
    state.itemStates[r.itemId] = {
      status: r.status || null,
      notes: '',
      notesHistory: r.notesHistory || (r.notes ? [{ text: r.notes, ts: r.timestamp, by: r.inspector }] : []),
      correctiveAction: r.correctiveAction || '',
      locked: !!r.locked,
      photoIds: r.photoIds || [],
    };
  });
}

function currentPhase() {
  return CHECKLIST_DATA.phases.find((p) => p.id === state.phaseId);
}

// ---------------------------------------------------------------
// Rendering: checklist
// ---------------------------------------------------------------
function renderChecklist() {
  const phase = currentPhase();
  const container = document.getElementById('itemsContainer');
  if (!phase) { container.innerHTML = ''; return; }

  if (!state.area) {
    container.innerHTML = `<div class="card muted">${
      currentLang === 'es'
        ? 'Ingrese un Área/Ubicación arriba para cargar la lista de verificación.'
        : 'Enter an Area/Location above to load the checklist.'
    }</div>`;
    updateProgress();
    return;
  }

  container.innerHTML = phase.items.map((item) => renderItemCard(item)).join('');
  updateProgress();
  hydratePhotoThumbnails();
}

function renderItemCard(item) {
  const s = state.itemStates[item.id] || { status: null, notes: '', notesHistory: [], correctiveAction: '', locked: false, photoIds: [] };
  const isLockedPass = s.locked && s.status === 'pass';
  const showCorrective = s.status === 'fail';

  const priorNotesHtml = s.notesHistory.length
    ? `<div class="prior-note"><b>${currentLang === 'es' ? 'Notas anteriores' : 'Prior notes'}:</b> ${
        s.notesHistory.map((n) => escapeHtml(n.text)).join(' · ')
      }</div>`
    : '';

  const photosHtml = (s.photoIds || [])
    .map((pid) => `<img class="photo-thumb" data-photo-id="${pid}" src="" alt="photo">`)
    .join('');

  return `
    <div class="item-card ${isLockedPass ? 'locked-pass' : ''} ${showCorrective ? 'show-corrective' : ''}" data-item-id="${item.id}">
      ${isLockedPass ? `
        <div class="locked-banner">
          <span>✓ ${currentLang === 'es' ? 'Aprobado en inspección anterior' : 'Passed in prior inspection'}</span>
          <button onclick="reopenItem('${item.id}')">${currentLang === 'es' ? 'Reabrir' : 'Reopen'}</button>
        </div>` : ''}
      <p class="item-card__text-en">${escapeHtml(item.en)}</p>
      <p class="item-card__text-es">${escapeHtml(item.es)}</p>
      ${priorNotesHtml}
      <div class="status-row">
        ${CHECKLIST_DATA.statusOptions.map((opt) => `
          <button class="status-btn ${s.status === opt.id ? 'active' : ''}" data-status="${opt.id}"
            onclick="selectStatus('${item.id}','${opt.id}')" ${isLockedPass ? 'disabled' : ''}>
            ${currentLang === 'es' ? opt.label_es : opt.label_en}
          </button>
        `).join('')}
      </div>
      <textarea class="notes" placeholder="${currentLang === 'es' ? 'Notas de seguimiento para esta ronda...' : 'Follow-up notes for this round...'}"
        oninput="updateNotes('${item.id}', this.value)" ${isLockedPass ? 'disabled' : ''}>${escapeHtml(s.notes || '')}</textarea>
      <div class="corrective-wrap">
        <label>${currentLang === 'es' ? 'Acción correctiva requerida' : 'Corrective action required'}</label>
        <textarea oninput="updateCorrective('${item.id}', this.value)">${escapeHtml(s.correctiveAction || '')}</textarea>
      </div>
      <div class="photo-row">
        ${!isLockedPass ? `
        <button class="photo-btn" onclick="document.getElementById('file_${item.id}').click()">
          📷 ${currentLang === 'es' ? 'Agregar foto' : 'Add photo'}
        </button>
        <input type="file" id="file_${item.id}" accept="image/*" class="hidden"
          onchange="addPhoto('${item.id}', this)">` : ''}
        ${photosHtml}
        ${s.photoIds && s.photoIds.length ? `<span class="photo-count">${s.photoIds.length} ${currentLang === 'es' ? 'foto(s)' : 'photo(s)'}</span>` : ''}
      </div>
    </div>
  `;
}

// Fill in photo thumbnails asynchronously (IndexedDB blobs -> object URLs)
async function hydratePhotoThumbnails() {
  const imgs = document.querySelectorAll('.photo-thumb[data-photo-id]');
  for (const img of imgs) {
    const pid = img.getAttribute('data-photo-id');
    if (!pid) continue;
    const rec = await Storage.getPhoto(pid);
    if (rec && rec.blob) img.src = URL.createObjectURL(rec.blob);
  }
}

function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---------------------------------------------------------------
// Item state mutations
// ---------------------------------------------------------------
function ensureItemState(itemId) {
  if (!state.itemStates[itemId]) {
    state.itemStates[itemId] = { status: null, notes: '', notesHistory: [], correctiveAction: '', locked: false, photoIds: [] };
  }
  return state.itemStates[itemId];
}

function selectStatus(itemId, statusId) {
  const s = ensureItemState(itemId);
  s.status = statusId;
  renderChecklist();
  hydratePhotoThumbnails();
  updateSubmitBar();
}

function updateNotes(itemId, value) {
  ensureItemState(itemId).notes = value;
}

function updateCorrective(itemId, value) {
  ensureItemState(itemId).correctiveAction = value;
}

function reopenItem(itemId) {
  const s = ensureItemState(itemId);
  s.locked = false;
  s.status = null;
  renderChecklist();
  hydratePhotoThumbnails();
  updateSubmitBar();
}

async function addPhoto(itemId, inputEl) {
  const file = inputEl.files[0];
  if (!file) return;
  const s = ensureItemState(itemId);
  const photoId = `${itemId}_${Date.now()}`;
  await Storage.putPhoto(photoId, file, { itemId, phaseId: state.phaseId, area: state.area, filename: file.name || `${photoId}.jpg` });
  s.photoIds = [...(s.photoIds || []), photoId];
  renderChecklist();
  hydratePhotoThumbnails();
  showToast(currentLang === 'es' ? 'Foto agregada' : 'Photo added', 'success');
}

// ---------------------------------------------------------------
// Progress + submit bar
// ---------------------------------------------------------------
function updateProgress() {
  const phase = currentPhase();
  if (!phase || !state.area) {
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('progressLabel').textContent = '0%';
    document.getElementById('progressDetail').textContent = '';
    return;
  }
  const items = phase.items;
  const passed = items.filter((i) => state.itemStates[i.id] && state.itemStates[i.id].status === 'pass').length;
  const naCount = items.filter((i) => state.itemStates[i.id] && state.itemStates[i.id].status === 'na').length;
  const applicable = items.length - naCount;
  const pct = applicable > 0 ? Math.round((passed / applicable) * 100) : 0;
  document.getElementById('progressFill').style.width = `${pct}%`;
  document.getElementById('progressLabel').textContent = `${pct}%`;
  document.getElementById('progressDetail').textContent =
    currentLang === 'es'
      ? `${passed} de ${applicable} elementos aplicables aprobados`
      : `${passed} of ${applicable} applicable items passed`;
}

function updateSubmitBar() {
  const phase = currentPhase();
  const btn = document.getElementById('submitBtn');
  const summary = document.getElementById('submitSummary');

  if (!phase || !state.area || !document.getElementById('fInspector').value.trim()) {
    btn.disabled = true;
    summary.textContent = currentLang === 'es'
      ? 'Ingrese inspector y área para continuar.'
      : 'Enter inspector and area to continue.';
    return;
  }

  const unset = phase.items.filter((i) => !state.itemStates[i.id] || !state.itemStates[i.id].status);
  if (unset.length > 0) {
    btn.disabled = true;
    summary.textContent = currentLang === 'es'
      ? `${unset.length} elemento(s) sin estado asignado.`
      : `${unset.length} item(s) still need a status.`;
    return;
  }

  const failCount = phase.items.filter((i) => state.itemStates[i.id].status === 'fail').length;
  const pendingCount = phase.items.filter((i) => state.itemStates[i.id].status === 'pending').length;
  btn.disabled = false;
  if (failCount === 0 && pendingCount === 0) {
    summary.textContent = currentLang === 'es'
      ? 'Todos los elementos aprobados — listo para enviar como completo.'
      : 'All items passed — ready to submit as complete.';
  } else {
    summary.textContent = currentLang === 'es'
      ? `${failCount} rechazado(s), ${pendingCount} pendiente(s) — se guardará como incompleto.`
      : `${failCount} failed, ${pendingCount} pending — will save as incomplete.`;
  }
}

// ---------------------------------------------------------------
// Submit
// ---------------------------------------------------------------
async function submitInspection() {
  const phase = currentPhase();
  const inspector = document.getElementById('fInspector').value.trim();
  const date = document.getElementById('fDate').value;
  const generalNotes = document.getElementById('fGeneralNotes').value.trim();
  state.inspector = inspector;
  state.generalNotes = generalNotes;
  await Storage.setMeta('lastInspector', inspector);

  const timestamp = new Date().toISOString();
  const itemResults = [];

  for (const item of phase.items) {
    const s = state.itemStates[item.id];
    const newNotesHistory = [...(s.notesHistory || [])];
    if (s.notes && s.notes.trim()) {
      newNotesHistory.push({ text: s.notes.trim(), ts: timestamp, by: inspector });
    }
    const shouldLock = s.status === 'pass' || s.status === 'na';
    const photoRef = (s.photoIds || []).length ? `${s.photoIds.length} photo(s) attached in app` : '';

    const record = {
      status: s.status,
      notes: s.notes && s.notes.trim() ? s.notes.trim() : (newNotesHistory.length ? newNotesHistory[newNotesHistory.length - 1].text : ''),
      notesHistory: newNotesHistory,
      correctiveAction: s.status === 'fail' ? (s.correctiveAction || '') : '',
      locked: shouldLock,
      photoIds: s.photoIds || [],
      inspector,
      timestamp,
      submittedBy: inspector,
    };
    await Storage.putItem(state.phaseId, state.area, item.id, record);
    itemResults.push({ itemId: item.id, ...record, photoRef });

    // Optional: push photos straight to OneDrive if connected, so they live
    // alongside the workbook rather than only on this device.
    if (GraphSync.isSignedIn() && (s.photoIds || []).length) {
      for (const pid of s.photoIds) {
        try {
          const photo = await Storage.getPhoto(pid);
          if (photo && photo.blob) {
            await GraphSync.uploadPhoto(photo.blob, photo.filename || `${pid}.jpg`);
          }
        } catch (e) {
          console.warn('Photo upload skipped:', e.message);
        }
      }
    }
    // Reflect what was just saved back into live state so the UI (prior-notes
    // box, lock state) is correct immediately, without waiting for a reload.
    state.itemStates[item.id].notes = '';
    state.itemStates[item.id].notesHistory = newNotesHistory;
    state.itemStates[item.id].correctiveAction = record.correctiveAction;
    state.itemStates[item.id].locked = shouldLock;
  }

  const failCount = itemResults.filter((r) => r.status === 'fail').length;
  const pendingCount = itemResults.filter((r) => r.status === 'pending').length;
  const overallStatus = (failCount === 0 && pendingCount === 0) ? 'complete' : 'incomplete';

  await Storage.addRecord({
    phaseId: state.phaseId,
    area: state.area,
    inspector,
    date,
    generalNotes,
    itemResults,
    timestamp,
    submittedBy: inspector,
    overallStatus,
  });

  await syncWorkbook();

  renderChecklist();
  hydratePhotoThumbnails();
  updateSubmitBar();
  updateProgress();

  const msg = overallStatus === 'complete'
    ? (currentLang === 'es' ? 'Inspección enviada como completa.' : 'Inspection submitted as complete.')
    : (currentLang === 'es' ? 'Inspección guardada como incompleta — quedan elementos por reinspeccionar.' : 'Inspection saved as incomplete — items remain for reinspection.');
  showToast(msg, overallStatus === 'complete' ? 'success' : 'error');
}

async function syncWorkbook() {
  try {
    const buffer = await ExportModule.buildWorkbookArrayBuffer();
    if (GraphSync.isSignedIn()) {
      await GraphSync.uploadWorkbook(buffer);
      showToast(currentLang === 'es' ? 'Libro de Excel sincronizado con OneDrive.' : 'Excel workbook synced to OneDrive.', 'success');
    } else {
      await ExportModule.downloadWorkbookLocally();
      showToast(
        currentLang === 'es'
          ? 'OneDrive no conectado — libro descargado localmente. Súbalo manualmente a la carpeta de OneDrive.'
          : 'OneDrive not connected — workbook downloaded locally. Please upload it to the OneDrive folder manually.',
        'error'
      );
    }
  } catch (e) {
    console.error(e);
    showToast((currentLang === 'es' ? 'Error al exportar: ' : 'Export error: ') + e.message, 'error');
  }
}

function downloadNow() {
  ExportModule.downloadWorkbookLocally();
}

// ---------------------------------------------------------------
// OneDrive sync bar
// ---------------------------------------------------------------
function markSyncConnected() {
  document.getElementById('syncDot').classList.add('connected');
  document.getElementById('syncDot').classList.remove('error');
}

function refreshSyncLabel() {
  const label = document.getElementById('syncLabel');
  const btn = document.getElementById('syncBtn');
  if (!GraphSync.isConfigured()) {
    label.textContent = currentLang === 'es' ? 'OneDrive no configurado (ver README)' : 'OneDrive not configured (see README)';
    btn.textContent = t('connect_onedrive');
    return;
  }
  if (GraphSync.isSignedIn()) {
    label.textContent = `${t('connected_as')}: ${GraphSync.getAccountName()}`;
    btn.textContent = t('disconnect');
    document.getElementById('syncDot').classList.add('connected');
  } else {
    label.textContent = t('not_connected');
    btn.textContent = t('connect_onedrive');
    document.getElementById('syncDot').classList.remove('connected');
  }
}

async function handleSyncButton() {
  try {
    if (GraphSync.isSignedIn()) {
      await GraphSync.signOut();
    } else {
      await GraphSync.signIn();
      showToast(currentLang === 'es' ? 'Conectado a OneDrive.' : 'Connected to OneDrive.', 'success');
    }
  } catch (e) {
    showToast((currentLang === 'es' ? 'Error de conexión: ' : 'Connection error: ') + e.message, 'error');
  }
  refreshSyncLabel();
}

// ---------------------------------------------------------------
// Views / tabs
// ---------------------------------------------------------------
function showView(view) {
  document.getElementById('viewChecklist').classList.toggle('hidden', view !== 'checklist');
  document.getElementById('viewSummary').classList.toggle('hidden', view !== 'summary');
  document.getElementById('tabChecklistBtn').classList.toggle('active', view === 'checklist');
  document.getElementById('tabSummaryBtn').classList.toggle('active', view === 'summary');
  if (view === 'summary') renderSummary();
}

async function renderSummary() {
  const allItems = await Storage.getAllItems();
  const allRecords = await Storage.getAllRecords();

  const total = allItems.length;
  const passed = allItems.filter((i) => i.status === 'pass').length;
  const failed = allItems.filter((i) => i.status === 'fail').length;
  const pending = allItems.filter((i) => i.status === 'pending').length;
  const na = allItems.filter((i) => i.status === 'na').length;
  const applicable = total - na;
  const pct = applicable > 0 ? Math.round((passed / applicable) * 100) : 0;

  document.getElementById('statTotal').textContent = total;
  document.getElementById('statPass').textContent = passed;
  document.getElementById('statFail').textContent = failed;
  document.getElementById('statPending').textContent = pending;
  document.getElementById('summaryProgressFill').style.width = `${pct}%`;
  document.getElementById('summaryProgressLabel').textContent = `${pct}%`;

  const lastDate = allRecords.length ? allRecords[allRecords.length - 1].timestamp : null;
  document.getElementById('lastInspectionLabel').textContent = lastDate
    ? (currentLang === 'es' ? `Última inspección: ${new Date(lastDate).toLocaleString()}` : `Last inspection: ${new Date(lastDate).toLocaleString()}`)
    : (currentLang === 'es' ? 'Aún no hay inspecciones registradas.' : 'No inspections recorded yet.');

  const list = document.getElementById('phaseOpenList');
  list.innerHTML = CHECKLIST_DATA.phases.map((p) => {
    const openCount = allItems.filter((i) => i.phaseId === p.id && (i.status === 'fail' || i.status === 'pending')).length;
    return `<div class="phase-open-row">
      <span>${p.number}. ${currentLang === 'es' ? p.name_es : p.name_en}</span>
      <span class="badge ${openCount === 0 ? 'zero' : ''}">${openCount}</span>
    </div>`;
  }).join('');
}

async function resetLocalData() {
  const msg = currentLang === 'es'
    ? '¿Está seguro? Esto eliminará todos los datos de inspección guardados en este dispositivo (no afecta a OneDrive).'
    : 'Are you sure? This will delete all inspection data stored on this device (does not affect OneDrive).';
  if (!confirm(msg)) return;
  if (!confirm(currentLang === 'es' ? 'Esta acción no se puede deshacer. Confirmar de nuevo.' : 'This cannot be undone. Confirm again.')) return;
  await Storage.clearAll();
  location.reload();
}

// ---------------------------------------------------------------
// Toast
// ---------------------------------------------------------------
let toastTimer = null;
function showToast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3800);
}

// ---------------------------------------------------------------
// Field listeners
// ---------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initApp();
  document.getElementById('fInspector').addEventListener('input', updateSubmitBar);
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
});
