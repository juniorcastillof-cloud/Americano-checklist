/**
 * export.js
 * Builds the Excel workbook (via SheetJS) from everything currently in the
 * local IndexedDB database. This is re-run in full on every submit, so the
 * workbook on OneDrive is always a complete, current snapshot — not an
 * incremental patch.
 *
 * Sheets produced:
 *   1. "Current Status"  - one row per checklist item per Phase+Area, latest state.
 *                           This is the punch-list view (Pass/Fail/Pending/N/A).
 *   2. "Inspection Log"  - one row per submission event, full audit trail with
 *                           bilingual item text, notes, corrective action, inspector, timestamp.
 *   3. "Summary"         - totals, percent complete, open items by phase, last inspection date.
 */

const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: 'FF1B2A4A' } };
const HEADER_FONT = { bold: true, color: { rgb: 'FFFFFFFF' } };

function styleHeaderRow(ws, rowIndex, colCount) {
  for (let c = 0; c < colCount; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c });
    if (!ws[cellRef]) continue;
    ws[cellRef].s = { font: HEADER_FONT, fill: HEADER_FILL, alignment: { vertical: 'center', wrapText: true } };
  }
}

function statusLabel(statusId) {
  const opt = CHECKLIST_DATA.statusOptions.find((s) => s.id === statusId);
  return opt ? `${opt.label_en} / ${opt.label_es}` : statusId || '';
}

function findItemText(phaseId, itemId) {
  const phase = CHECKLIST_DATA.phases.find((p) => p.id === phaseId);
  if (!phase) return { en: '', es: '', phaseName_en: '', phaseName_es: '', phaseNumber: '' };
  const item = phase.items.find((i) => i.id === itemId);
  return {
    en: item ? item.en : '',
    es: item ? item.es : '',
    phaseName_en: phase.name_en,
    phaseName_es: phase.name_es,
    phaseNumber: phase.number,
  };
}

async function buildWorkbook() {
  const wb = XLSX.utils.book_new();

  const allItems = await Storage.getAllItems();
  const allRecords = await Storage.getAllRecords();

  // ---------- Sheet 1: Current Status ----------
  const statusHeader = [
    'Phase #', 'Phase (EN)', 'Phase (ES)', 'Area/Location',
    'Checklist Item (EN)', 'Checklist Item (ES)',
    'Status', 'Notes', 'Corrective Action Required',
    'Inspector', 'Last Updated', 'Submitted By',
  ];
  const statusRows = allItems
    .map((it) => {
      const txt = findItemText(it.phaseId, it.itemId);
      return {
        row: [
          txt.phaseNumber, txt.phaseName_en, txt.phaseName_es, it.area,
          txt.en, txt.es,
          statusLabel(it.status), it.notes || '', it.correctiveAction || '',
          it.inspector || '', it.timestamp || '', it.submittedBy || '',
        ],
        sortKey: `${txt.phaseNumber}-${it.area}-${it.itemId}`,
      };
    })
    .sort((a, b) => a.sortKey.localeCompare(b.sortKey))
    .map((x) => x.row);

  const wsStatus = XLSX.utils.aoa_to_sheet([statusHeader, ...statusRows]);
  wsStatus['!cols'] = [
    { wch: 8 }, { wch: 26 }, { wch: 26 }, { wch: 18 },
    { wch: 42 }, { wch: 42 },
    { wch: 20 }, { wch: 30 }, { wch: 28 },
    { wch: 16 }, { wch: 20 }, { wch: 16 },
  ];
  styleHeaderRow(wsStatus, 0, statusHeader.length);
  XLSX.utils.book_append_sheet(wb, wsStatus, 'Current Status');

  // ---------- Sheet 2: Inspection Log (audit trail of every submission) ----------
  const logHeader = [
    'Timestamp', 'Project', 'Subcontractor', 'Phase #', 'Phase (EN)', 'Phase (ES)',
    'Area/Location', 'Inspector', 'Checklist Item (EN)', 'Checklist Item (ES)',
    'Status', 'Notes', 'Corrective Action Required', 'Photo Reference',
    'General Inspection Notes', 'Submitted By',
  ];
  const logRows = [];
  allRecords.forEach((rec) => {
    (rec.itemResults || []).forEach((ir) => {
      const txt = findItemText(rec.phaseId, ir.itemId);
      logRows.push([
        rec.timestamp, CHECKLIST_DATA.project.name_en, CHECKLIST_DATA.project.subcontractor_en,
        txt.phaseNumber, txt.phaseName_en, txt.phaseName_es,
        rec.area, rec.inspector, txt.en, txt.es,
        statusLabel(ir.status), ir.notes || '', ir.correctiveAction || '', ir.photoRef || '',
        rec.generalNotes || '', rec.submittedBy || '',
      ]);
    });
  });
  const wsLog = XLSX.utils.aoa_to_sheet([logHeader, ...logRows]);
  wsLog['!cols'] = logHeader.map(() => ({ wch: 22 }));
  styleHeaderRow(wsLog, 0, logHeader.length);
  XLSX.utils.book_append_sheet(wb, wsLog, 'Inspection Log');

  // ---------- Sheet 3: Summary ----------
  const total = allItems.length;
  const passed = allItems.filter((i) => i.status === 'pass').length;
  const failed = allItems.filter((i) => i.status === 'fail').length;
  const pending = allItems.filter((i) => i.status === 'pending').length;
  const na = allItems.filter((i) => i.status === 'na').length;
  const applicable = total - na;
  const pctComplete = applicable > 0 ? passed / applicable : 0;
  const lastDate = allRecords.length ? allRecords[allRecords.length - 1].timestamp : '';

  const summaryTop = [
    ['Verizon North Royalton UPS - Inspection Summary', ''],
    ['Generated', new Date().toISOString()],
    ['Subcontractor', CHECKLIST_DATA.project.subcontractor_en],
    ['General Contractor', CHECKLIST_DATA.project.general_contractor_en],
    [''],
    ['Total Checklist Items Tracked', total],
    ['Passed', passed],
    ['Failed', failed],
    ['Pending', pending],
    ['Not Applicable', na],
    ['Percent Complete (of applicable items)', pctComplete],
    ['Last Inspection Date', lastDate],
    [''],
    ['Open Items by Phase', ''],
    ['Phase', 'Open (Fail + Pending)'],
  ];
  CHECKLIST_DATA.phases.forEach((p) => {
    const openCount = allItems.filter(
      (i) => i.phaseId === p.id && (i.status === 'fail' || i.status === 'pending')
    ).length;
    summaryTop.push([`${p.number}. ${p.name_en} / ${p.name_es}`, openCount]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryTop);
  wsSummary['!cols'] = [{ wch: 50 }, { wch: 20 }];
  // percent format on the "Percent Complete" row
  const pctRowIndex = summaryTop.findIndex((r) => r[0] === 'Percent Complete (of applicable items)');
  const pctCell = XLSX.utils.encode_cell({ r: pctRowIndex, c: 1 });
  if (wsSummary[pctCell]) wsSummary[pctCell].z = '0.0%';
  styleHeaderRow(wsSummary, 0, 2);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  return wb;
}

/** Returns the workbook as an ArrayBuffer, ready to upload to OneDrive or download locally. */
async function buildWorkbookArrayBuffer() {
  const wb = await buildWorkbook();
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
}

/** Triggers a local file download (fallback path when OneDrive isn't connected). */
async function downloadWorkbookLocally() {
  const wb = await buildWorkbook();
  XLSX.writeFile(wb, GraphSync.config.workbookName);
}

window.ExportModule = { buildWorkbook, buildWorkbookArrayBuffer, downloadWorkbookLocally };
