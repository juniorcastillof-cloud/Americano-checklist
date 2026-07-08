/**
 * i18n.js
 * Translations for app CHROME (buttons, labels, headings). The checklist
 * item text itself is always shown in both languages simultaneously
 * (see checklist-data.js) per the bilingual requirement — this toggle only
 * controls which language leads in labels/buttons/status names for a
 * reader who prefers one language, plus sets the primary <html lang>.
 */

const I18N = {
  en: {
    hdrTitle: 'Field Inspection Checklist',
    tab_checklist: 'Checklist',
    tab_summary: 'Summary',
    setup_heading: 'Inspection Setup',
    lbl_project: 'Project',
    lbl_subcontractor: 'Subcontractor',
    lbl_phase: 'Inspection Phase',
    lbl_area: 'Area / Location',
    lbl_inspector: 'Inspector Name',
    lbl_date: 'Inspection Date',
    progress_heading: 'Progress — This Phase / Area',
    general_notes_heading: 'General Inspection Notes',
    stat_total: 'Total Items',
    stat_pass: 'Passed',
    stat_fail: 'Failed',
    stat_pending: 'Pending',
    stat_pct: 'Percent Complete',
    stat_open_by_phase: 'Open Items by Phase',
    export_heading: 'Export',
    export_note: 'Download a full copy of the current workbook (all phases, all recorded inspections) at any time.',
    btn_download_now: 'Download Excel Workbook Now',
    btn_reset: 'Reset Local Data…',
    btn_submit: 'Submit Inspection',
    connect_onedrive: 'Connect OneDrive',
    disconnect: 'Disconnect',
    not_connected: 'Not connected to OneDrive',
    connected_as: 'Connected as',
  },
  es: {
    hdrTitle: 'Lista de Verificación de Inspección de Campo',
    tab_checklist: 'Lista',
    tab_summary: 'Resumen',
    setup_heading: 'Configuración de la Inspección',
    lbl_project: 'Proyecto',
    lbl_subcontractor: 'Subcontratista',
    lbl_phase: 'Fase de Inspección',
    lbl_area: 'Área / Ubicación',
    lbl_inspector: 'Nombre del Inspector',
    lbl_date: 'Fecha de Inspección',
    progress_heading: 'Progreso — Esta Fase / Área',
    general_notes_heading: 'Notas Generales de la Inspección',
    stat_total: 'Total de Elementos',
    stat_pass: 'Aprobados',
    stat_fail: 'Rechazados',
    stat_pending: 'Pendientes',
    stat_pct: 'Porcentaje Completo',
    stat_open_by_phase: 'Elementos Abiertos por Fase',
    export_heading: 'Exportar',
    export_note: 'Descargue una copia completa del libro actual (todas las fases, todas las inspecciones registradas) en cualquier momento.',
    btn_download_now: 'Descargar Libro de Excel Ahora',
    btn_reset: 'Reiniciar Datos Locales…',
    btn_submit: 'Enviar Inspección',
    connect_onedrive: 'Conectar OneDrive',
    disconnect: 'Desconectar',
    not_connected: 'No conectado a OneDrive',
    connected_as: 'Conectado como',
  },
};

let currentLang = 'en';

function setLang(lang) {
  currentLang = lang;
  document.documentElement.lang = lang;
  document.getElementById('langEnBtn').classList.toggle('active', lang === 'en');
  document.getElementById('langEsBtn').classList.toggle('active', lang === 'es');
  applyTranslations();
}

function t(key) {
  return (I18N[currentLang] && I18N[currentLang][key]) || I18N.en[key] || key;
}

function applyTranslations() {
  document.querySelectorAll('[data-t]').forEach((el) => {
    el.textContent = t(el.getAttribute('data-t'));
  });
  document.getElementById('hdrTitle').textContent = t('hdrTitle');
  if (typeof refreshSyncLabel === 'function') refreshSyncLabel();
  if (typeof renderChecklist === 'function') renderChecklist();
  if (typeof renderSummary === 'function') renderSummary();
}
