
const STORAGE_KEY = "inserimento_lavori_senza_excel_v3";

const DEFAULT_JOBS = [
  { id: "1", jobCode: "31030941", client: "FAMULARI GIUSEPPE", type: "RIPARAZIONE", operator: "Giuseppe", listPrice: 41, appliedPrice: 41, quantity: 1, standardDays: 1, entryDate: "2026-03-31", deliveryDate: "2026-03-31", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Riparazione", pieces: 1 },
  { id: "2", jobCode: "01040937", client: "STUDIO DENT. VILLA DANTE DTT.RI METRO E ZA", type: "MASCHERIBA BREGA", operator: "Laboratorio", listPrice: 20, appliedPrice: 20, quantity: 1, standardDays: 1, entryDate: "2026-03-31", deliveryDate: "2026-03-31", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Lavorazione standard", pieces: 1 },
  { id: "3", jobCode: "01041111", client: "BARRESI ANTONIO", type: "CORONA IN RESINA PROVVISORIA", operator: "Laboratorio", listPrice: 180, appliedPrice: 180, quantity: 1, standardDays: 2, entryDate: "2026-03-31", deliveryDate: "2026-03-31", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Provvisorio", pieces: 1 },
  { id: "4", jobCode: "30031626", client: "GDENTAL CLINIC SRLS", type: "CARICO IMMEDIATO", operator: "Laboratorio", listPrice: 300, appliedPrice: 300, quantity: 1, standardDays: 1, entryDate: "2026-03-30", deliveryDate: "2026-03-30", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Protocollo urgente", pieces: 4 },
  { id: "5", jobCode: "30031415", client: "FAMULARI GIUSEPPE", type: "RIOARAZIONE", operator: "Giuseppe", listPrice: 41, appliedPrice: 41, quantity: 1, standardDays: 1, entryDate: "2026-03-30", deliveryDate: "2026-03-30", trialOutDate: "", trialReturnDate: "", urgent: false, redo: false, status: "Consegnato", note: "", priceInfo: "Riparazione", pieces: 1 },
  { id: "6", jobCode: "30031124-A", client: "FAMULARI GIUSEPPE", type: "RIBASATURA PARZIALE", operator: "Giuseppe", listPrice: 81.5, appliedPrice: 81.5, quantity: 1, standardDays: 1, entryDate: "2026-03-30", deliveryDate: "2026-03-30", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Ribasatura", pieces: 1 },
  { id: "7", jobCode: "30031238", client: "STUDIO DENT. VILLA DANTE DTT.RI METRO E ZA", type: "ELEMENTO IN ZIRCONIA MONOLITICA", operator: "Cad Cam", listPrice: 360, appliedPrice: 360, quantity: 1, standardDays: 2, entryDate: "2026-03-30", deliveryDate: "2026-03-31", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Zirconia", pieces: 1 },
  { id: "8", jobCode: "30031530", client: "LIFE S.R.L.", type: "ELEMENTO IN ZIRCONIA MONOLITICA", operator: "Cad Cam", listPrice: 120, appliedPrice: 120, quantity: 1, standardDays: 2, entryDate: "2026-03-30", deliveryDate: "2026-03-30", trialOutDate: "", trialReturnDate: "", urgent: false, redo: false, status: "Consegnato", note: "", priceInfo: "Zirconia", pieces: 1 },
  { id: "9", jobCode: "30031124-B", client: "FAMULARI GIUSEPPE", type: "AGG DENTE SU SAL", operator: "Laboratorio", listPrice: 81.5, appliedPrice: 81.5, quantity: 1, standardDays: 1, entryDate: "2026-03-30", deliveryDate: "", trialOutDate: "2026-03-31", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "", priceInfo: "Aggiunta", pieces: 1 },
  { id: "10", jobCode: "30031522", client: "STUDIO DENTISTICO DOTT.SSA CAFIERO ANGELA", type: "CORONA IN RESINA PROVVISORIA", operator: "Laboratorio", listPrice: 10, appliedPrice: 10, quantity: 1, standardDays: 1, entryDate: "2026-03-30", deliveryDate: "2026-03-31", trialOutDate: "", trialReturnDate: "", urgent: true, redo: false, status: "Consegnato", note: "", priceInfo: "Provvisorio", pieces: 1 },
  { id: "11", jobCode: "30032001", client: "GIORGIANNI", type: "ELEMENTO IN CR.CO - CERAMICA", operator: "Ceramica", listPrice: 220, appliedPrice: 220, quantity: 1, standardDays: 3, entryDate: "2026-03-27", deliveryDate: "", trialOutDate: "2026-03-27", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "", priceInfo: "Ceramica", pieces: 1 },
  { id: "12", jobCode: "30032002", client: "D'ANDREA MASSIMO", type: "ELEMENTO IN CR.CO - CERAMICA", operator: "Ceramica", listPrice: 220, appliedPrice: 220, quantity: 1, standardDays: 3, entryDate: "2026-03-27", deliveryDate: "", trialOutDate: "2026-03-27", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "", priceInfo: "Ceramica", pieces: 1 },
  { id: "13", jobCode: "30032003", client: "STUDIO M. D. DEI DOTT.RI ALBANESE S. E DI BE", type: "ELEMENTO IN CR.CO - CERAMICA", operator: "Ceramica", listPrice: 220, appliedPrice: 220, quantity: 1, standardDays: 3, entryDate: "2026-03-26", deliveryDate: "", trialOutDate: "2026-03-26", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "", priceInfo: "Ceramica", pieces: 1 },
  { id: "14", jobCode: "30032004", client: "GRASSO SANDRO", type: "ELEMENTO IN CR.CO - CERAMICA", operator: "Ceramica", listPrice: 220, appliedPrice: 220, quantity: 1, standardDays: 3, entryDate: "2026-03-20", deliveryDate: "", trialOutDate: "2026-03-20", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "", priceInfo: "Ceramica", pieces: 1 },
  { id: "15", jobCode: "30032005", client: "VELO ALESSIA", type: "SCHELETRATO SUPERIORE O INFERIORE", operator: "Scheletrica", listPrice: 480, appliedPrice: 450, quantity: 1, standardDays: 5, entryDate: "2026-03-26", deliveryDate: "", trialOutDate: "2026-03-26", trialReturnDate: "", urgent: false, redo: false, status: "In prova", note: "Sconto applicato", priceInfo: "Listino adattato", pieces: 2 }
];

const IMPORTED_JOBS = Array.isArray(window.DEFAULT_SPERANZA_JOBS) && window.DEFAULT_SPERANZA_JOBS.length
  ? window.DEFAULT_SPERANZA_JOBS.map((job, index) => ({
      id: job.id || `xlsx-${index + 1}`,
      jobCode: job.jobCode || "",
      client: job.client || "",
      type: job.type || "",
      operator: job.operator || "",
      listPrice: Number(job.listPrice || 0),
      appliedPrice: Number(job.appliedPrice || job.listPrice || 0),
      quantity: Number(job.quantity || 1),
      standardDays: Number(job.standardDays || 0),
      entryDate: job.entryDate || "",
      deliveryDate: job.deliveryDate || "",
      trialOutDate: job.trialOutDate || "",
      trialReturnDate: job.trialReturnDate || "",
      urgent: !!job.urgent,
      redo: !!job.redo,
      status: job.status || "In lavorazione",
      note: job.note || "",
      priceInfo: job.priceInfo || "",
      pieces: Number(job.pieces || job.quantity || 1)
    }))
  : DEFAULT_JOBS;

const ANALYSIS_MODES = ["Prova non rientrati","Consegnati nel periodo","SLA per studio","Distribuzione per medico","Distribuzione per lavorazione","Pareto clienti 80/20","Pareto lavorazioni 80/20","Numeri economici","Valore potenziale","Clienti piu' scontati","Lavorazioni per volume","Margine perso per lavorazione","Indicatori flusso"];
const state = {
  jobs: [],
  activeView: "operativita",
  activeAnalysis: ANALYSIS_MODES[0],
  analysisDetailRecords: [],
  analysisDetailFilterKey: "",
  analysisDetailFilterLabel: "",
  analysisDetailEmptyLabel: "Nessun dettaglio disponibile."
};
const viewButtons = Array.from(document.querySelectorAll(".view-tab"));
const viewPanels = { operativita: document.getElementById("view-operativita"), analisi: document.getElementById("view-analisi") };
const jobForm = document.getElementById("jobForm");
const jobIdHidden = document.getElementById("jobIdHidden");
const jobCodeInput = document.getElementById("jobCodeInput");
const jobClientInput = document.getElementById("jobClientInput");
const jobTypeInput = document.getElementById("jobTypeInput");
const jobOperatorInput = document.getElementById("jobOperatorInput");
const jobListPriceInput = document.getElementById("jobListPriceInput");
const jobStandardDaysInput = document.getElementById("jobStandardDaysInput");
const jobQuantityInput = document.getElementById("jobQuantityInput");
const jobAppliedPriceInput = document.getElementById("jobAppliedPriceInput");
const jobFinalPricePreview = document.getElementById("jobFinalPricePreview");
const jobStatusInput = document.getElementById("jobStatusInput");
const jobEntryDateInput = document.getElementById("jobEntryDateInput");
const jobDeliveryDateInput = document.getElementById("jobDeliveryDateInput");
const jobTrialOutDateInput = document.getElementById("jobTrialOutDateInput");
const jobTrialReturnDateInput = document.getElementById("jobTrialReturnDateInput");
const jobUrgentInput = document.getElementById("jobUrgentInput");
const jobRedoInput = document.getElementById("jobRedoInput");
const jobPriceInfoInput = document.getElementById("jobPriceInfoInput");
const jobNoteInput = document.getElementById("jobNoteInput");
const jobListTotalValue = document.getElementById("jobListTotalValue");
const jobAppliedTotalValue = document.getElementById("jobAppliedTotalValue");
const cancelEditButton = document.getElementById("cancelEditButton");
const exportCsvButton = document.getElementById("exportCsvButton");
const exportJsonButton = document.getElementById("exportJsonButton");
const resetArchiveButton = document.getElementById("resetArchiveButton");
const jobSearchInput = document.getElementById("jobSearchInput");
const jobClientFilterInput = document.getElementById("jobClientFilterInput");
const jobOperatorFilterInput = document.getElementById("jobOperatorFilterInput");
const jobUrgencyFilter = document.getElementById("jobUrgencyFilter");
const applyFiltersButton = document.getElementById("applyFiltersButton");
const resetFiltersButton = document.getElementById("resetFiltersButton");
const tableSearchInput = document.getElementById("tableSearchInput");
const filterStatus = document.getElementById("filterStatus");
const recentEntriesList = document.getElementById("recentEntriesList");
const jobTableBody = document.getElementById("jobTableBody");
const analysisStartInput = document.getElementById("analysisStartInput");
const analysisEndInput = document.getElementById("analysisEndInput");
const analysisKpiGrid = document.getElementById("analysisKpiGrid");
const analysisChipBar = document.getElementById("analysisChipBar");
const analysisTitle = document.getElementById("analysisTitle");
const analysisIntro = document.getElementById("analysisIntro");
const analysisPrimary = document.getElementById("analysisPrimary");
const analysisDetailList = document.getElementById("analysisDetailList");
const analysisDetailSearchInput = document.getElementById("analysisDetailSearchInput");
const clientOptions = document.getElementById("clientOptions");
const clientFilterOptions = document.getElementById("clientFilterOptions");
const workTypeOptions = document.getElementById("workTypeOptions");
const operatorOptions = document.getElementById("operatorOptions");

const CLEAN_IMPORTED_JOBS = IMPORTED_JOBS.map(normalizeJobRecord).filter(isUsefulJobRecord);
const CLIENT_CHOICES = getUniqueSorted(CLEAN_IMPORTED_JOBS.map((job) => job.client));
const WORK_TYPE_CHOICES = getUniqueSorted(CLEAN_IMPORTED_JOBS.map((job) => job.type));
const OPERATOR_CHOICES = getUniqueSorted(CLEAN_IMPORTED_JOBS.map((job) => job.operator));
const WORK_TYPE_META = buildWorkTypeMeta(CLEAN_IMPORTED_JOBS);

initialize();
function initialize() {
  loadState();
  renderDatalists();
  renderOperatorFilterOptions();
  wireEvents();
  resetForm();
  initializeAnalysisRange();
  renderAnalysisChips();
  setActiveView("operativita");
  render();
}

function wireEvents() {
  viewButtons.forEach((button) => button.addEventListener("click", () => setActiveView(button.dataset.view || "operativita")));
  jobForm.addEventListener("submit", handleSubmit);
  cancelEditButton.addEventListener("click", resetForm);
  exportCsvButton.addEventListener("click", exportCsv);
  exportJsonButton.addEventListener("click", exportJson);
  resetArchiveButton.addEventListener("click", resetArchive);
  applyFiltersButton.addEventListener("click", render);
  resetFiltersButton.addEventListener("click", resetFilters);
  jobSearchInput.addEventListener("change", render);
  jobClientFilterInput.addEventListener("input", render);
  jobOperatorFilterInput.addEventListener("change", render);
  jobUrgencyFilter.addEventListener("change", render);
  tableSearchInput.addEventListener("input", render);
  analysisStartInput.addEventListener("change", render);
  analysisEndInput.addEventListener("change", render);
  analysisDetailSearchInput.addEventListener("input", renderAnalysisDetailList);
  [jobListPriceInput, jobAppliedPriceInput, jobQuantityInput].forEach((input) => input.addEventListener("input", updatePricePreview));
  jobTypeInput.addEventListener("change", handleWorkTypeSelection);
  jobTypeInput.addEventListener("blur", handleWorkTypeSelection);
}

function loadState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      state.jobs = Array.isArray(parsed.jobs) && parsed.jobs.length ? parsed.jobs.map(normalizeJobRecord).filter(isUsefulJobRecord) : [...IMPORTED_JOBS].map(normalizeJobRecord).filter(isUsefulJobRecord);
      return;
    }
  } catch (error) {
    console.error(error);
  }
  state.jobs = [...IMPORTED_JOBS].map(normalizeJobRecord).filter(isUsefulJobRecord);
  persistState();
}

function persistState() { window.localStorage.setItem(STORAGE_KEY, JSON.stringify({ jobs: state.jobs })); }
function initializeAnalysisRange() {
  const dates = state.jobs.flatMap((job) => [job.entryDate, job.deliveryDate, job.trialOutDate, job.trialReturnDate]).filter(Boolean).sort();
  analysisStartInput.value = dates[0] || "2026-02-01";
  analysisEndInput.value = dates[dates.length - 1] || "2026-04-30";
}
function setActiveView(viewName) {
  state.activeView = viewName;
  viewButtons.forEach((button) => button.classList.toggle("active", button.dataset.view === viewName));
  Object.entries(viewPanels).forEach(([name, panel]) => panel.classList.toggle("active-view", name === viewName));
}
function handleSubmit(event) {
  event.preventDefault();
  const sanitizedCode = clean(jobCodeInput.value).replace(/\D/g, "").slice(0, 8);
  jobCodeInput.value = sanitizedCode;
  const quantity = Math.max(parseInt(jobQuantityInput.value || "1", 10) || 1, 1);
  const listPrice = parseNumber(jobListPriceInput.value);
  const appliedUnitPrice = jobAppliedPriceInput.value ? parseNumber(jobAppliedPriceInput.value) : listPrice;
  const job = {
    id: jobIdHidden.value || createId(),
    jobCode: sanitizedCode, client: clean(jobClientInput.value), type: clean(jobTypeInput.value), operator: clean(jobOperatorInput.value),
    listPrice, appliedPrice: appliedUnitPrice, quantity, standardDays: parseInt(jobStandardDaysInput.value || "0", 10) || 0,
    entryDate: jobEntryDateInput.value, deliveryDate: jobDeliveryDateInput.value, trialOutDate: jobTrialOutDateInput.value, trialReturnDate: jobTrialReturnDateInput.value,
    urgent: jobUrgentInput.checked, redo: jobRedoInput.checked, status: clean(jobStatusInput.value), note: clean(jobNoteInput.value), priceInfo: clean(jobPriceInfoInput.value), pieces: quantity
  };
  if (!job.jobCode || !job.client || !job.type || !job.entryDate) {
    window.alert("Completa almeno ID lavoro, cliente, lavorazione e data ingresso.");
    return;
  }
  if (!/^\d{1,8}$/.test(job.jobCode)) {
    window.alert("L'ID lavoro deve contenere solo numeri e massimo 8 cifre.");
    return;
  }
  const index = state.jobs.findIndex((item) => item.id === job.id);
  if (index >= 0) state.jobs.splice(index, 1, job); else state.jobs.unshift(job);
  persistState(); resetForm(); render();
}
function resetForm() {
  jobForm.reset(); jobIdHidden.value = ""; jobQuantityInput.value = "1"; jobStandardDaysInput.value = "4"; jobStatusInput.value = "In lavorazione"; jobEntryDateInput.value = isoToday(); updatePricePreview();
}
function resetFilters() { jobSearchInput.value = ""; jobClientFilterInput.value = ""; jobOperatorFilterInput.value = ""; jobUrgencyFilter.value = ""; tableSearchInput.value = ""; render(); }
function resetArchive() {
  if (!window.confirm("Vuoi davvero ripristinare l'archivio base dei lavori?")) return;
  state.jobs = [...IMPORTED_JOBS]; persistState(); resetForm(); initializeAnalysisRange(); render();
}
function exportJson() { downloadFile("inserimento-lavori-senza-excel.json", JSON.stringify(state.jobs, null, 2), "application/json;charset=utf-8;"); }
function exportCsv() {
  const rows = [["id_lavoro","cliente","lavorazione","operatore","listino_unitario","prezzo_applicato_unitario","quantita","totale","stato","ingresso","consegna","uscita_prova","rientro_prova","urgente","rifacimento","info_listino","note"], ...state.jobs.map((job) => [job.jobCode, job.client, job.type, job.operator, job.listPrice, job.appliedPrice, job.quantity, computeAppliedTotal(job), job.status, job.entryDate, job.deliveryDate, job.trialOutDate, job.trialReturnDate, job.urgent ? "Si" : "No", job.redo ? "Si" : "No", job.priceInfo, job.note])];
  const csv = rows.map((row) => row.map(escapeCsv).join(";")).join("\n");
  downloadFile("inserimento-lavori-senza-excel.csv", csv, "text/csv;charset=utf-8;");
}
function updatePricePreview() {
  const quantity = Math.max(parseInt(jobQuantityInput.value || "1", 10) || 1, 1);
  const listPrice = parseNumber(jobListPriceInput.value);
  const applied = jobAppliedPriceInput.value ? parseNumber(jobAppliedPriceInput.value) : listPrice;
  const listTotal = listPrice * quantity;
  const appliedTotal = applied * quantity;
  jobFinalPricePreview.value = appliedTotal ? formatCurrency(appliedTotal) : "Calcolato automaticamente";
  jobListTotalValue.textContent = formatCurrency(listTotal);
  jobAppliedTotalValue.textContent = `EUR ${formatDecimal(appliedTotal)}`;
}
function handleWorkTypeSelection() {
  const key = normalizeText(jobTypeInput.value);
  if (!key || !WORK_TYPE_META.has(key)) return;
  const meta = WORK_TYPE_META.get(key);
  if (!jobListPriceInput.value || parseNumber(jobListPriceInput.value) === 0) {
    jobListPriceInput.value = meta.listPrice ? String(meta.listPrice).replace(".", ",") : "";
  }
  if (!jobStandardDaysInput.value || Number(jobStandardDaysInput.value) === 0) {
    jobStandardDaysInput.value = meta.standardDays ? String(meta.standardDays) : "4";
  }
  if (!jobPriceInfoInput.value && meta.priceInfo) {
    jobPriceInfoInput.value = meta.priceInfo;
  }
  updatePricePreview();
}
function render() { const rows = getFilteredJobs(); renderFilterStatus(rows); renderRecentEntries(); renderTable(rows); renderAnalysis(); }
function getFilteredJobs() {
  const quickSearch = normalizeText(jobSearchInput.value);
  const clientFilter = normalizeText(jobClientFilterInput.value);
  const operatorFilter = normalizeText(jobOperatorFilterInput.value);
  const urgencyFilter = jobUrgencyFilter.value;
  const tableSearch = normalizeText(tableSearchInput.value);
  return state.jobs.filter((job) => {
    const haystack = normalizeText(`${job.jobCode} ${job.client} ${job.type} ${job.status} ${job.operator} ${job.note}`);
    const matchesQuick = matchesQuickFilter(job, quickSearch, haystack);
    const matchesClient = !clientFilter || normalizeText(job.client).includes(clientFilter);
    const matchesOperator = !operatorFilter || normalizeText(job.operator).includes(operatorFilter);
    const matchesUrgency = !urgencyFilter || (urgencyFilter === "urgent" ? job.urgent : !job.urgent);
    const matchesTable = !tableSearch || haystack.includes(tableSearch);
    return matchesQuick && matchesClient && matchesOperator && matchesUrgency && matchesTable;
  });
}
function renderFilterStatus(rows) {
  const parts = [];
  if (jobSearchInput.value) parts.push(`commesse "${jobSearchInput.options[jobSearchInput.selectedIndex].text}"`);
  if (jobClientFilterInput.value) parts.push(`cliente "${jobClientFilterInput.value}"`);
  if (jobOperatorFilterInput.value) parts.push(`operatore "${jobOperatorFilterInput.value}"`);
  if (jobUrgencyFilter.value === "urgent") parts.push("solo urgenti");
  if (jobUrgencyFilter.value === "normal") parts.push("no urgenti");
  filterStatus.textContent = parts.length ? `${rows.length} lavori con filtri attivi: ${parts.join(" | ")}` : "Nessun filtro attivo.";
}
function renderRecentEntries() {
  const latest = [...state.jobs].sort((a, b) => compareDate(b.entryDate, a.entryDate)).slice(0, 5);
  recentEntriesList.innerHTML = latest.map((job) => `<li>${escapeHtml(job.client)} - ${escapeHtml(job.type)} - ${escapeHtml(job.status)}</li>`).join("");
}
function renderTable(rows) {
  const sorted = [...rows].sort((a, b) => compareDate(b.entryDate, a.entryDate));
  jobTableBody.innerHTML = sorted.length ? sorted.map((job) => `
    <tr>
      <td data-label="ID">${escapeHtml(job.jobCode)}</td>
      <td data-label="Cliente">${escapeHtml(job.client)}</td>
      <td data-label="Lavorazione">${escapeHtml(job.type)}</td>
      <td data-label="Stato"><span class="status-pill ${getStatusTone(job.status)}">${escapeHtml(job.status)}</span></td>
      <td data-label="Ingresso">${escapeHtml(formatDate(job.entryDate))}</td>
      <td data-label="Consegna">${escapeHtml(formatDate(job.deliveryDate))}</td>
      <td data-label="Urgente"><span class="flag-pill ${job.urgent ? "flag-yes" : "flag-no"}">${job.urgent ? "Si" : "No"}</span></td>
      <td data-label="Rifacimento"><span class="flag-pill ${job.redo ? "flag-yes" : "flag-no"}">${job.redo ? "Si" : "No"}</span></td>
      <td data-label="Totale">${escapeHtml(formatCurrency(computeAppliedTotal(job)))}</td>
      <td data-label="Azioni"><div class="table-actions"><button class="table-action secondary" type="button" data-action="edit" data-id="${escapeHtml(job.id)}">Modifica</button><button class="table-action danger" type="button" data-action="delete" data-id="${escapeHtml(job.id)}">Elimina</button></div></td>
    </tr>`).join("") : `<tr><td colspan="10" class="empty-state">Nessun lavoro corrisponde ai filtri selezionati.</td></tr>`;
  jobTableBody.querySelectorAll("[data-action]").forEach((button) => button.addEventListener("click", handleTableAction));
}

function handleTableAction(event) {
  const action = event.currentTarget.dataset.action;
  const id = event.currentTarget.dataset.id;
  const job = state.jobs.find((item) => item.id === id);
  if (!job) return;
  if (action === "edit") {
    jobIdHidden.value = job.id; jobCodeInput.value = job.jobCode || ""; jobClientInput.value = job.client || ""; jobTypeInput.value = job.type || ""; jobOperatorInput.value = job.operator || ""; jobListPriceInput.value = job.listPrice || ""; jobStandardDaysInput.value = job.standardDays || ""; jobQuantityInput.value = job.quantity || 1; jobAppliedPriceInput.value = job.appliedPrice || ""; jobStatusInput.value = job.status || "In lavorazione"; jobEntryDateInput.value = job.entryDate || ""; jobDeliveryDateInput.value = job.deliveryDate || ""; jobTrialOutDateInput.value = job.trialOutDate || ""; jobTrialReturnDateInput.value = job.trialReturnDate || ""; jobUrgentInput.checked = !!job.urgent; jobRedoInput.checked = !!job.redo; jobPriceInfoInput.value = job.priceInfo || ""; jobNoteInput.value = job.note || ""; updatePricePreview(); setActiveView("operativita"); window.scrollTo({ top: 0, behavior: "smooth" }); return;
  }
  if (action === "delete") {
    if (!window.confirm(`Eliminare il lavoro ${job.jobCode}?`)) return;
    state.jobs = state.jobs.filter((item) => item.id !== id); persistState(); render();
  }
}

function renderAnalysisChips() {
  analysisChipBar.innerHTML = ANALYSIS_MODES.map((mode) => `<button class="chip-button ${mode === state.activeAnalysis ? "active" : ""}" data-analysis="${escapeHtml(mode)}" type="button">${escapeHtml(mode)}</button>`).join("");
  analysisChipBar.querySelectorAll("[data-analysis]").forEach((button) => button.addEventListener("click", () => { state.activeAnalysis = button.dataset.analysis; renderAnalysisChips(); renderAnalysis(); }));
}

function renderAnalysis() { const rows = getJobsInRange(); renderAnalysisKpis(rows); renderAnalysisMode(rows); }

function renderAnalysisKpis(rows) {
  const kpis = [
    { title: "Lavori nel periodo", value: rows.length },
    { title: "Pezzi entrati", value: rows.reduce((sum, job) => sum + (job.pieces || job.quantity || 1), 0) },
    { title: "Usciti in prova", value: rows.filter((job) => job.trialOutDate).length },
    { title: "Rientrati da prova", value: rows.filter((job) => job.trialReturnDate).length },
    { title: "Consegnati", value: rows.filter((job) => job.status === "Consegnato").length },
    { title: "Urgenze", value: rows.filter((job) => job.urgent).length },
    { title: "Rifacimenti", value: rows.filter((job) => job.redo).length },
    { title: "Valore totale", value: formatCurrency(rows.reduce((sum, job) => sum + computeAppliedTotal(job), 0)) }
  ];
  analysisKpiGrid.innerHTML = kpis.map((item) => `<article class="kpi-card"><span>${escapeHtml(item.title)}</span><strong>${escapeHtml(String(item.value))}</strong></article>`).join("");
}

function renderAnalysisMode(rows) {
  if (state.activeAnalysis === "Prova non rientrati") {
    const noReturn = rows.filter((job) => job.trialOutDate && !job.trialReturnDate);
    const grouped = summarizeGroupMetrics(noReturn, (job) => job.client, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Lavori usciti in prova e non rientrati",
      intro: `Ci sono ${noReturn.length} lavori fuori in prova senza rientro registrato nel periodo selezionato.`,
      summaryHeaders: ["Medici coinvolti", "Lavori", "Pezzi", "Valore stimato", "Da rientrare"],
      summaryValues: [String(grouped.length), String(noReturn.length), formatInteger(sumPieces(noReturn)), formatCurrency(sumApplied(noReturn)), String(noReturn.length)],
      tableHeaders: ["Medico", "Lavori", "Pezzi", "Valore", "Ultima uscita prova"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatDate(item.latestDate)], detailKey: item.label })),
      detailRecords: noReturn.map((job) => buildJobDetailRecord(job, job.client, [`Uscita prova ${formatDate(job.trialOutDate)}`, "Da rientrare"]))
    });
    return;
  } else if (state.activeAnalysis === "Consegnati nel periodo") {
    const delivered = rows.filter((job) => job.status === "Consegnato");
    const grouped = summarizeGroupMetrics(delivered, (job) => job.client, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Lavori consegnati nel periodo",
      intro: `Elenco dei lavori consegnati nel periodo, raggruppati per medico/studio con riepilogo economico e produttivo.`,
      summaryHeaders: ["Medici", "Lavori", "Pezzi", "Valore totale", "Ticket medio"],
      summaryValues: [String(grouped.length), String(delivered.length), formatInteger(sumPieces(delivered)), formatCurrency(sumApplied(delivered)), formatCurrency(delivered.length ? sumApplied(delivered) / delivered.length : 0)],
      tableHeaders: ["Medico", "Lavori", "Pezzi", "Valore", "Ultima consegna"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatDate(item.latestDate)], detailKey: item.label })),
      detailRecords: delivered.map((job) => buildJobDetailRecord(job, job.client, [`Consegnato il ${formatDate(job.deliveryDate)}`]))
    });
    return;
  } else if (state.activeAnalysis === "SLA per studio") {
    renderSlaAnalysis();
    return;
  } else if (state.activeAnalysis === "Distribuzione per medico") {
    const grouped = summarizeGroupMetrics(rows, (job) => job.client, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Distribuzione per medico",
      intro: "Distribuzione del volume lavori per medico con lettura congiunta di lavori, pezzi e valore.",
      summaryHeaders: ["Medici", "Lavori", "Pezzi", "Valore totale", "Media per medico"],
      summaryValues: [String(grouped.length), String(rows.length), formatInteger(sumPieces(rows)), formatCurrency(sumApplied(rows)), formatCurrency(grouped.length ? sumApplied(rows) / grouped.length : 0)],
      tableHeaders: ["Medico", "Lavori", "Pezzi", "Valore", "Valore medio/lavoro"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatCurrency(item.jobs ? item.value / item.jobs : 0)], detailKey: item.label })),
      detailRecords: rows.map((job) => buildJobDetailRecord(job, job.client, [`Ingresso ${formatDate(job.entryDate)}`, `Stato ${job.status || "-"}`]))
    });
    return;
  } else if (state.activeAnalysis === "Distribuzione per lavorazione") {
    const grouped = summarizeGroupMetrics(rows, (job) => job.type, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Distribuzione per lavorazione",
      intro: "Le lavorazioni con maggiore ricorrenza nel periodo selezionato, lette in chiave operativa ed economica.",
      summaryHeaders: ["Lavorazioni", "Lavori", "Pezzi", "Valore totale", "Media per lavorazione"],
      summaryValues: [String(grouped.length), String(rows.length), formatInteger(sumPieces(rows)), formatCurrency(sumApplied(rows)), formatCurrency(grouped.length ? sumApplied(rows) / grouped.length : 0)],
      tableHeaders: ["Lavorazione", "Lavori", "Pezzi", "Valore", "Valore medio/lavoro"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatCurrency(item.jobs ? item.value / item.jobs : 0)], detailKey: item.label })),
      detailRecords: rows.map((job) => buildJobDetailRecord(job, job.type, [`Operatore ${job.operator || "-"}`, `Ingresso ${formatDate(job.entryDate)}`]))
    });
    return;
  } else if (state.activeAnalysis === "Pareto clienti 80/20") {
    renderParetoAnalysis({
      title: "Analisi 80/20 dei medici",
      periodRows: rows,
      groupLabel: "Medici con valore",
      itemLabel: "Medico",
      noteTop: "Medico trainante",
      noteMid: "Presenza marginale",
      noteZero: "Nessun valore nel periodo",
      groupBy: (job) => clean(job.client) || "Medico non indicato",
      valueBy: (job) => computeAppliedTotal(job),
      piecesBy: (job) => Number(job.pieces || job.quantity || 1)
    });
    return;
  } else if (state.activeAnalysis === "Pareto lavorazioni 80/20") {
    renderParetoAnalysis({
      title: "Analisi 80/20 delle lavorazioni prodotte",
      periodRows: rows,
      groupLabel: "Lavorazioni con pezzi",
      itemLabel: "Lavorazione",
      noteTop: "Lavorazione trainante",
      noteMid: "Presenza marginale",
      noteZero: "Nessun pezzo nel periodo",
      groupBy: (job) => clean(job.type) || "Lavorazione non indicata",
      valueBy: (job) => Number(job.pieces || job.quantity || 1),
      piecesBy: (job) => Number(job.pieces || job.quantity || 1)
    });
    return;
  } else if (state.activeAnalysis === "Numeri economici") {
    const delivered = rows.filter((job) => normalizeText(job.status) === "consegnato");
    renderEconomicAnalysis(delivered);
    return;
  } else if (state.activeAnalysis === "Valore potenziale") {
    const openJobs = rows.filter((job) => job.status !== "Consegnato");
    const grouped = summarizeGroupMetrics(openJobs, (job) => job.client, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Valore potenziale aperto",
      intro: "Valore ancora aperto sui lavori non consegnati, utile per leggere il portafoglio in essere.",
      summaryHeaders: ["Medici aperti", "Lavori aperti", "Pezzi", "Valore aperto", "Media aperto/lavoro"],
      summaryValues: [String(grouped.length), String(openJobs.length), formatInteger(sumPieces(openJobs)), formatCurrency(sumApplied(openJobs)), formatCurrency(openJobs.length ? sumApplied(openJobs) / openJobs.length : 0)],
      tableHeaders: ["Medico", "Lavori", "Pezzi", "Valore aperto", "Ultimo ingresso"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatDate(item.latestDate)], detailKey: item.label })),
      detailRecords: openJobs.map((job) => buildJobDetailRecord(job, job.client, [`Stato ${job.status || "-"}`, `Ingresso ${formatDate(job.entryDate)}`]))
    });
    return;
  } else if (state.activeAnalysis === "Clienti piu' scontati") {
    const discountedRows = rows.filter((job) => computeListTotal(job) > computeAppliedTotal(job));
    const grouped = summarizeGroupMetrics(discountedRows, (job) => job.client, (job) => computeListTotal(job) - computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Clienti piu' scontati",
      intro: "Clienti con maggiore differenza tra listino teorico e prezzo effettivamente applicato.",
      summaryHeaders: ["Clienti scontati", "Lavori scontati", "Pezzi", "Sconto totale", "Sconto medio"],
      summaryValues: [String(grouped.length), String(discountedRows.length), formatInteger(sumPieces(discountedRows)), formatCurrency(grouped.reduce((sum, item) => sum + item.value, 0)), formatCurrency(discountedRows.length ? grouped.reduce((sum, item) => sum + item.value, 0) / discountedRows.length : 0)],
      tableHeaders: ["Medico", "Lavori", "Pezzi", "Sconto totale", "Sconto medio/lavoro"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatCurrency(item.jobs ? item.value / item.jobs : 0)], detailKey: item.label })),
      detailRecords: discountedRows.map((job) => buildJobDetailRecord(job, job.client, [`Sconto ${formatCurrency(computeListTotal(job) - computeAppliedTotal(job))}`, `Listino ${formatCurrency(computeListTotal(job))}`]))
    });
    return;
  } else if (state.activeAnalysis === "Lavorazioni per volume") {
    const grouped = summarizeGroupMetrics(rows, (job) => job.type, (job) => Number(job.pieces || job.quantity || 1));
    renderUniformAnalysis({
      title: "Lavorazioni per volume",
      intro: "Volume pezzi per tipologia di lavorazione, utile per leggere il carico produttivo reale.",
      summaryHeaders: ["Lavorazioni", "Lavori", "Pezzi", "Pezzi medi/lavoro", "Valore totale"],
      summaryValues: [String(grouped.length), String(rows.length), formatInteger(sumPieces(rows)), formatDecimal(rows.length ? sumPieces(rows) / rows.length : 0), formatCurrency(sumApplied(rows))],
      tableHeaders: ["Lavorazione", "Lavori", "Pezzi", "Pezzi medi/lavoro", "Valore"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatDecimal(item.jobs ? item.pieces / item.jobs : 0), formatCurrency(item.appliedValue)], detailKey: item.label })),
      detailRecords: rows.map((job) => buildJobDetailRecord(job, job.type, [`Operatore ${job.operator || "-"}`, `Pezzi ${formatInteger(job.pieces || job.quantity || 1)}`]))
    });
    return;
  } else if (state.activeAnalysis === "Margine perso per lavorazione") {
    const discountedRows = rows.filter((job) => computeListTotal(job) > computeAppliedTotal(job));
    const grouped = summarizeGroupMetrics(discountedRows, (job) => job.type, (job) => computeListTotal(job) - computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Margine perso per lavorazione",
      intro: "Differenza economica cumulata sulle lavorazioni scontate, per capire dove si concentra la perdita di margine.",
      summaryHeaders: ["Lavorazioni", "Lavori scontati", "Pezzi", "Margine perso", "Perdita media"],
      summaryValues: [String(grouped.length), String(discountedRows.length), formatInteger(sumPieces(discountedRows)), formatCurrency(grouped.reduce((sum, item) => sum + item.value, 0)), formatCurrency(discountedRows.length ? grouped.reduce((sum, item) => sum + item.value, 0) / discountedRows.length : 0)],
      tableHeaders: ["Lavorazione", "Lavori", "Pezzi", "Margine perso", "Perdita media/lavoro"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatCurrency(item.jobs ? item.value / item.jobs : 0)], detailKey: item.label })),
      detailRecords: discountedRows.map((job) => buildJobDetailRecord(job, job.type, [`Margine perso ${formatCurrency(computeListTotal(job) - computeAppliedTotal(job))}`, `Cliente ${job.client}`]))
    });
    return;
  } else {
    const metrics = [
      ["Tempo medio standard", `${formatDecimal(average(rows.map((job) => job.standardDays || 0)))} gg`, "Media dei giorni standard impostati sulle lavorazioni"],
      ["Lavori in prova", String(rows.filter((job) => normalizeText(job.status) === "in prova").length), "Lavori attualmente in stato In prova"],
      ["Consegne registrate", String(rows.filter((job) => job.deliveryDate).length), "Lavori con data consegna compilata"],
      ["Urgenze", String(rows.filter((job) => job.urgent).length), "Lavori marcati come urgenti"],
      ["Rifacimenti", String(rows.filter((job) => job.redo).length), "Lavori marcati come rifacimento"]
    ];
    renderUniformAnalysis({
      title: "Indicatori di flusso",
      intro: "Lettura sintetica dei principali indicatori di avanzamento e carico del laboratorio.",
      summaryHeaders: ["Lavori", "Pezzi", "Valore", "In prova", "Consegnati"],
      summaryValues: [String(rows.length), formatInteger(sumPieces(rows)), formatCurrency(sumApplied(rows)), String(rows.filter((job) => normalizeText(job.status) === "in prova").length), String(rows.filter((job) => job.deliveryDate).length)],
      tableHeaders: ["Indicatore", "Valore", "Nota"],
      tableRows: metrics,
      detailRecords: rows.map((job) => buildJobDetailRecord(job, job.status || "stato", [`Ingresso ${formatDate(job.entryDate)}`, `Stato ${job.status || "-"}`])),
      detailEmptyLabel: "Nessun dettaglio di flusso disponibile."
    });
    return;
  }
}
function renderUniformAnalysis({ title, intro, summaryHeaders, summaryValues, tableHeaders, tableRows, detailRecords, detailEmptyLabel }) {
  analysisTitle.textContent = title;
  analysisIntro.textContent = intro;
  analysisPrimary.innerHTML = `
    <section class="sla-report">
      <div class="sla-summary-block">
        <table class="sla-table sla-summary-table">
          <thead>
            <tr>${summaryHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
          </thead>
          <tbody>
            <tr>${summaryValues.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}</tr>
          </tbody>
        </table>
      </div>
      <div class="sla-summary-block">
        <table class="sla-table">
          <thead>
            <tr>${tableHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${tableRows.length ? tableRows.map((row) => renderAnalysisTableRow(row)).join("") : `<tr><td colspan="${tableHeaders.length}">Nessun dato disponibile nel periodo selezionato.</td></tr>`}
          </tbody>
        </table>
      </div>
    </section>
  `;
  bindAnalysisDetailRowClicks();
  setAnalysisDetailContext(detailRecords, detailEmptyLabel);
}
function renderAnalysisTableRow(row) {
  const normalizedRow = Array.isArray(row) ? { cells: row } : row;
  const detailKey = normalizedRow.detailKey ? normalizeText(normalizedRow.detailKey) : "";
  const detailLabel = normalizedRow.detailLabel || normalizedRow.detailKey || "";
  return `
    <tr${detailKey ? ` class="analysis-clickable-row" data-detail-key="${escapeHtml(detailKey)}" data-detail-label="${escapeHtml(detailLabel)}"` : ""}>
      ${normalizedRow.cells.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}
    </tr>
  `;
}
function setAnalysisDetailContext(records = [], emptyLabel = "Nessun dettaglio disponibile.") {
  state.analysisDetailRecords = Array.isArray(records) ? records : [];
  state.analysisDetailFilterKey = "";
  state.analysisDetailFilterLabel = "";
  state.analysisDetailEmptyLabel = emptyLabel;
  analysisDetailSearchInput.value = "";
  analysisDetailSearchInput.placeholder = "Cerca nel dettaglio o seleziona una riga a sinistra";
  renderAnalysisDetailList();
}
function bindAnalysisDetailRowClicks() {
  analysisPrimary.querySelectorAll("[data-detail-key]").forEach((row) => {
    row.addEventListener("click", () => {
      const nextKey = row.dataset.detailKey || "";
      const nextLabel = row.dataset.detailLabel || "";
      if (state.analysisDetailFilterKey === nextKey) {
        state.analysisDetailFilterKey = "";
        state.analysisDetailFilterLabel = "";
      } else {
        state.analysisDetailFilterKey = nextKey;
        state.analysisDetailFilterLabel = nextLabel;
      }
      updateAnalysisDetailSelection();
      renderAnalysisDetailList();
    });
  });
  updateAnalysisDetailSelection();
}
function updateAnalysisDetailSelection() {
  analysisPrimary.querySelectorAll("[data-detail-key]").forEach((row) => {
    row.classList.toggle("analysis-row-selected", row.dataset.detailKey === state.analysisDetailFilterKey && !!state.analysisDetailFilterKey);
  });
}
function renderAnalysisDetailList() {
  const searchValue = normalizeText(analysisDetailSearchInput.value);
  const selectedKey = normalizeText(state.analysisDetailFilterKey);
  const filtered = state.analysisDetailRecords.filter((record) => {
    const matchesSelected = !selectedKey || normalizeText(record.group) === selectedKey;
    const haystack = normalizeText(`${record.group} ${record.title} ${record.meta} ${record.searchText || ""}`);
    const matchesSearch = !searchValue || haystack.includes(searchValue);
    return matchesSelected && matchesSearch;
  });
  if (!filtered.length) {
    const message = state.analysisDetailFilterLabel
      ? `Nessun dettaglio trovato per ${state.analysisDetailFilterLabel.toLowerCase()}.`
      : state.analysisDetailEmptyLabel;
    analysisDetailList.innerHTML = `<li>${escapeHtml(message)}</li>`;
    return;
  }
  analysisDetailList.innerHTML = filtered.map((record) => `
    <li>
      <strong>${escapeHtml(record.title)}</strong>
      ${record.meta ? `<br><span class="detail-meta">${escapeHtml(record.meta)}</span>` : ""}
    </li>
  `).join("");
}
function buildJobDetailRecord(job, groupKey, metaParts = []) {
  const resolvedGroup = groupKey || job.client || job.type || "non indicato";
  const details = [
    job.client ? `Medico: ${job.client}` : "",
    job.type ? `Lavorazione: ${job.type}` : "",
    job.jobCode ? `ID ${job.jobCode}` : "",
    `Pezzi ${formatInteger(job.pieces || job.quantity || 1)}`,
    formatCurrency(computeAppliedTotal(job)),
    ...metaParts.filter(Boolean)
  ].filter(Boolean);
  return {
    group: normalizeText(resolvedGroup),
    title: `${job.client || "Medico non indicato"} - ${job.type || "Lavorazione non indicata"}`,
    meta: details.join(" | "),
    searchText: `${job.operator || ""} ${job.note || ""} ${job.status || ""}`
  };
}
function normalizeImportedText(value) {
  const text = clean(value);
  if (!text) return "";
  if (text === "System.Xml.XmlElement") return "";
  return text
    .replace(/[–—]/g, "-")
    .replace(/\uFFFD/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
function normalizeJobRecord(job) {
  return {
    ...job,
    jobCode: normalizeImportedText(job.jobCode),
    client: normalizeImportedText(job.client),
    type: normalizeImportedText(job.type),
    operator: normalizeImportedText(job.operator),
    note: normalizeImportedText(job.note),
    priceInfo: normalizeImportedText(job.priceInfo),
    status: normalizeImportedText(job.status) || "In lavorazione"
  };
}
function isUsefulJobRecord(job) {
  return !!(normalizeImportedText(job.client) || normalizeImportedText(job.type) || normalizeImportedText(job.jobCode));
}
function renderEconomicAnalysis(delivered) {
  const monthMap = new Map();
  delivered.forEach((job) => {
    const key = job.month || (job.entryDate ? job.entryDate.slice(0, 7) : "senza-mese");
    if (!monthMap.has(key)) {
      monthMap.set(key, { key, applied: 0, theoretical: 0, manual: 0, pieces: 0, jobs: 0 });
    }
    const bucket = monthMap.get(key);
    bucket.applied += computeAppliedTotal(job);
    bucket.theoretical += computeListTotal(job);
    bucket.manual += (parseNumber(job.appliedPrice) || 0) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1);
    bucket.pieces += Number(job.pieces || job.quantity || 1);
    bucket.jobs += 1;
  });
  const monthlyRows = Array.from(monthMap.values()).sort((a, b) => a.key.localeCompare(b.key));
  const totalTheoretical = monthlyRows.reduce((sum, item) => sum + item.theoretical, 0);
  const totalApplied = monthlyRows.reduce((sum, item) => sum + item.applied, 0);
  const totalManual = monthlyRows.reduce((sum, item) => sum + item.manual, 0);
  const totalPieces = monthlyRows.reduce((sum, item) => sum + item.pieces, 0);
  const totalDifference = totalTheoretical - totalApplied;
  const maxMonthlyValue = Math.max(...monthlyRows.map((item) => item.theoretical), totalTheoretical, 1);

  analysisTitle.textContent = "Economico mensile";
  analysisIntro.textContent = `Lettura economica dei lavori consegnati nel periodo selezionato, con confronto mese per mese tra listino teorico e valore applicato.`;
  analysisPrimary.innerHTML = `
    <section class="economic-report">
      <div class="economic-rows">
        ${[
          { label: "Totale listino", value: formatCurrency(totalTheoretical), width: (totalTheoretical / maxMonthlyValue) * 100, tone: "strong" },
          { label: "Totale applicato", value: formatCurrency(totalApplied), width: (totalApplied / maxMonthlyValue) * 100, tone: "medium" },
          { label: "Scostamento", value: formatCurrency(totalDifference), width: (Math.abs(totalDifference) / maxMonthlyValue) * 100, tone: "soft" },
          { label: "Pezzi complessivi", value: `${formatInteger(totalPieces)} pezzi`, width: (totalPieces / Math.max(totalPieces, 1)) * 100, tone: "neutral" }
        ].map((item) => `
          <article class="chart-row chart-row-economic economic-row">
            <div class="chart-meta">
              <strong>${escapeHtml(item.label)}</strong>
              <span>${escapeHtml(item.value)}</span>
            </div>
            <div class="chart-bar economic-bar ${escapeHtml(item.tone)}"><span style="width:${Math.max(Math.min(item.width, 100), 6)}%"></span></div>
          </article>
        `).join("")}
      </div>
      <div class="sla-summary-block">
        <table class="sla-table economic-month-table">
          <thead>
            <tr>
              <th>Mese</th>
              <th>Lavori</th>
              <th>Pezzi</th>
              <th>Applicato</th>
              <th>Listino</th>
              <th>Scostamento</th>
            </tr>
          </thead>
          <tbody>
            ${monthlyRows.length ? monthlyRows.map((item) => `
              <tr class="analysis-clickable-row" data-detail-key="${escapeHtml(item.key)}" data-detail-label="${escapeHtml(formatMonthLabel(item.key))}">
                <td>${escapeHtml(formatMonthLabel(item.key))}</td>
                <td>${escapeHtml(formatInteger(item.jobs))}</td>
                <td>${escapeHtml(formatInteger(item.pieces))}</td>
                <td>${escapeHtml(formatCurrency(item.applied))}</td>
                <td>${escapeHtml(formatCurrency(item.theoretical))}</td>
                <td>${escapeHtml(formatCurrency(item.theoretical - item.applied))}</td>
              </tr>
            `).join("") : `<tr><td colspan="6">Nessun dato economico disponibile nel periodo selezionato.</td></tr>`}
          </tbody>
        </table>
      </div>
    </section>
  `;
  bindAnalysisDetailRowClicks();
  setAnalysisDetailContext(
    monthlyRows.map((item) => ({
      group: normalizeText(item.key),
      title: formatMonthLabel(item.key),
      meta: `Lavori ${formatInteger(item.jobs)} | Pezzi ${formatInteger(item.pieces)} | Applicato ${formatCurrency(item.applied)} | Listino ${formatCurrency(item.theoretical)} | Scostamento ${formatCurrency(item.theoretical - item.applied)}`,
      searchText: `${formatMonthLabel(item.key)} ${item.key}`
    })).concat(delivered.map((job) => {
      const monthKey = job.month || (job.entryDate ? job.entryDate.slice(0, 7) : "senza-mese");
      return buildJobDetailRecord(job, monthKey, [`Mese ${formatMonthLabel(monthKey)}`, `Listino ${formatCurrency(computeListTotal(job))}`, `Applicato ${formatCurrency(computeAppliedTotal(job))}`]);
    })),
    "Seleziona un mese a sinistra per vedere il dettaglio economico."
  );
  analysisDetailSearchInput.placeholder = "Cerca un mese o seleziona una riga a sinistra";
  if (!state.analysisDetailFilterKey) {
    analysisDetailList.innerHTML = `
      <li><strong>Lettura consigliata</strong><br><span class="detail-meta">Seleziona una riga del mese per vedere a destra il dettaglio dei lavori collegati, invece di una lista completa poco leggibile.</span></li>
      <li><strong>Totale periodo</strong><br><span class="detail-meta">Applicato ${escapeHtml(formatCurrency(totalApplied))} | Listino ${escapeHtml(formatCurrency(totalTheoretical))} | Scostamento ${escapeHtml(formatCurrency(totalDifference))} | Pezzi ${escapeHtml(formatInteger(totalPieces))}</span></li>
    `;
  }
}
function renderParetoAnalysis({ title, periodRows, groupLabel, itemLabel, noteTop, noteMid, noteZero, groupBy, valueBy, piecesBy }) {
  const groupedRows = summarizePareto(periodRows, groupBy, valueBy);
  const groupedPieces = summarizePareto(periodRows, groupBy, piecesBy || (() => 0));
  const piecesMap = new Map(groupedPieces.map((row) => [row.label, row.value]));
  const activeRows = groupedRows.filter((row) => row.value > 0);
  const totalValue = groupedRows.reduce((sum, row) => sum + row.value, 0);
  const topTwoValue = activeRows.slice(0, 2).reduce((sum, row) => sum + row.value, 0);
  const topFiveValue = activeRows.slice(0, 5).reduce((sum, row) => sum + row.value, 0);
  const incidenceTopTwo = totalValue > 0 ? (topTwoValue / totalValue) * 100 : 0;
  const incidenceTopFive = totalValue > 0 ? (topFiveValue / totalValue) * 100 : 0;
  const totalPieces = Array.from(piecesMap.values()).reduce((sum, value) => sum + value, 0);
  const topTwoPieces = activeRows.slice(0, 2).reduce((sum, row) => sum + (piecesMap.get(row.label) || 0), 0);

  analysisTitle.textContent = title;
  analysisIntro.textContent = `Periodo analizzato: ${formatDate(analysisStartInput.value)} - ${formatDate(analysisEndInput.value)}`;

  analysisPrimary.innerHTML = `
    <section class="sla-report">
      <div class="sla-summary-block">
        <table class="sla-table sla-summary-table pareto-summary-table">
          <thead>
            <tr>
              <th>${itemLabel === "Lavorazione" ? "Pezzi totali" : "Valore totale"}</th>
              <th>${itemLabel === "Lavorazione" ? "Lavorazioni a listino" : "Medici totali"}</th>
              <th>${groupLabel}</th>
              <th>Top 2 ${itemLabel.toLowerCase()}${itemLabel === "Lavorazione" ? "" : "i"}</th>
              <th>Incidenza Top 2</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>${escapeHtml(formatParetoMetric(totalValue, itemLabel))}</td>
              <td>${escapeHtml(String(groupedRows.length))}</td>
              <td>${escapeHtml(String(activeRows.length))}</td>
              <td>${escapeHtml(formatParetoMetric(topTwoValue, itemLabel))}</td>
              <td>${escapeHtml(`${formatDecimal(incidenceTopTwo)}%`)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="sla-summary-block">
        <table class="sla-table">
          <thead>
            <tr>
              <th>${itemLabel}</th>
              <th>${itemLabel === "Lavorazione" ? "Pezzi" : "Valore"}</th>
              <th>Pezzi</th>
              <th>% su totale</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody>
            ${groupedRows.map((row, index) => `
              <tr class="${row.value > 0 && index < 2 ? "pareto-row-top" : row.value > 0 ? "pareto-row-mid" : "pareto-row-zero"} analysis-clickable-row" data-detail-key="${escapeHtml(normalizeText(row.label))}" data-detail-label="${escapeHtml(row.label)}">
                <td>${escapeHtml(row.label)}</td>
                <td>${escapeHtml(formatParetoMetric(row.value, itemLabel))}</td>
                <td>${escapeHtml(formatInteger(piecesMap.get(row.label) || 0))}</td>
                <td>${escapeHtml(`${formatDecimal(row.share)}%`)}</td>
                <td>${escapeHtml(row.value === 0 ? noteZero : index < 2 ? noteTop : noteMid)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
  bindAnalysisDetailRowClicks();
  setAnalysisDetailContext(
    periodRows.map((job) => buildJobDetailRecord(job, groupBy(job), [`Pezzi ${formatInteger(piecesBy(job))}`, `Ingresso ${formatDate(job.entryDate)}`])),
    "Nessun dettaglio Pareto disponibile."
  );
}
function renderSlaAnalysis() {
  const rows = getSlaRowsInRange();
  const grouped = buildSlaByClient(rows);
  const deliveredRows = rows.filter((job) => !!job.deliveryDate);
  const validRows = rows.filter((job) => getDeliveryLeadDays(job) !== null);
  const flaggedRows = rows.filter((job) => needsSlaReview(job));
  const totalPieces = rows.reduce((sum, job) => sum + Number(job.pieces || job.quantity || 1), 0);
  const avgOverall = validRows.length ? validRows.reduce((sum, job) => sum + getDeliveryLeadDays(job), 0) / validRows.length : 0;

  analysisTitle.textContent = "SLA consegne per studio";
  analysisIntro.textContent = "Raggruppati per studio con periodo letto dalla data consegna e tempi di lavorazione da Data ingresso lavoro a Data consegna.";

  analysisPrimary.innerHTML = `
    <section class="sla-report">
      <div class="sla-summary-block">
        <table class="sla-table sla-summary-table">
          <thead>
            <tr>
              <th>Studi</th>
              <th>Lavori</th>
              <th>Pezzi</th>
              <th>Tempo medio complessivo</th>
              <th>Righe da verificare</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>${escapeHtml(String(grouped.length))}</td>
              <td>${escapeHtml(String(deliveredRows.length))}</td>
              <td>${escapeHtml(String(totalPieces))}</td>
              <td>${escapeHtml(`${formatDecimal(avgOverall)} giorni`)}</td>
              <td>${escapeHtml(String(flaggedRows.length))}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="sla-studio-list">
        ${grouped.map((group) => {
          const studioRows = group.rows.map((job) => {
            const leadDays = getDeliveryLeadDays(job);
            return `
              <tr class="analysis-clickable-row" data-detail-key="${escapeHtml(normalizeText(group.client))}" data-detail-label="${escapeHtml(group.client)}">
                <td>${escapeHtml(job.jobCode || "-")}</td>
                <td>${escapeHtml(job.type || "-")}</td>
                <td>${escapeHtml(String(job.quantity || 1))}</td>
                <td>${escapeHtml(formatDate(job.entryDate))}</td>
                <td>${escapeHtml(formatDate(job.deliveryDate))}</td>
                <td>${leadDays === null ? "Da verificare" : `${escapeHtml(formatDecimal(leadDays))} gg`}</td>
              </tr>
            `;
          }).join("");
          return `
            <section class="sla-studio-card">
              <div class="sla-studio-head">
                <h3>${escapeHtml(group.client)}</h3>
                <p>Lavori: ${escapeHtml(String(group.jobs))} | Pezzi: ${escapeHtml(String(group.pieces))} | Tempo medio: ${escapeHtml(`${formatDecimal(group.avgDays)} giorni`)}</p>
              </div>
              <table class="sla-table sla-studio-table">
                <thead>
                  <tr>
                    <th>ID lavoro</th>
                    <th>Articolo</th>
                    <th>Q.ta'</th>
                    <th>Ingresso</th>
                    <th>Consegna</th>
                    <th>Giorni lavorazione</th>
                  </tr>
                </thead>
                <tbody>${studioRows}</tbody>
                <tfoot>
                  <tr>
                    <td colspan="2">Totale studio</td>
                    <td>${escapeHtml(String(group.pieces))}</td>
                    <td></td>
                    <td></td>
                    <td>${escapeHtml(`${formatDecimal(group.avgDays)} giorni`)}</td>
                  </tr>
                </tfoot>
              </table>
            </section>
          `;
        }).join("")}
      </div>
    </section>
  `;
  bindAnalysisDetailRowClicks();
  setAnalysisDetailContext(
    rows.map((job) => buildJobDetailRecord(job, job.client, [`Ingresso ${formatDate(job.entryDate)}`, `Consegna ${formatDate(job.deliveryDate)}`, getDeliveryLeadDays(job) === null ? "Da verificare" : `${formatDecimal(getDeliveryLeadDays(job))} giorni`])),
    flaggedRows.length ? "Nessuna riga da verificare nel periodo selezionato." : "Nessun dettaglio SLA disponibile."
  );
}
function getJobsInRange() { return state.jobs.filter((job) => isBetween(job.entryDate, analysisStartInput.value, analysisEndInput.value)); }
function getSlaRowsInRange() {
  return state.jobs.filter((job) => job.deliveryDate && isBetween(job.deliveryDate, analysisStartInput.value, analysisEndInput.value));
}
function summarizePareto(rows, groupBy, valueBy) {
  const map = new Map();
  rows.forEach((job) => {
    const key = groupBy(job);
    const value = Number(valueBy(job) || 0);
    map.set(key, (map.get(key) || 0) + value);
  });
  const total = Array.from(map.values()).reduce((sum, value) => sum + value, 0);
  return Array.from(map.entries())
    .map(([label, value]) => ({
      label,
      value,
      share: total > 0 ? (value / total) * 100 : 0
    }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
}
function summarizeGroupMetrics(rows, groupBy, valueBy) {
  const map = new Map();
  rows.forEach((job) => {
    const key = groupBy(job) || "Non indicato";
    if (!map.has(key)) {
      map.set(key, { label: key, jobs: 0, pieces: 0, value: 0, appliedValue: 0, latestDate: "" });
    }
    const bucket = map.get(key);
    bucket.jobs += 1;
    bucket.pieces += Number(job.pieces || job.quantity || 1);
    bucket.value += Number(valueBy(job) || 0);
    bucket.appliedValue += computeAppliedTotal(job);
    const candidateDate = job.deliveryDate || job.trialOutDate || job.entryDate || "";
    if (candidateDate && (!bucket.latestDate || candidateDate > bucket.latestDate)) {
      bucket.latestDate = candidateDate;
    }
  });
  return Array.from(map.values()).sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
}
function aggregateBy(items, keySelector, valueSelector) {
  const grouped = items.reduce((map, item) => { const key = keySelector(item); map.set(key, (map.get(key) || 0) + valueSelector(item)); return map; }, new Map());
  return Array.from(grouped.entries(), ([key, total]) => ({ key, total })).sort((a, b) => b.total - a.total).slice(0, 10);
}
function computeBarWidth(value, list) { const numbers = list.map((item) => toBarNumber(item.value)); const max = Math.max(...numbers, 1); return (toBarNumber(value) / max) * 100; }
function toBarNumber(value) { if (typeof value === "number") return value; const normalized = String(value).replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", "."); return Number(normalized) || 0; }
function computeListTotal(job) { return parseNumber(job.listPrice) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1); }
function computeAppliedTotal(job) { return parseNumber(job.appliedPrice) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1); }
function sumPieces(rows) { return rows.reduce((sum, job) => sum + Number(job.pieces || job.quantity || 1), 0); }
function sumApplied(rows) { return rows.reduce((sum, job) => sum + computeAppliedTotal(job), 0); }
function formatParetoMetric(value, itemLabel) {
  if (itemLabel === "Lavorazione") return `${formatInteger(value)} pezzi`;
  return formatCurrency(value);
}
function formatInteger(value) {
  const number = Math.round(Number(value) || 0);
  return number.toLocaleString("it-IT");
}
function getDeliveryLeadDays(job) {
  if (!job.entryDate || !job.deliveryDate) return null;
  const start = new Date(`${job.entryDate}T00:00:00`);
  const end = new Date(`${job.deliveryDate}T00:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null;
  const diff = Math.round((end.getTime() - start.getTime()) / 86400000);
  if (diff < 0) return null;
  return diff;
}
function needsSlaReview(job) {
  return !!job.entryDate && (!job.deliveryDate || getDeliveryLeadDays(job) === null);
}
function buildSlaByClient(rows) {
  const grouped = new Map();
  rows.forEach((job) => {
    if (!job.entryDate) return;
    const client = clean(job.client) || "Studio non indicato";
    if (!grouped.has(client)) {
      grouped.set(client, { client, jobs: 0, pieces: 0, totalDays: 0, validCount: 0, minDays: null, maxDays: null, reviewCount: 0, rows: [] });
    }
    const bucket = grouped.get(client);
    bucket.jobs += 1;
    bucket.pieces += Number(job.pieces || job.quantity || 1);
    bucket.rows.push(job);
    const leadDays = getDeliveryLeadDays(job);
    if (leadDays === null) {
      bucket.reviewCount += 1;
      return;
    }
    bucket.validCount += 1;
    bucket.totalDays += leadDays;
    bucket.minDays = bucket.minDays === null ? leadDays : Math.min(bucket.minDays, leadDays);
    bucket.maxDays = bucket.maxDays === null ? leadDays : Math.max(bucket.maxDays, leadDays);
  });
  return Array.from(grouped.values())
    .map((item) => ({
      ...item,
      avgDays: item.validCount ? item.totalDays / item.validCount : 0,
      minDays: item.minDays ?? 0,
      maxDays: item.maxDays ?? 0
    }))
    .sort((a, b) => a.client.localeCompare(b.client));
}
function getStatusTone(status) { return { "In lavorazione": "status-neutral", "In prova": "status-warn", Consegnato: "status-ok", Pronto: "status-neutral", Sospeso: "status-stop" }[status] || "status-neutral"; }
function isBetween(value, start, end) { if (!value) return false; if (start && value < start) return false; if (end && value > end) return false; return true; }
function compareDate(left, right) { return (left || "").localeCompare(right || ""); }
function average(values) { const filtered = values.filter((value) => Number.isFinite(value)); return filtered.length ? filtered.reduce((sum, value) => sum + value, 0) / filtered.length : 0; }
function createId() { if (window.crypto && typeof window.crypto.randomUUID === "function") return window.crypto.randomUUID(); return `job-${Date.now()}`; }
function downloadFile(filename, content, contentType) { const blob = new Blob([content], { type: contentType }); const url = window.URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = filename; anchor.click(); window.URL.revokeObjectURL(url); }
function formatDate(value) { if (!value) return "-"; const [year, month, day] = value.split("-"); return `${day}/${month}/${year}`; }
function formatMonthLabel(value) {
  if (!value || !/^\d{4}-\d{2}$/.test(value)) return value || "nessun dato";
  const [year, month] = value.split("-");
  const monthNames = ["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"];
  return `${monthNames[Number(month) - 1]} ${year}`;
}
function formatCurrency(value) { return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(parseNumber(value)); }
function formatDecimal(value) { return new Intl.NumberFormat("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(parseNumber(value)); }
function parseNumber(value) { if (typeof value === "number") return value; const normalized = String(value || "").replace(/\s/g, "").replace(/\./g, "").replace(",", "."); return Number(normalized) || 0; }
function normalizeText(value) { return String(value || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(); }
function clean(value) { return String(value || "").trim(); }
function isoToday() { const today = new Date(); const month = String(today.getMonth() + 1).padStart(2, "0"); const day = String(today.getDate()).padStart(2, "0"); return `${today.getFullYear()}-${month}-${day}`; }
function escapeHtml(value) { return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;"); }
function escapeCsv(value) { const text = String(value ?? ""); return /[;"\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text; }
function getUniqueSorted(values) { return [...new Set(values.map((value) => clean(value)).filter(Boolean))].sort((a, b) => a.localeCompare(b, "it")); }
function renderDatalists() {
  clientOptions.innerHTML = CLIENT_CHOICES.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  clientFilterOptions.innerHTML = CLIENT_CHOICES.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  workTypeOptions.innerHTML = WORK_TYPE_CHOICES.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  operatorOptions.innerHTML = OPERATOR_CHOICES.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
}
function renderOperatorFilterOptions() {
  jobOperatorFilterInput.innerHTML = `<option value="">Filtra per operatore</option>${OPERATOR_CHOICES.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}`;
}
function buildWorkTypeMeta(items) {
  const grouped = new Map();
  items.forEach((job) => {
    const key = normalizeText(job.type);
    if (!key) return;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(job);
  });
  const meta = new Map();
  grouped.forEach((jobs, key) => {
    const priceCounts = new Map();
    const dayCounts = new Map();
    jobs.forEach((job) => {
      const price = Number(job.listPrice || 0);
      const days = Number(job.standardDays || 0);
      if (price > 0) priceCounts.set(price, (priceCounts.get(price) || 0) + 1);
      if (days > 0) dayCounts.set(days, (dayCounts.get(days) || 0) + 1);
    });
    meta.set(key, {
      listPrice: getMostFrequentNumber(priceCounts),
      standardDays: getMostFrequentNumber(dayCounts),
      priceInfo: jobs.find((job) => clean(job.priceInfo))?.priceInfo || ""
    });
  });
  return meta;
}
function getMostFrequentNumber(counter) {
  let bestValue = 0;
  let bestCount = -1;
  counter.forEach((count, value) => {
    if (count > bestCount) {
      bestValue = Number(value);
      bestCount = count;
    }
  });
  return bestValue;
}
function matchesQuickFilter(job, filterKey, haystack) {
  if (!filterKey) return true;

  const normalizedStatus = normalizeText(job.status);
  const isDelivered = normalizedStatus === "consegnato";
  const isReady = normalizedStatus === "pronto";
  const isWorking = normalizedStatus === "in lavorazione";
  const isTrial = normalizedStatus === "in prova";
  const isSuspended = normalizedStatus === "sospeso";

  if (filterKey === "aperte") return isWorking || isTrial || isSuspended;
  if (filterKey === "in lavorazione") return isWorking;
  if (filterKey === "in prova") return isTrial;
  if (filterKey === "da rientrare") return !!job.trialOutDate && !job.trialReturnDate;
  if (filterKey === "consegnato" || filterKey === "consegnate") return isDelivered;
  if (filterKey === "pronto") return isReady;
  if (filterKey === "urgenti") return !!job.urgent;
  if (filterKey === "rifacimenti") return !!job.redo;

  return haystack.includes(filterKey);
}

jobCodeInput.addEventListener("input", () => {
  jobCodeInput.value = clean(jobCodeInput.value).replace(/\D/g, "").slice(0, 8);
});
