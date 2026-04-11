
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

const ANALYSIS_MODES = ["Prova non rientrati","Consegnati nel periodo","Distribuzione per medico","Distribuzione per lavorazione","Pareto clienti 80/20","Pareto lavorazioni 80/20","Numeri economici","Valore potenziale","Clienti piu' scontati","Lavorazioni per volume","Margine perso per lavorazione","Indicatori flusso"];
const state = { jobs: [], activeView: "operativita", activeAnalysis: ANALYSIS_MODES[0] };
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
const clientOptions = document.getElementById("clientOptions");
const workTypeOptions = document.getElementById("workTypeOptions");
const operatorOptions = document.getElementById("operatorOptions");

const CLIENT_CHOICES = getUniqueSorted(IMPORTED_JOBS.map((job) => job.client));
const WORK_TYPE_CHOICES = getUniqueSorted(IMPORTED_JOBS.map((job) => job.type));
const OPERATOR_CHOICES = getUniqueSorted(IMPORTED_JOBS.map((job) => job.operator));
const WORK_TYPE_META = buildWorkTypeMeta(IMPORTED_JOBS);

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
  tableSearchInput.addEventListener("input", render);
  analysisStartInput.addEventListener("change", render);
  analysisEndInput.addEventListener("change", render);
  [jobListPriceInput, jobAppliedPriceInput, jobQuantityInput].forEach((input) => input.addEventListener("input", updatePricePreview));
  jobTypeInput.addEventListener("change", handleWorkTypeSelection);
  jobTypeInput.addEventListener("blur", handleWorkTypeSelection);
}

function loadState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      state.jobs = Array.isArray(parsed.jobs) && parsed.jobs.length ? parsed.jobs : [...IMPORTED_JOBS];
      return;
    }
  } catch (error) {
    console.error(error);
  }
  state.jobs = [...IMPORTED_JOBS];
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
  let intro = "";
  let primaryItems = [];
  let detailItems = [];

  if (state.activeAnalysis === "Prova non rientrati") {
    const noReturn = rows.filter((job) => job.trialOutDate && !job.trialReturnDate);
    const grouped = aggregateBy(noReturn, (job) => job.client, () => 1);
    intro = `Ci sono ${noReturn.length} lavori fuori in prova senza rientro registrato.`;
    primaryItems = grouped.map((item) => ({ label: item.key, value: item.total }));
    detailItems = noReturn.map((job) => `${job.client} - ${job.type} (${formatDate(job.trialOutDate)})`);
  } else if (state.activeAnalysis === "Consegnati nel periodo") {
    const delivered = rows.filter((job) => job.status === "Consegnato");
    const grouped = aggregateBy(delivered, (job) => job.client, (job) => computeAppliedTotal(job));
    intro = `Nel periodo risultano ${delivered.length} lavori consegnati.`;
    primaryItems = grouped.map((item) => ({ label: item.key, value: formatCurrency(item.total) }));
    detailItems = delivered.map((job) => `${job.jobCode} - ${job.client} - ${job.type}`);
  } else if (state.activeAnalysis === "Distribuzione per medico") {
    const grouped = aggregateBy(rows, (job) => job.client, () => 1);
    intro = "Distribuzione del volume lavori per cliente o medico.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: item.total }));
    detailItems = grouped.map((item) => `${item.key}: ${item.total} lavori`);
  } else if (state.activeAnalysis === "Distribuzione per lavorazione") {
    const grouped = aggregateBy(rows, (job) => job.type, () => 1);
    intro = "Le lavorazioni con maggiore ricorrenza nel periodo selezionato.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: item.total }));
    detailItems = grouped.map((item) => `${item.key}: ${item.total} lavori`);
  } else if (state.activeAnalysis === "Pareto clienti 80/20") {
    const grouped = aggregateBy(rows, (job) => job.client, (job) => computeAppliedTotal(job));
    intro = "Clienti che generano la quota maggiore di valore.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: formatCurrency(item.total) }));
    detailItems = grouped.map((item) => `${item.key}: ${formatCurrency(item.total)}`);
  } else if (state.activeAnalysis === "Pareto lavorazioni 80/20") {
    const grouped = aggregateBy(rows, (job) => job.type, (job) => computeAppliedTotal(job));
    intro = "Lavorazioni che concentrano la maggior parte del valore.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: formatCurrency(item.total) }));
    detailItems = grouped.map((item) => `${item.key}: ${formatCurrency(item.total)}`);
  } else if (state.activeAnalysis === "Numeri economici") {
    const delivered = rows.filter((job) => normalizeText(job.status) === "consegnato");
    const monthMap = new Map();
    delivered.forEach((job) => {
      const key = job.month || (job.entryDate ? job.entryDate.slice(0, 7) : "senza-mese");
      if (!monthMap.has(key)) {
        monthMap.set(key, { key, applied: 0, theoretical: 0, manual: 0, pieces: 0 });
      }
      const bucket = monthMap.get(key);
      bucket.applied += computeAppliedTotal(job);
      bucket.theoretical += computeListTotal(job);
      bucket.manual += (parseNumber(job.appliedPrice) || 0) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1);
      bucket.pieces += Number(job.pieces || job.quantity || 1);
    });
    const monthlyRows = Array.from(monthMap.values()).sort((a, b) => a.key.localeCompare(b.key));
    const selectedMonth = monthlyRows[monthlyRows.length - 1] || { key: "nessun dato", applied: 0, theoretical: 0, manual: 0, pieces: 0 };
    const diff = selectedMonth.theoretical - selectedMonth.applied;
    const diffPercent = selectedMonth.theoretical > 0 ? (diff / selectedMonth.theoretical) * 100 : 0;
    intro = `Incassato mensile sui lavori consegnati nel periodo selezionato. Totale applicato ${formatCurrency(selectedMonth.applied)} su teorico da listino ${formatCurrency(selectedMonth.theoretical)}.`;
    primaryItems = monthlyRows.map((item) => ({
      label: formatMonthLabel(item.key),
      value: item.applied,
      valueText: formatCurrency(item.applied)
    }));
    detailItems = [
      `Totale teorico da listino: ${formatCurrency(selectedMonth.theoretical)}`,
      `Totale applicato: ${formatCurrency(selectedMonth.applied)}`,
      `Totale con prezzo manuale: ${formatCurrency(selectedMonth.manual)}`,
      `Differenza listino vs applicato: ${formatCurrency(diff)}`,
      `Scostamento medio sul valore teorico: ${formatDecimal(diffPercent)}%`,
      `Pezzi complessivi registrati: ${selectedMonth.pieces}`,
      `${formatMonthLabel(selectedMonth.key)} - incassato ${formatCurrency(selectedMonth.applied)} | listino ${formatCurrency(selectedMonth.theoretical)} | differenza ${formatCurrency(diff)} (${formatDecimal(diffPercent)}%)`
    ];
  } else if (state.activeAnalysis === "Valore potenziale") {
    const openJobs = rows.filter((job) => job.status !== "Consegnato");
    const grouped = aggregateBy(openJobs, (job) => job.client, (job) => computeAppliedTotal(job));
    intro = "Valore ancora aperto sui lavori non consegnati.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: formatCurrency(item.total) }));
    detailItems = openJobs.map((job) => `${job.client} - ${job.type} - ${formatCurrency(computeAppliedTotal(job))}`);
  } else if (state.activeAnalysis === "Clienti piu' scontati") {
    const discounted = rows.filter((job) => computeListTotal(job) > computeAppliedTotal(job)).map((job) => ({ label: job.client, total: computeListTotal(job) - computeAppliedTotal(job), detail: `${job.type} - sconto ${formatCurrency(computeListTotal(job) - computeAppliedTotal(job))}` })).sort((a, b) => b.total - a.total);
    intro = "Clienti con maggiore differenza tra listino e prezzo applicato.";
    primaryItems = discounted.map((item) => ({ label: item.label, value: formatCurrency(item.total) }));
    detailItems = discounted.map((item) => `${item.label}: ${item.detail}`);
  } else if (state.activeAnalysis === "Lavorazioni per volume") {
    const grouped = aggregateBy(rows, (job) => job.type, (job) => job.quantity || 1);
    intro = "Volume pezzi per tipologia di lavorazione.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: item.total }));
    detailItems = grouped.map((item) => `${item.key}: ${item.total} pezzi`);
  } else if (state.activeAnalysis === "Margine perso per lavorazione") {
    const grouped = aggregateBy(rows.filter((job) => computeListTotal(job) > computeAppliedTotal(job)), (job) => job.type, (job) => computeListTotal(job) - computeAppliedTotal(job));
    intro = "Differenza economica cumulata sulle lavorazioni scontate.";
    primaryItems = grouped.map((item) => ({ label: item.key, value: formatCurrency(item.total) }));
    detailItems = grouped.map((item) => `${item.key}: ${formatCurrency(item.total)}`);
  } else {
    intro = "Indicatori sintetici del flusso di laboratorio.";
    primaryItems = [{ label: "Tempo medio standard", value: `${formatDecimal(average(rows.map((job) => job.standardDays || 0)))} gg` }, { label: "Lavori in prova", value: rows.filter((job) => job.status === "In prova").length }, { label: "Consegne registrate", value: rows.filter((job) => job.deliveryDate).length }];
    detailItems = rows.slice(0, 8).map((job) => `${job.client} - ${job.status} - ingresso ${formatDate(job.entryDate)}`);
  }

  analysisTitle.textContent = state.activeAnalysis === "Numeri economici" ? "Economico mensile" : state.activeAnalysis;
  analysisIntro.textContent = intro;
  analysisPrimary.innerHTML = primaryItems.length ? primaryItems.map((item) => `<article class="chart-row ${state.activeAnalysis === "Numeri economici" ? "chart-row-economic" : ""}"><div class="chart-meta"><strong>${escapeHtml(String(item.label))}</strong><span>${escapeHtml(String(item.valueText || item.value))}</span></div><div class="chart-bar"><span style="width:${computeBarWidth(item.value, primaryItems)}%"></span></div></article>`).join("") : `<p>Nessun dato disponibile nel periodo selezionato.</p>`;
  analysisDetailList.innerHTML = detailItems.length ? detailItems.map((item) => `<li>${escapeHtml(item)}</li>`).join("") : `<li>Nessun dettaglio disponibile.</li>`;
}
function getJobsInRange() { return state.jobs.filter((job) => isBetween(job.entryDate, analysisStartInput.value, analysisEndInput.value)); }
function aggregateBy(items, keySelector, valueSelector) {
  const grouped = items.reduce((map, item) => { const key = keySelector(item); map.set(key, (map.get(key) || 0) + valueSelector(item)); return map; }, new Map());
  return Array.from(grouped.entries(), ([key, total]) => ({ key, total })).sort((a, b) => b.total - a.total).slice(0, 10);
}
function computeBarWidth(value, list) { const numbers = list.map((item) => toBarNumber(item.value)); const max = Math.max(...numbers, 1); return (toBarNumber(value) / max) * 100; }
function toBarNumber(value) { if (typeof value === "number") return value; const normalized = String(value).replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", "."); return Number(normalized) || 0; }
function computeListTotal(job) { return parseNumber(job.listPrice) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1); }
function computeAppliedTotal(job) { return parseNumber(job.appliedPrice) * Math.max(parseInt(job.quantity || 1, 10) || 1, 1); }
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
  if (filterKey === "aperte") return job.status !== "Consegnato";
  if (filterKey === "in lavorazione") return normalizeText(job.status) === "in lavorazione";
  if (filterKey === "in prova") return normalizeText(job.status) === "in prova";
  if (filterKey === "da rientrare") return !!job.trialOutDate && !job.trialReturnDate;
  if (filterKey === "consegnato") return normalizeText(job.status) === "consegnato";
  if (filterKey === "urgenti") return !!job.urgent;
  if (filterKey === "rifacimenti") return !!job.redo;
  return haystack.includes(filterKey);
}

jobCodeInput.addEventListener("input", () => {
  jobCodeInput.value = clean(jobCodeInput.value).replace(/\D/g, "").slice(0, 8);
});
