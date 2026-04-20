
const STORAGE_KEY = "inserimento_lavori_senza_excel_v4";
const GOOGLE_MAPS_API_KEY = document.querySelector('meta[name="google-maps-api-key"]')?.getAttribute("content")?.trim() || "";
const GOOGLE_MAPS_AUTOCOMPLETE_SCRIPT_ID = "google-maps-places-script";
const GOOGLE_MAPS_AUTOCOMPLETE_CALLBACK = "initGoogleMapsPlacesIntegration";

const DEFAULT_ADDRESS_BOOK = [
  { clientName: "AMBULATORIO ODONTOIATRICO DOTT. RUSSO ED ANNUNZIATA CARLO", address: "Via Lorenzo Cutugno 42, 98051 Barcellona Pozzo di Gotto (ME)", city: "Barcellona Pozzo di Gotto", province: "ME", verified: true, notes: "" },
  { clientName: "BARRESI ANTONIO", address: "Viale Regina Margherita 15/D, 98122 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "BARRESI PAOLO", address: "Via Adolfo Celi 139, 98125 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "BRIGUGLIO", address: "Via Errico Scipione 1, Villaggio Aldisio, 98100 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "CACICI RAFFAELE", address: "Via Ionica 17, 96010 Siracusa (SR)", city: "Siracusa", province: "SR", verified: true, notes: "" },
  { clientName: "CATARSINI ALFREDO", address: "Via Liberta 65, Brolo (ME)", city: "Brolo", province: "ME", verified: true, notes: "" },
  { clientName: "CENTRO ODONTOIATRICO OLIVA S.R.L. - S.T.P.", address: "Viale Aldo Moro 1, 89054 Galatro (RC)", city: "Galatro", province: "RC", verified: true, notes: "" },
  { clientName: "CORDARO RAFFAELE", address: "Viale della Liberta 393, 98121 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "D'ANDREA MASSIMO", address: "Via Giuseppe Garibaldi 126, 98122 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "D'ANGELO ALBERTO", address: "Via dei Marinai 82, 98049 Villafranca Tirrena (ME)", city: "Villafranca Tirrena", province: "ME", verified: true, notes: "" },
  { clientName: "DENTAL SRLS", address: "Strada Panoramica dello Stretto 2725, 98167 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "DI MARIA", address: "Via Giacomo Leopardi 23, Catania (CT)", city: "Catania", province: "CT", verified: true, notes: "" },
  { clientName: "FAMULARI GIUSEPPE", address: "Via Cesare Battisti 104, 98123 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "FIUMARA SALVATORE", address: "Via Industriale 166, Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "FRANZA TERESA GIACOMINA", address: "Via Centonze 38, 98122 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "G DENTAL SRL", address: "Via Tremonti 32, 98121 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "GDENTAL CLINIC SRLS", address: "Viale San Martino 315, 98124 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "GIORGIANNI", address: "Viale Europa 83, 98124 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "GRASSO SANDRO", address: "Via Cristoforo Colombo 10, Santa Lucia del Mela (ME)", city: "Santa Lucia del Mela", province: "ME", verified: true, notes: "" },
  { clientName: "LA SPADA CENTROMEDICO DENTISTICO SRL", address: "Via Caduti sul Lavoro 25, 98051 Barcellona Pozzo di Gotto (ME)", city: "Barcellona Pozzo di Gotto", province: "ME", verified: true, notes: "" },
  { clientName: "LATINO GIOVANNI", address: "Viale Giostra 98 Pal. C, 98100 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "LIFE S.R.L.", address: "Localita Santa Margherita 101, 98134 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "LO FORTE MASSIMO", address: "Via Maddalena 24, 98123 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "MINNITI GABRIELE", address: "Via dei Verdi 5, 98122 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "MIRACOLA LORENZO", address: "Via Nazionale 42, fraz. Scala, 98040 Torregrotta (ME)", city: "Torregrotta", province: "ME", verified: true, notes: "" },
  { clientName: "OLIVA FRANCESCO", address: "Via P. Togliatti, palazzina F s.n.c., Milazzo (ME)", city: "Milazzo", province: "ME", verified: true, notes: "" },
  { clientName: "PANZERA ROBERTO", address: "Piazza Matteotti 11, 98168 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "PETRELLI DIEGO", address: "Via Belvedere 16, Mirto (ME)", city: "Mirto", province: "ME", verified: true, notes: "" },
  { clientName: "ROTONDO FEDERICO", address: "Via Centonze 120, Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "SIRNA ANTONIO", address: "Via Calabria 44, 98076 Sant'Agata di Militello (ME)", city: "Sant'Agata di Militello", province: "ME", verified: true, notes: "" },
  { clientName: "SORACI ANNIBALE", address: "Via della Zecca 58/A, 98122 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "SPAGNOLO SABINA", address: "Via Tripoli 124, 98071 Capo d'Orlando (ME)", city: "Capo d'Orlando", province: "ME", verified: true, notes: "" },
  { clientName: "STAITI GIORGIO", address: "Traversa II Gorizia 21, 98165 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO ASSOCIATO ZAMPOGNA", address: "Via San Giovanni Bosco 14, Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO COPPIN SRL", address: "Via Giordano Bruno 120, 98123 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO DENT. DO.SSA SICLARI RITA E DR. TUZZA", address: "Viale Italia 101, 98124 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO DENT. VILLA DANTE DTT.RI METRO E ZA", address: "Via Catania, Residence Villa Dante, 98124 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO DENTISTICO ASSOCIATO K.S DENTAL", address: "Via Nino Bixio 89, 98123 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO DENTISTICO DOTT.SSA CAFIERO ANGELA", address: "Via Stagno 10, 98125 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "STUDIO M. D. DEI DOTT.RI ALBANESE S. E DI BE", address: "Via Santo Agostino 11, 98100 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "TEDESCO MICHELE", address: "Via Consolare Pompea 290, 98168 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" },
  { clientName: "VELO ALESSIA", address: "Via Nazionale 52, Valdina (ME)", city: "Valdina", province: "ME", verified: true, notes: "" },
  { clientName: "VENUTO MARCO", address: "Viale Regina Margherita 65, 98121 Messina (ME)", city: "Messina", province: "ME", verified: true, notes: "" }
];

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

const ANALYSIS_MODES = ["Prova non rientrati","Consegnati nel periodo","SLA per studio","Distribuzione per medico","Distribuzione per lavorazione","Pareto clienti 80/20","Pareto lavorazioni 80/20","Numeri economici","Valore potenziale","Clienti piu' scontati","Lavorazioni per volume","Margine perso per lavorazione","Saldo diretto"];
const DELIVERY_CITY_COORDINATES = {
  "barcellona pozzo di gotto|me": { lat: 38.1476, lon: 15.2142 },
  "brolo|me": { lat: 38.1567, lon: 14.8278 },
  "capo d'orlando|me": { lat: 38.1597, lon: 14.7343 },
  "capo d orlando|me": { lat: 38.1597, lon: 14.7343 },
  "catania|ct": { lat: 37.5079, lon: 15.0830 },
  "galatro|rc": { lat: 38.4595, lon: 15.1031 },
  "messina|me": { lat: 38.1938, lon: 15.5540 },
  "milazzo|me": { lat: 38.2207, lon: 15.2402 },
  "mirto|me": { lat: 38.0713, lon: 14.7517 },
  "santa lucia del mela|me": { lat: 38.1398, lon: 15.2832 },
  "sant'agata di militello|me": { lat: 38.0673, lon: 14.6362 },
  "sant agata di militello|me": { lat: 38.0673, lon: 14.6362 },
  "siracusa|sr": { lat: 37.0755, lon: 15.2866 },
  "torregrotta|me": { lat: 38.2043, lon: 15.3501 },
  "valdina|me": { lat: 38.1969, lon: 15.3687 },
  "villafranca tirrena|me": { lat: 38.2398, lon: 15.4384 }
};
const state = {
  jobs: [],
  addressBook: [],
  manualTasks: [],
  deliveryOrigin: "Laboratorio Dental G, Messina",
  deliverySelection: {},
  deliveryConfirmedJobIds: [],
  deliveryConfirmedAt: "",
  deliveryActionFilter: "all",
  activeView: "operativita",
  activeAnalysis: ANALYSIS_MODES[0],
  analysisDetailRecords: [],
  analysisDetailFilterKey: "",
  analysisDetailFilterLabel: "",
  analysisDetailEmptyLabel: "Nessun dettaglio disponibile."
};
let activeViewedJobId = "";
let isJobModalEditing = false;
let activeAddressClientName = "";
let activeAddressOriginalKey = "";
let googlePlacesLoadPromise = null;
const viewButtons = Array.from(document.querySelectorAll(".view-tab"));
const viewPanels = {
  operativita: document.getElementById("view-operativita"),
  analisi: document.getElementById("view-analisi"),
  consegne: document.getElementById("view-consegne")
};
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
const jobSaldoDirettoInput = document.getElementById("jobSaldoDirettoInput");
const jobPickupInput = document.getElementById("jobPickupInput");
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
const deliveryKpiGrid = document.getElementById("deliveryKpiGrid");
const deliveryOriginInput = document.getElementById("deliveryOriginInput");
const deliveryActionFilterInput = document.getElementById("deliveryActionFilterInput");
const openDeliveryRouteButton = document.getElementById("openDeliveryRouteButton");
const selectAllDeliveryButton = document.getElementById("selectAllDeliveryButton");
const clearSelectionDeliveryButton = document.getElementById("clearSelectionDeliveryButton");
const confirmDeliveryPlanButton = document.getElementById("confirmDeliveryPlanButton");
const printDeliveryPlanButton = document.getElementById("printDeliveryPlanButton");
const deliveryHint = document.getElementById("deliveryHint");
const deliveryRoutePreview = document.getElementById("deliveryRoutePreview");
const deliverySelectionSummary = document.getElementById("deliverySelectionSummary");
const deliveryConfirmedSummary = document.getElementById("deliveryConfirmedSummary");
const deliveryGroupedList = document.getElementById("deliveryGroupedList");
const deliveryAddressSearchInput = document.getElementById("deliveryAddressSearchInput");
const deliveryAddressBookList = document.getElementById("deliveryAddressBookList");
const manualTaskForm = document.getElementById("manualTaskForm");
const manualTaskIdHidden = document.getElementById("manualTaskIdHidden");
const manualTaskTitleInput = document.getElementById("manualTaskTitleInput");
const manualTaskClientInput = document.getElementById("manualTaskClientInput");
const manualTaskCreatedDateInput = document.getElementById("manualTaskCreatedDateInput");
const manualTaskAddressInput = document.getElementById("manualTaskAddressInput");
const manualTaskAddressHint = document.getElementById("manualTaskAddressHint");
const manualTaskCityInput = document.getElementById("manualTaskCityInput");
const manualTaskProvinceInput = document.getElementById("manualTaskProvinceInput");
const manualTaskKindInput = document.getElementById("manualTaskKindInput");
const manualTaskTimeSlotInput = document.getElementById("manualTaskTimeSlotInput");
const manualTaskUrgentInput = document.getElementById("manualTaskUrgentInput");
const manualTaskNoteInput = document.getElementById("manualTaskNoteInput");
const manualTaskCancelButton = document.getElementById("manualTaskCancelButton");
const manualTaskList = document.getElementById("manualTaskList");
const clientOptions = document.getElementById("clientOptions");
const clientFilterOptions = document.getElementById("clientFilterOptions");
const workTypeOptions = document.getElementById("workTypeOptions");
const operatorOptions = document.getElementById("operatorOptions");
const jobViewModal = document.getElementById("jobViewModal");
const jobViewBackdrop = document.getElementById("jobViewBackdrop");
const jobViewCloseButton = document.getElementById("jobViewCloseButton");
const jobViewCloseSecondary = document.getElementById("jobViewCloseSecondary");
const jobViewEditButton = document.getElementById("jobViewEditButton");
const jobViewTitle = document.getElementById("jobViewTitle");
const jobViewSubtitle = document.getElementById("jobViewSubtitle");
const jobViewContent = document.getElementById("jobViewContent");
const addressModal = document.getElementById("addressModal");
const addressModalBackdrop = document.getElementById("addressModalBackdrop");
const addressModalCloseButton = document.getElementById("addressModalCloseButton");
const addressModalCancelButton = document.getElementById("addressModalCancelButton");
const addressModalTitle = document.getElementById("addressModalTitle");
const addAddressRecordButton = document.getElementById("addAddressRecordButton");
const addressForm = document.getElementById("addressForm");
const addressClientInput = document.getElementById("addressClientInput");
const addressStreetInput = document.getElementById("addressStreetInput");
const addressCityInput = document.getElementById("addressCityInput");
const addressProvinceInput = document.getElementById("addressProvinceInput");
const addressNotesInput = document.getElementById("addressNotesInput");
const addressModalSaveButton = document.getElementById("addressModalSaveButton");
const manualTaskModal = document.getElementById("manualTaskModal");
const manualTaskModalBackdrop = document.getElementById("manualTaskModalBackdrop");
const manualTaskModalCloseButton = document.getElementById("manualTaskModalCloseButton");
const manualTaskModalCancelButton = document.getElementById("manualTaskModalCancelButton");
const manualTaskModalForm = document.getElementById("manualTaskModalForm");
const manualTaskModalIdHidden = document.getElementById("manualTaskModalIdHidden");
const manualTaskModalTitleInput = document.getElementById("manualTaskModalTitleInput");
const manualTaskModalClientInput = document.getElementById("manualTaskModalClientInput");
const manualTaskModalCreatedDateInput = document.getElementById("manualTaskModalCreatedDateInput");
const manualTaskModalAddressInput = document.getElementById("manualTaskModalAddressInput");
const manualTaskModalCityInput = document.getElementById("manualTaskModalCityInput");
const manualTaskModalProvinceInput = document.getElementById("manualTaskModalProvinceInput");
const manualTaskModalKindInput = document.getElementById("manualTaskModalKindInput");
const manualTaskModalTimeSlotInput = document.getElementById("manualTaskModalTimeSlotInput");
const manualTaskModalUrgentInput = document.getElementById("manualTaskModalUrgentInput");
const manualTaskModalNoteInput = document.getElementById("manualTaskModalNoteInput");

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
  resetManualTaskForm();
  initializeGoogleMapsAddressAutocomplete();
  initializeParentFrameReporting("lavori");
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
  deliveryOriginInput.addEventListener("input", handleDeliveryOriginChange);
  deliveryActionFilterInput.addEventListener("change", () => {
    state.deliveryActionFilter = normalizeDeliveryFilterValue(deliveryActionFilterInput.value || "all");
    persistState();
    renderDeliveryPlanner();
  });
  if (deliveryAddressSearchInput) {
    deliveryAddressSearchInput.addEventListener("input", renderDeliveryAddressBook);
  }
  if (deliveryAddressBookList) {
    deliveryAddressBookList.addEventListener("click", handleAddressBookClick);
  }
  deliveryGroupedList.addEventListener("change", handleDeliverySelectionChange);
  deliveryGroupedList.addEventListener("click", handleDeliveryListClick);
  if (manualTaskForm) {
    manualTaskForm.addEventListener("submit", handleManualTaskSubmit);
  }
  if (manualTaskCancelButton) {
    manualTaskCancelButton.addEventListener("click", resetManualTaskForm);
  }
  if (manualTaskList) {
    manualTaskList.addEventListener("click", handleManualTaskListClick);
  }
  if (manualTaskModalBackdrop) {
    manualTaskModalBackdrop.addEventListener("click", closeManualTaskModal);
  }
  if (manualTaskModalCloseButton) {
    manualTaskModalCloseButton.addEventListener("click", closeManualTaskModal);
  }
  if (manualTaskModalCancelButton) {
    manualTaskModalCancelButton.addEventListener("click", closeManualTaskModal);
  }
  if (manualTaskModalForm) {
    manualTaskModalForm.addEventListener("submit", handleManualTaskModalSubmit);
  }
  selectAllDeliveryButton.addEventListener("click", () => setAllDeliverySelection(true));
  clearSelectionDeliveryButton.addEventListener("click", () => setAllDeliverySelection(false));
  confirmDeliveryPlanButton.addEventListener("click", confirmDeliveryPlan);
  printDeliveryPlanButton.addEventListener("click", printDeliveryPlan);
  openDeliveryRouteButton.addEventListener("click", openDeliveryRouteInMaps);
  [jobListPriceInput, jobAppliedPriceInput, jobQuantityInput].forEach((input) => input.addEventListener("input", updatePricePreview));
  jobTypeInput.addEventListener("change", handleWorkTypeSelection);
  jobTypeInput.addEventListener("blur", handleWorkTypeSelection);
  jobViewBackdrop.addEventListener("click", closeJobModal);
  jobViewCloseButton.addEventListener("click", closeJobModal);
  jobViewCloseSecondary.addEventListener("click", handleModalSecondaryAction);
  jobViewEditButton.addEventListener("click", handleEditFromModal);
  addressModalBackdrop.addEventListener("click", closeAddressModal);
  addressModalCloseButton.addEventListener("click", closeAddressModal);
  addressModalCancelButton.addEventListener("click", closeAddressModal);
  addressForm.addEventListener("submit", handleAddressSave);
  if (addAddressRecordButton) {
    addAddressRecordButton.addEventListener("click", openNewAddressModal);
  }
  document.addEventListener("keydown", handleGlobalKeydown);
}

function loadState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      state.jobs = mergeSavedJobsWithImported(Array.isArray(parsed.jobs) ? parsed.jobs : []);
      state.addressBook = mergeAddressBook(Array.isArray(parsed.addressBook) ? parsed.addressBook : []);
      state.manualTasks = Array.isArray(parsed.manualTasks) ? parsed.manualTasks.map(normalizeManualTaskRecord).filter((task) => task.title || task.client) : [];
      state.deliveryOrigin = clean(parsed.deliveryOrigin) || "Laboratorio Dental G, Messina";
      state.deliverySelection = parsed.deliverySelection && typeof parsed.deliverySelection === "object" ? parsed.deliverySelection : {};
      state.deliveryConfirmedJobIds = Array.isArray(parsed.deliveryConfirmedJobIds) ? parsed.deliveryConfirmedJobIds.map((id) => String(id)) : [];
      state.deliveryConfirmedAt = clean(parsed.deliveryConfirmedAt);
      state.deliveryActionFilter = clean(parsed.deliveryActionFilter) || "all";
      return;
    }
  } catch (error) {
    console.error(error);
  }
  state.jobs = [...IMPORTED_JOBS].map(normalizeJobRecord).filter(isUsefulJobRecord);
  state.addressBook = mergeAddressBook([]);
  state.manualTasks = [];
  state.deliveryOrigin = "Laboratorio Dental G, Messina";
  state.deliverySelection = {};
  state.deliveryConfirmedJobIds = [];
  state.deliveryConfirmedAt = "";
  state.deliveryActionFilter = "all";
  persistState();
}

function mergeSavedJobsWithImported(savedJobs) {
  const importedJobs = [...IMPORTED_JOBS].map(normalizeJobRecord).filter(isUsefulJobRecord);
  const normalizedSavedJobs = Array.isArray(savedJobs)
    ? savedJobs.map(normalizeJobRecord).filter(isUsefulJobRecord)
    : [];

  if (!normalizedSavedJobs.length) {
    return importedJobs;
  }

  const mergedById = new Map();
  importedJobs.forEach((job) => {
    mergedById.set(String(job.id), job);
  });
  normalizedSavedJobs.forEach((job) => {
    mergedById.set(String(job.id), job);
  });

  const mergedJobs = Array.from(mergedById.values());
  const importedEligibleCount = importedJobs.filter((job) => isDeliveryReadyJob(job) || isTrialPickupJob(job) || isRequestedPickupJob(job)).length;
  const mergedEligibleCount = mergedJobs.filter((job) => isDeliveryReadyJob(job) || isTrialPickupJob(job) || isRequestedPickupJob(job)).length;

  // If the browser kept an old local archive that wipes out every delivery-stage job,
  // fall back to the bundled dataset so the giro planner remains operational.
  if (mergedEligibleCount === 0 && importedEligibleCount > 0) {
    return importedJobs;
  }

  return mergedJobs;
}

function persistState() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
    jobs: state.jobs,
    addressBook: state.addressBook,
    manualTasks: state.manualTasks,
    deliveryOrigin: state.deliveryOrigin,
    deliverySelection: state.deliverySelection,
    deliveryConfirmedJobIds: state.deliveryConfirmedJobIds,
    deliveryConfirmedAt: state.deliveryConfirmedAt,
    deliveryActionFilter: state.deliveryActionFilter
  }));
}
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
    urgent: jobUrgentInput.checked, redo: jobRedoInput.checked, saldoDirect: jobSaldoDirettoInput.checked, pickupRequested: jobPickupInput.checked, status: clean(jobStatusInput.value), note: clean(jobNoteInput.value), priceInfo: clean(jobPriceInfoInput.value), pieces: quantity
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
  jobForm.reset(); jobIdHidden.value = ""; jobQuantityInput.value = "1"; jobStandardDaysInput.value = "4"; jobStatusInput.value = "In lavorazione"; jobEntryDateInput.value = isoToday(); jobSaldoDirettoInput.checked = false; jobPickupInput.checked = false; updatePricePreview();
}
function resetFilters() { jobSearchInput.value = ""; jobClientFilterInput.value = ""; jobOperatorFilterInput.value = ""; jobUrgencyFilter.value = ""; tableSearchInput.value = ""; render(); }
function resetArchive() {
  if (!window.confirm("Vuoi davvero ripristinare l'archivio base dei lavori?")) return;
  const archivePassword = window.prompt("Per confermare lo svuotamento archivio inserisci la password.");
  if (archivePassword === null) return;
  if (archivePassword.trim() !== "2026") {
    window.alert("Password non corretta. Archivio non modificato.");
    return;
  }
  state.jobs = [...IMPORTED_JOBS];
  state.manualTasks = [];
  state.deliverySelection = {};
  state.deliveryConfirmedJobIds = [];
  state.deliveryConfirmedAt = "";
  persistState(); resetForm(); initializeAnalysisRange(); render();
}
function exportJson() { downloadFile("inserimento-lavori-senza-excel.json", JSON.stringify(state.jobs, null, 2), "application/json;charset=utf-8;"); }
function exportCsv() {
  const rows = [["id_lavoro","cliente","lavorazione","operatore","listino_unitario","prezzo_applicato_unitario","quantita","totale","stato","ingresso","consegna","uscita_prova","rientro_prova","urgente","rifacimento","saldo_diretto","ritiro_ingresso","info_listino","note"], ...state.jobs.map((job) => [job.jobCode, job.client, job.type, job.operator, job.listPrice, job.appliedPrice, job.quantity, computeAppliedTotal(job), job.status, job.entryDate, job.deliveryDate, job.trialOutDate, job.trialReturnDate, job.urgent ? "Si" : "No", job.redo ? "Si" : "No", job.saldoDirect ? "Si" : "No", job.pickupRequested ? "Si" : "No", job.priceInfo, job.note])];
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
function render() {
  const rows = getFilteredJobs();
  renderFilterStatus(rows);
  renderRecentEntries();
  renderTable(rows);
  renderAnalysis();
  renderDeliveryPlanner();
}

function shouldHideFromOperationalView(job) {
  return !!job.saldoDirect && normalizeText(job.status) === "consegnato";
}

function getFilteredJobs() {
  const quickSearch = normalizeText(jobSearchInput.value);
  const clientFilter = normalizeText(jobClientFilterInput.value);
  const operatorFilter = normalizeText(jobOperatorFilterInput.value);
  const urgencyFilter = jobUrgencyFilter.value;
  const tableSearch = normalizeText(tableSearchInput.value);
  return state.jobs.filter((job) => {
    if (shouldHideFromOperationalView(job)) return false;
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
  const latest = state.jobs
    .filter((job) => !shouldHideFromOperationalView(job))
    .sort((a, b) => compareDate(b.entryDate, a.entryDate))
    .slice(0, 5);
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
      <td data-label="Azioni"><div class="table-actions"><button class="table-action info" type="button" data-action="view" data-id="${escapeHtml(job.id)}">Visualizza</button><button class="table-action secondary" type="button" data-action="edit" data-id="${escapeHtml(job.id)}">Modifica</button><button class="table-action danger" type="button" data-action="delete" data-id="${escapeHtml(job.id)}">Elimina</button></div></td>
    </tr>`).join("") : `<tr><td colspan="10" class="empty-state">Nessun lavoro corrisponde ai filtri selezionati.</td></tr>`;
  jobTableBody.querySelectorAll("[data-action]").forEach((button) => button.addEventListener("click", handleTableAction));
}

function handleTableAction(event) {
  const action = event.currentTarget.dataset.action;
  const id = event.currentTarget.dataset.id;
  const job = state.jobs.find((item) => item.id === id);
  if (!job) return;
  if (action === "view") {
    openJobModal(job);
    return;
  }
  if (action === "edit") {
    populateFormForEdit(job);
    return;
  }
  if (action === "delete") {
    if (!window.confirm(`Eliminare il lavoro ${job.jobCode}?`)) return;
    state.jobs = state.jobs.filter((item) => item.id !== id); persistState(); render();
  }
}

function populateFormForEdit(job) {
  jobIdHidden.value = job.id;
  jobCodeInput.value = job.jobCode || "";
  jobClientInput.value = job.client || "";
  jobTypeInput.value = job.type || "";
  jobOperatorInput.value = job.operator || "";
  jobListPriceInput.value = job.listPrice || "";
  jobStandardDaysInput.value = job.standardDays || "";
  jobQuantityInput.value = job.quantity || 1;
  jobAppliedPriceInput.value = job.appliedPrice || "";
  jobStatusInput.value = job.status || "In lavorazione";
  jobEntryDateInput.value = job.entryDate || "";
  jobDeliveryDateInput.value = job.deliveryDate || "";
  jobTrialOutDateInput.value = job.trialOutDate || "";
  jobTrialReturnDateInput.value = job.trialReturnDate || "";
  jobUrgentInput.checked = !!job.urgent;
  jobRedoInput.checked = !!job.redo;
  jobSaldoDirettoInput.checked = !!job.saldoDirect;
  jobPickupInput.checked = !!job.pickupRequested;
  jobPriceInfoInput.value = job.priceInfo || "";
  jobNoteInput.value = job.note || "";
  updatePricePreview();
  setActiveView("operativita");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function openJobModal(job) {
  activeViewedJobId = job.id;
  isJobModalEditing = false;
  renderJobModal(job);
  jobViewModal.hidden = false;
  jobViewModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeJobModal() {
  activeViewedJobId = "";
  isJobModalEditing = false;
  jobViewModal.hidden = true;
  jobViewModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

function handleEditFromModal() {
  if (!activeViewedJobId) return;
  const job = state.jobs.find((item) => item.id === activeViewedJobId);
  if (!job) return;
  if (!isJobModalEditing) {
    isJobModalEditing = true;
    renderJobModal(job);
    return;
  }
  saveJobFromModal(job);
}

function handleGlobalKeydown(event) {
  if (event.key === "Escape" && !jobViewModal.hidden) {
    if (isJobModalEditing) {
      handleModalSecondaryAction();
      return;
    }
    closeJobModal();
    return;
  }
  if (event.key === "Escape" && !addressModal.hidden) {
    closeAddressModal();
    return;
  }
  if (event.key === "Escape" && manualTaskModal && !manualTaskModal.hidden) {
    closeManualTaskModal();
  }
}

function handleModalSecondaryAction() {
  if (!activeViewedJobId) {
    closeJobModal();
    return;
  }
  if (!isJobModalEditing) {
    closeJobModal();
    return;
  }
  const job = state.jobs.find((item) => item.id === activeViewedJobId);
  if (!job) {
    closeJobModal();
    return;
  }
  isJobModalEditing = false;
  renderJobModal(job);
}

function buildJobModalMarkup(job) {
  const pieces = Number(job.pieces || job.quantity || 1);
  const listUnit = parseNumber(job.listPrice);
  const appliedUnit = parseNumber(job.appliedPrice);
  const listTotal = computeListTotal(job);
  const appliedTotal = computeAppliedTotal(job);
  const fields = [
    ["ID lavoro", job.jobCode],
    ["Cliente", job.client],
    ["Lavorazione", job.type],
    ["Operatore", job.operator || "-"],
    ["Stato lavoro", job.status || "-"],
    ["Quantita'", formatInteger(job.quantity || 1)],
    ["Pezzi", formatInteger(pieces)],
    ["Urgente", job.urgent ? "Si" : "No"],
    ["Rifacimento", job.redo ? "Si" : "No"],
    ["Saldo diretto", job.saldoDirect ? "Si" : "No"],
    ["Ingresso", formatDate(job.entryDate)],
    ["Consegna", formatDate(job.deliveryDate)],
    ["Uscita prova", formatDate(job.trialOutDate)],
    ["Rientro prova", formatDate(job.trialReturnDate)],
    ["Tempo standard", `${formatInteger(job.standardDays || 0)} giorni`],
    ["Listino unitario", formatCurrency(listUnit)],
    ["Applicato unitario", formatCurrency(appliedUnit)],
    ["Totale listino", formatCurrency(listTotal)],
    ["Totale applicato", formatCurrency(appliedTotal)],
    ["Info listino", job.priceInfo || "-"],
    ["Note", job.note || "-"]
  ];

  return `
    <div class="job-modal-grid">
      ${fields.map(([label, value]) => `
        <article class="job-modal-card">
          <span class="job-modal-label">${escapeHtml(label)}</span>
          <strong class="job-modal-value">${escapeHtml(String(value ?? "-"))}</strong>
        </article>
      `).join("")}
    </div>
  `;
}

function renderJobModal(job) {
  jobViewTitle.textContent = `${job.jobCode || "Lavoro"} - ${job.client || "Cliente"}`;
  if (isJobModalEditing) {
    jobViewSubtitle.textContent = "Modifica la scheda direttamente qui e salva senza tornare alla compilazione.";
    jobViewContent.innerHTML = buildJobModalEditMarkup(job);
    jobViewEditButton.textContent = "Salva modifiche";
    jobViewCloseSecondary.textContent = "Annulla modifica";
    wireJobModalEditEvents();
    updateJobModalPricePreview();
    return;
  }
  jobViewSubtitle.textContent = job.type ? `Scheda completa della lavorazione ${job.type}.` : "Scheda completa in sola lettura.";
  jobViewContent.innerHTML = buildJobModalMarkup(job);
  jobViewEditButton.textContent = "Modifica questo lavoro";
  jobViewCloseSecondary.textContent = "Chiudi";
}

function buildJobModalEditMarkup(job) {
  return `
    <div class="job-modal-content-block">
      <p class="job-modal-hint">Puoi aggiornare i campi qui sotto e salvare direttamente dal popup, senza tornare alla compilazione principale.</p>
    </div>
    <div class="job-modal-edit-grid">
      <label><span>ID lavoro</span><input id="jobModalCodeInput" type="text" inputmode="numeric" maxlength="8" value="${escapeHtml(job.jobCode || "")}" required></label>
      <label><span>Cliente</span><input id="jobModalClientInput" type="text" list="clientOptions" value="${escapeHtml(job.client || "")}" required></label>
      <label><span>Lavorazione</span><input id="jobModalTypeInput" type="text" list="workTypeOptions" value="${escapeHtml(job.type || "")}" required></label>
      <label><span>Operatore</span><input id="jobModalOperatorInput" type="text" list="operatorOptions" value="${escapeHtml(job.operator || "")}"></label>
      <label><span>Prezzo listino automatico (EUR)</span><input id="jobModalListPriceInput" type="text" value="${escapeHtml(String(job.listPrice || ""))}" readonly></label>
      <label><span>Tempo standard (giorni)</span><input id="jobModalStandardDaysInput" type="number" min="0" step="1" value="${escapeHtml(String(job.standardDays || 0))}"></label>
      <label><span>Quantita'</span><input id="jobModalQuantityInput" type="number" min="1" step="1" value="${escapeHtml(String(job.quantity || 1))}"></label>
      <label><span>Prezzo manuale applicato (EUR)</span><input id="jobModalAppliedPriceInput" type="text" inputmode="decimal" value="${escapeHtml(String(job.appliedPrice || ""))}" placeholder="Se diverso dal listino"></label>
      <label><span>Prezzo finale</span><input id="jobModalFinalPricePreview" type="text" value="" disabled></label>
      <label>
        <span>Stato lavoro</span>
        <select id="jobModalStatusInput">
          ${["In lavorazione","In prova","Consegnato","Pronto","Sospeso"].map((status) => `<option value="${escapeHtml(status)}"${status === (job.status || "In lavorazione") ? " selected" : ""}>${escapeHtml(status)}</option>`).join("")}
        </select>
      </label>
      <label><span>Data ingresso</span><input id="jobModalEntryDateInput" type="date" value="${escapeHtml(job.entryDate || "")}" required></label>
      <label><span>Data consegna</span><input id="jobModalDeliveryDateInput" type="date" value="${escapeHtml(job.deliveryDate || "")}"></label>
      <label><span>Data uscita prova</span><input id="jobModalTrialOutDateInput" type="date" value="${escapeHtml(job.trialOutDate || "")}"></label>
      <label><span>Data rientro prova</span><input id="jobModalTrialReturnDateInput" type="date" value="${escapeHtml(job.trialReturnDate || "")}"></label>
      <label class="checkbox-field"><input id="jobModalUrgentInput" type="checkbox"${job.urgent ? " checked" : ""}><span>Urgente</span></label>
      <label class="checkbox-field"><input id="jobModalRedoInput" type="checkbox"${job.redo ? " checked" : ""}><span>Rifacimento</span></label>
      <label class="checkbox-field"><input id="jobModalSaldoDirettoInput" type="checkbox"${job.saldoDirect ? " checked" : ""}><span>Saldo diretto</span></label>
      <label class="checkbox-field"><input id="jobModalPickupInput" type="checkbox"${job.pickupRequested ? " checked" : ""}><span>Da ritirare / ingresso</span></label>
      <label class="full-width"><span>Info listino</span><input id="jobModalPriceInfoInput" type="text" value="${escapeHtml(job.priceInfo || "")}" placeholder="Lavorazione libera o non presente nel listino"></label>
      <div class="totals-grid full-width">
        <article class="mini-total">
          <span>Totale listino</span>
          <strong id="jobModalListTotalValue">0,00 EUR</strong>
        </article>
        <article class="mini-total alert">
          <span>Totale applicato</span>
          <strong id="jobModalAppliedTotalValue">EUR 0,00</strong>
        </article>
      </div>
      <label class="full-width"><span>Note</span><textarea id="jobModalNoteInput" rows="4" placeholder="Annotazioni sul lavoro">${escapeHtml(job.note || "")}</textarea></label>
    </div>
  `;
}

function wireJobModalEditEvents() {
  const modalTypeInput = document.getElementById("jobModalTypeInput");
  const modalAppliedPriceInput = document.getElementById("jobModalAppliedPriceInput");
  const modalQuantityInput = document.getElementById("jobModalQuantityInput");
  [modalAppliedPriceInput, modalQuantityInput].forEach((input) => input && input.addEventListener("input", updateJobModalPricePreview));
  if (modalTypeInput) {
    modalTypeInput.addEventListener("change", handleJobModalWorkTypeSelection);
    modalTypeInput.addEventListener("blur", handleJobModalWorkTypeSelection);
  }
}

function handleJobModalWorkTypeSelection() {
  const modalTypeInput = document.getElementById("jobModalTypeInput");
  const modalListPriceInput = document.getElementById("jobModalListPriceInput");
  const modalStandardDaysInput = document.getElementById("jobModalStandardDaysInput");
  const modalPriceInfoInput = document.getElementById("jobModalPriceInfoInput");
  if (!modalTypeInput || !modalListPriceInput || !modalStandardDaysInput || !modalPriceInfoInput) return;
  const key = normalizeText(modalTypeInput.value);
  if (!key || !WORK_TYPE_META.has(key)) return;
  const meta = WORK_TYPE_META.get(key);
  modalListPriceInput.value = meta.listPrice ? String(meta.listPrice).replace(".", ",") : "";
  modalStandardDaysInput.value = meta.standardDays ? String(meta.standardDays) : "4";
  if (!modalPriceInfoInput.value && meta.priceInfo) {
    modalPriceInfoInput.value = meta.priceInfo;
  }
  updateJobModalPricePreview();
}

function updateJobModalPricePreview() {
  const modalListPriceInput = document.getElementById("jobModalListPriceInput");
  const modalAppliedPriceInput = document.getElementById("jobModalAppliedPriceInput");
  const modalQuantityInput = document.getElementById("jobModalQuantityInput");
  const modalFinalPricePreview = document.getElementById("jobModalFinalPricePreview");
  const modalListTotalValue = document.getElementById("jobModalListTotalValue");
  const modalAppliedTotalValue = document.getElementById("jobModalAppliedTotalValue");
  if (!modalListPriceInput || !modalQuantityInput || !modalFinalPricePreview || !modalListTotalValue || !modalAppliedTotalValue) return;
  const quantity = Math.max(parseInt(modalQuantityInput.value || "1", 10) || 1, 1);
  const listPrice = parseNumber(modalListPriceInput.value);
  const applied = modalAppliedPriceInput && modalAppliedPriceInput.value ? parseNumber(modalAppliedPriceInput.value) : listPrice;
  const listTotal = listPrice * quantity;
  const appliedTotal = applied * quantity;
  modalFinalPricePreview.value = appliedTotal ? formatCurrency(appliedTotal) : "Calcolato automaticamente";
  modalListTotalValue.textContent = formatCurrency(listTotal);
  modalAppliedTotalValue.textContent = `EUR ${formatDecimal(appliedTotal)}`;
}

function saveJobFromModal(job) {
  const modalCodeInput = document.getElementById("jobModalCodeInput");
  const modalClientInput = document.getElementById("jobModalClientInput");
  const modalTypeInput = document.getElementById("jobModalTypeInput");
  const modalOperatorInput = document.getElementById("jobModalOperatorInput");
  const modalListPriceInput = document.getElementById("jobModalListPriceInput");
  const modalStandardDaysInput = document.getElementById("jobModalStandardDaysInput");
  const modalQuantityInput = document.getElementById("jobModalQuantityInput");
  const modalAppliedPriceInput = document.getElementById("jobModalAppliedPriceInput");
  const modalStatusInput = document.getElementById("jobModalStatusInput");
  const modalEntryDateInput = document.getElementById("jobModalEntryDateInput");
  const modalDeliveryDateInput = document.getElementById("jobModalDeliveryDateInput");
  const modalTrialOutDateInput = document.getElementById("jobModalTrialOutDateInput");
  const modalTrialReturnDateInput = document.getElementById("jobModalTrialReturnDateInput");
  const modalUrgentInput = document.getElementById("jobModalUrgentInput");
  const modalRedoInput = document.getElementById("jobModalRedoInput");
  const modalSaldoDirettoInput = document.getElementById("jobModalSaldoDirettoInput");
  const modalPickupInput = document.getElementById("jobModalPickupInput");
  const modalPriceInfoInput = document.getElementById("jobModalPriceInfoInput");
  const modalNoteInput = document.getElementById("jobModalNoteInput");
  if (!modalCodeInput || !modalClientInput || !modalTypeInput || !modalEntryDateInput) return;

  const sanitizedCode = clean(modalCodeInput.value).replace(/\D/g, "").slice(0, 8);
  modalCodeInput.value = sanitizedCode;
  const quantity = Math.max(parseInt(modalQuantityInput?.value || "1", 10) || 1, 1);
  const listPrice = parseNumber(modalListPriceInput?.value || 0);
  const appliedUnitPrice = modalAppliedPriceInput?.value ? parseNumber(modalAppliedPriceInput.value) : listPrice;
  const updatedJob = {
    ...job,
    jobCode: sanitizedCode,
    client: clean(modalClientInput.value),
    type: clean(modalTypeInput.value),
    operator: clean(modalOperatorInput?.value),
    listPrice,
    appliedPrice: appliedUnitPrice,
    quantity,
    standardDays: parseInt(modalStandardDaysInput?.value || "0", 10) || 0,
    entryDate: modalEntryDateInput.value,
    deliveryDate: modalDeliveryDateInput?.value || "",
    trialOutDate: modalTrialOutDateInput?.value || "",
    trialReturnDate: modalTrialReturnDateInput?.value || "",
    urgent: !!modalUrgentInput?.checked,
    redo: !!modalRedoInput?.checked,
    saldoDirect: !!modalSaldoDirettoInput?.checked,
    pickupRequested: !!modalPickupInput?.checked,
    status: clean(modalStatusInput?.value),
    note: clean(modalNoteInput?.value),
    priceInfo: clean(modalPriceInfoInput?.value),
    pieces: quantity
  };

  if (!updatedJob.jobCode || !updatedJob.client || !updatedJob.type || !updatedJob.entryDate) {
    window.alert("Completa almeno ID lavoro, cliente, lavorazione e data ingresso.");
    return;
  }
  if (!/^\d{1,8}$/.test(updatedJob.jobCode)) {
    window.alert("L'ID lavoro deve contenere solo numeri e massimo 8 cifre.");
    return;
  }
  const index = state.jobs.findIndex((item) => item.id === updatedJob.id);
  if (index < 0) return;
  state.jobs.splice(index, 1, updatedJob);
  persistState();
  render();
  isJobModalEditing = false;
  renderJobModal(updatedJob);
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
    const saldoRows = rows.filter((job) => !!job.saldoDirect);
    const grouped = summarizeGroupMetrics(saldoRows, (job) => job.client, (job) => computeAppliedTotal(job));
    renderUniformAnalysis({
      title: "Saldo diretto",
      intro: "Lavori marcati come saldo diretto, utili per controllare pezzi, valore e studi coinvolti fuori dal ciclo ordinario.",
      summaryHeaders: ["Studi", "Lavori", "Pezzi", "Valore saldo diretto", "Ticket medio"],
      summaryValues: [String(grouped.length), String(saldoRows.length), formatInteger(sumPieces(saldoRows)), formatCurrency(sumApplied(saldoRows)), formatCurrency(saldoRows.length ? sumApplied(saldoRows) / saldoRows.length : 0)],
      tableHeaders: ["Studio", "Lavori", "Pezzi", "Valore", "Ultimo ingresso"],
      tableRows: grouped.map((item) => ({ cells: [item.label, formatInteger(item.jobs), formatInteger(item.pieces), formatCurrency(item.value), formatDate(item.latestDate)], detailKey: item.label })),
      detailRecords: saldoRows.map((job) => buildJobDetailRecord(job, job.client, [`Saldo diretto`, `Stato ${job.status || "-"}`, `Ingresso ${formatDate(job.entryDate)}`])),
      detailEmptyLabel: "Nessun lavoro marcato come saldo diretto nel periodo selezionato."
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
            <tr>${summaryValues.map((value, index) => `<td data-label="${escapeHtml(summaryHeaders[index] || "")}">${escapeHtml(value)}</td>`).join("")}</tr>
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
    status: normalizeImportedText(job.status) || "In lavorazione",
    saldoDirect: !!job.saldoDirect,
    pickupRequested: !!job.pickupRequested
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
              <td data-label="${escapeHtml(itemLabel === "Lavorazione" ? "Pezzi totali" : "Valore totale")}">${escapeHtml(formatParetoMetric(totalValue, itemLabel))}</td>
              <td data-label="${escapeHtml(itemLabel === "Lavorazione" ? "Lavorazioni a listino" : "Medici totali")}">${escapeHtml(String(groupedRows.length))}</td>
              <td data-label="${escapeHtml(groupLabel)}">${escapeHtml(String(activeRows.length))}</td>
              <td data-label="${escapeHtml(`Top 2 ${itemLabel.toLowerCase()}${itemLabel === "Lavorazione" ? "" : "i"}`)}">${escapeHtml(formatParetoMetric(topTwoValue, itemLabel))}</td>
              <td data-label="Incidenza Top 2">${escapeHtml(`${formatDecimal(incidenceTopTwo)}%`)}</td>
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
              <td data-label="Studi">${escapeHtml(String(grouped.length))}</td>
              <td data-label="Lavori">${escapeHtml(String(deliveredRows.length))}</td>
              <td data-label="Pezzi">${escapeHtml(String(totalPieces))}</td>
              <td data-label="Tempo medio complessivo">${escapeHtml(`${formatDecimal(avgOverall)} giorni`)}</td>
              <td data-label="Righe da verificare">${escapeHtml(String(flaggedRows.length))}</td>
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
  if (filterKey === "sospeso" || filterKey === "sospese") return isSuspended;
  if (filterKey === "urgenti") return !!job.urgent;
  if (filterKey === "rifacimenti") return !!job.redo;

  return haystack.includes(filterKey);
}

function normalizeAddressRecord(record) {
  return {
    clientName: normalizeImportedText(record.clientName),
    address: normalizeImportedText(record.address),
    city: normalizeImportedText(record.city),
    province: normalizeImportedText(record.province),
    verified: !!record.verified,
    notes: normalizeImportedText(record.notes)
  };
}

function normalizeManualTaskRecord(task) {
  const scheduledAt = clean(task.scheduledAt) || clean(task.createdAt);
  return {
    id: clean(task.id) || `manual-${Date.now()}`,
    title: normalizeImportedText(task.title),
    client: normalizeImportedText(task.client),
    address: normalizeImportedText(task.address),
    city: normalizeImportedText(task.city),
    province: normalizeImportedText(task.province).toUpperCase(),
    kind: normalizeImportedText(task.kind) || "Commissione",
    timeSlot: normalizeImportedText(task.timeSlot),
    note: normalizeImportedText(task.note),
    urgent: !!task.urgent,
    scheduledAt: scheduledAt || new Date().toISOString()
  };
}

function createManualTaskEntry(task) {
  const normalizedTask = normalizeManualTaskRecord(task);
  const clientLabel = normalizedTask.client || normalizedTask.title || "Commissione manuale";
  return {
    id: normalizedTask.id,
    client: clientLabel,
    type: normalizedTask.title || normalizedTask.kind || "Commissione",
    jobCode: normalizedTask.kind || "Commissione",
    status: "Commissione",
    note: normalizedTask.note,
    urgent: !!normalizedTask.urgent,
    pickupRequested: false,
    manualTask: true,
    manualTaskData: normalizedTask,
    addressRecord: {
      clientName: clientLabel,
      address: normalizedTask.address,
      city: normalizedTask.city,
      province: normalizedTask.province,
      verified: !!normalizedTask.address,
      notes: normalizedTask.note
    },
    pieces: 1,
    quantity: 1,
    timeSlot: normalizedTask.timeSlot
  };
}

function loadDefaultAddressBook() {
  return DEFAULT_ADDRESS_BOOK.map(normalizeAddressRecord);
}

function mergeAddressBook(savedAddressBook) {
  const defaults = loadDefaultAddressBook();
  const savedMap = new Map(
    savedAddressBook
      .map(normalizeAddressRecord)
      .filter((record) => record.clientName)
      .map((record) => [normalizeText(record.clientName), record])
  );

  const merged = defaults.map((record) => {
    const saved = savedMap.get(normalizeText(record.clientName));
    return saved ? { ...record, ...saved } : record;
  });

  const mergedKeys = new Set(merged.map((record) => normalizeText(record.clientName)));
  savedMap.forEach((record, key) => {
    if (!mergedKeys.has(key)) {
      merged.push(record);
    }
  });

  state.jobs
    .map((job) => normalizeImportedText(job.client))
    .filter(Boolean)
    .forEach((clientName) => {
      const key = normalizeText(clientName);
      if (mergedKeys.has(key)) return;
      merged.push({
        clientName,
        address: "",
        city: "",
        province: "",
        verified: false,
        notes: ""
      });
      mergedKeys.add(key);
    });

  return merged.sort((a, b) => normalizeText(a.clientName).localeCompare(normalizeText(b.clientName)));
}

function handleDeliveryOriginChange() {
  state.deliveryOrigin = clean(deliveryOriginInput.value);
  persistState();
  renderDeliveryPlanner();
}

function getDeliveryEligibleJobs() {
  const eligibleJobs = state.jobs.filter((job) => {
    return isDeliveryReadyJob(job) || isTrialPickupJob(job) || isRequestedPickupJob(job);
  });
  const manualEntries = state.manualTasks.map(createManualTaskEntry);
  return [...eligibleJobs, ...manualEntries];
}

function normalizeDeliveryFilterValue(filterValue) {
  const normalized = String(filterValue || "all").trim().toLowerCase();
  if (["all", "delivery", "trial", "pickup", "manual"].includes(normalized)) return normalized;
  return "all";
}

function isDeliveryReadyJob(job) {
  return normalizeText(job.status) === "pronto";
}

function isTrialPickupJob(job) {
  const status = normalizeText(job.status);
  return status === "in prova" || (!!job.trialOutDate && !job.trialReturnDate);
}

function isRequestedPickupJob(job) {
  return !!job.pickupRequested;
}

function getDeliveryActionType(job) {
  if (job.manualTask) return "manual";
  if (isDeliveryReadyJob(job)) return "delivery";
  if (isTrialPickupJob(job)) return "trial";
  if (isRequestedPickupJob(job)) return "pickup";
  return "";
}

function matchesDeliveryActionFilter(job, filterValue = state.deliveryActionFilter || "all") {
  const normalizedFilter = normalizeDeliveryFilterValue(filterValue);
  if (normalizedFilter === "all") return true;
  return getDeliveryActionType(job) === normalizedFilter;
}

function getVisibleDeliveryJobs(filterValue = state.deliveryActionFilter || "all") {
  return getDeliveryEligibleJobs().filter((job) => matchesDeliveryActionFilter(job, filterValue));
}

function syncDeliverySelection() {
  const eligibleJobs = getDeliveryEligibleJobs();
  const eligibleIds = new Set(eligibleJobs.map((job) => String(job.id)));
  const nextSelection = {};
  const previousSelection = JSON.stringify(state.deliverySelection);
  const previousConfirmed = JSON.stringify(state.deliveryConfirmedJobIds);

  eligibleJobs.forEach((job) => {
    const key = String(job.id);
    nextSelection[key] = Object.prototype.hasOwnProperty.call(state.deliverySelection, key)
      ? !!state.deliverySelection[key]
      : true;
  });

  state.deliverySelection = nextSelection;
  state.deliveryConfirmedJobIds = state.deliveryConfirmedJobIds.filter((id) => eligibleIds.has(String(id)));
  if (!state.deliveryConfirmedJobIds.length) {
    state.deliveryConfirmedAt = "";
  }
  if (previousSelection !== JSON.stringify(state.deliverySelection) || previousConfirmed !== JSON.stringify(state.deliveryConfirmedJobIds)) {
    persistState();
  }
}

function getSelectedDeliveryJobs() {
  syncDeliverySelection();
  return getDeliveryEligibleJobs().filter((job) => !!state.deliverySelection[String(job.id)]);
}

function getSelectedVisibleDeliveryJobs(filterValue = state.deliveryActionFilter || "all") {
  return getSelectedDeliveryJobs().filter((job) => matchesDeliveryActionFilter(job, filterValue));
}

function getConfirmedDeliveryJobs() {
  if (!state.deliveryConfirmedJobIds.length) return [];
  const confirmedSet = new Set(state.deliveryConfirmedJobIds.map((id) => String(id)));
  return getDeliveryEligibleJobs().filter((job) => confirmedSet.has(String(job.id)));
}

function getConfirmedVisibleDeliveryJobs(filterValue = state.deliveryActionFilter || "all") {
  return getConfirmedDeliveryJobs().filter((job) => matchesDeliveryActionFilter(job, filterValue));
}

function setAllDeliverySelection(value) {
  syncDeliverySelection();
  getVisibleDeliveryJobs().forEach((job) => {
    state.deliverySelection[String(job.id)] = value;
  });
  state.deliveryConfirmedJobIds = [];
  state.deliveryConfirmedAt = "";
  persistState();
  renderDeliveryPlanner();
}

function handleDeliverySelectionChange(event) {
  const jobCheckbox = event.target.closest("[data-delivery-job]");
  if (jobCheckbox) {
    syncDeliverySelection();
    state.deliverySelection[jobCheckbox.getAttribute("data-delivery-job")] = !!jobCheckbox.checked;
    state.deliveryConfirmedJobIds = [];
    state.deliveryConfirmedAt = "";
    persistState();
    renderDeliveryPlanner();
    return;
  }

  const stopCheckbox = event.target.closest("[data-delivery-stop]");
  if (stopCheckbox) {
    const stopClient = stopCheckbox.getAttribute("data-delivery-stop") || "";
    syncDeliverySelection();
    getVisibleDeliveryJobs()
      .filter((job) => normalizeImportedText(job.client) === normalizeImportedText(stopClient))
      .forEach((job) => {
        state.deliverySelection[String(job.id)] = !!stopCheckbox.checked;
      });
    state.deliveryConfirmedJobIds = [];
    state.deliveryConfirmedAt = "";
    persistState();
    renderDeliveryPlanner();
  }
}

function handleDeliveryListClick(event) {
  const actionButton = event.target.closest("[data-delivery-action]");
  if (!actionButton) return;
  const action = actionButton.getAttribute("data-delivery-action");
  if (action === "stop-all") {
    const clientName = actionButton.getAttribute("data-delivery-client") || "";
    syncDeliverySelection();
    getVisibleDeliveryJobs()
      .filter((job) => normalizeImportedText(job.client) === normalizeImportedText(clientName))
      .forEach((job) => {
        state.deliverySelection[String(job.id)] = true;
      });
    state.deliveryConfirmedJobIds = [];
    state.deliveryConfirmedAt = "";
    persistState();
    renderDeliveryPlanner();
    return;
  }

  if (action === "edit-address") {
    openAddressModal(actionButton.getAttribute("data-delivery-client") || "");
  }
}

function confirmDeliveryPlan() {
  const selectedJobs = getSelectedDeliveryJobs();
  if (!selectedJobs.length) {
    window.alert("Seleziona almeno un lavoro o una commissione prima di confermare il giro.");
    return;
  }
  state.deliveryConfirmedJobIds = selectedJobs.map((job) => String(job.id));
  state.deliveryConfirmedAt = new Date().toLocaleString("it-IT");
  persistState();
  renderDeliveryPlanner();
}

function printDeliveryPlan() {
  const confirmedJobs = getConfirmedDeliveryJobs();
  const jobsToPrint = confirmedJobs.length ? confirmedJobs : getSelectedDeliveryJobs();
  if (!jobsToPrint.length) {
    window.alert("Non ci sono lavori o commissioni selezionati da stampare.");
    return;
  }

  const stops = groupDeliveryStops(jobsToPrint);
  const printableMarkup = buildPrintableDeliveryMarkupV2(stops);
  const printWindow = window.open("", "_blank", "width=1200,height=900");
  if (!printWindow) {
    window.alert("Consenti l'apertura del popup per stampare il giro.");
    return;
  }

  printWindow.document.write(printableMarkup);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function getAddressRecordByClient(clientName) {
  const target = normalizeText(clientName);
  return state.addressBook.find((record) => normalizeText(record.clientName) === target) || null;
}

function groupDeliveryStops(jobs) {
  const grouped = new Map();
  jobs.forEach((job) => {
    const key = normalizeImportedText(job.client) || "Cliente non definito";
    if (!grouped.has(key)) {
      const addressRecord = job.manualTask ? normalizeAddressRecord(job.addressRecord || {}) : getAddressRecordByClient(key);
      grouped.set(key, {
        client: key,
        addressRecord,
        jobs: [],
        urgent: false,
        deliveryCount: 0,
        trialCount: 0,
        pickupCount: 0,
        manualCount: 0
      });
    }
    const bucket = grouped.get(key);
    bucket.jobs.push(job);
    bucket.urgent = bucket.urgent || !!job.urgent;
    const status = normalizeText(job.status);
    if (status === "pronto") bucket.deliveryCount += 1;
    if (status === "in prova") bucket.trialCount += 1;
    if (job.pickupRequested) bucket.pickupCount += 1;
    if (job.manualTask) bucket.manualCount += 1;
  });

  return optimizeDeliveryStops(
    Array.from(grouped.values()).map((stop) => ({
      ...stop,
      jobs: sortDeliveryJobs(stop.jobs),
      addressText: stop.addressRecord?.address || stop.addressRecord?.notes || "Indirizzo non disponibile",
      totalJobs: stop.jobs.length,
      totalPieces: stop.jobs.reduce((sum, job) => sum + Number(job.pieces || job.quantity || 1), 0),
      cityLabel: buildStopCityLabel(stop.addressRecord),
      geoPoint: getStopGeoPoint(stop.addressRecord, stop.client)
    }))
  )
    .map((stop, index) => ({
      ...stop,
      sequence: index + 1
    }));
}

function sortDeliveryJobs(jobs) {
  const actionPriority = { pickup: 0, trial: 1, delivery: 2 };
  return [...jobs].sort((a, b) => {
    const actionDelta = (actionPriority[getDeliveryActionType(a)] ?? 9) - (actionPriority[getDeliveryActionType(b)] ?? 9);
    if (actionDelta !== 0) return actionDelta;
    if (!!a.urgent !== !!b.urgent) return a.urgent ? -1 : 1;
    const dateA = a.deliveryDate || a.trialOutDate || a.entryDate || "";
    const dateB = b.deliveryDate || b.trialOutDate || b.entryDate || "";
    if (dateA !== dateB) return dateA.localeCompare(dateB);
    return normalizeText(`${a.client} ${a.type} ${a.jobCode}`).localeCompare(normalizeText(`${b.client} ${b.type} ${b.jobCode}`));
  });
}

function buildStopCityLabel(addressRecord) {
  const city = normalizeImportedText(addressRecord?.city);
  const province = normalizeImportedText(addressRecord?.province).toUpperCase();
  return [city, province].filter(Boolean).join(" - ");
}

function normalizeLocationKey(city, province) {
  const cityKey = normalizeText(city).replace(/\s+/g, " ");
  const provinceKey = normalizeText(province).slice(0, 2);
  if (!cityKey) return "";
  return `${cityKey}|${provinceKey}`;
}

function extractProvinceFromText(value) {
  const match = clean(value).match(/\(([A-Za-z]{2})\)\s*$/);
  return match ? match[1].toUpperCase() : "";
}

function guessCityFromAddress(address) {
  const normalizedAddress = normalizeImportedText(address);
  if (!normalizedAddress) return "";
  const withoutProvince = normalizedAddress.replace(/\([A-Za-z]{2}\)\s*$/g, "");
  const parts = withoutProvince.split(",").map((part) => clean(part)).filter(Boolean);
  if (!parts.length) return "";
  const lastPart = parts[parts.length - 1].replace(/\b\d{5}\b/g, "").trim();
  return normalizeImportedText(lastPart);
}

function getLocationPoint(city, province) {
  const key = normalizeLocationKey(city, province);
  return DELIVERY_CITY_COORDINATES[key] || null;
}

function getStopGeoPoint(addressRecord, fallbackClientName = "") {
  const city = normalizeImportedText(addressRecord?.city) || guessCityFromAddress(addressRecord?.address || addressRecord?.notes || fallbackClientName);
  const province = normalizeImportedText(addressRecord?.province).toUpperCase() || extractProvinceFromText(addressRecord?.address || "");
  return getLocationPoint(city, province);
}

function getOriginGeoPoint() {
  const originText = clean(state.deliveryOrigin);
  if (!originText) return getLocationPoint("Messina", "ME");
  const province = extractProvinceFromText(originText) || "ME";
  const guessedCity = guessCityFromAddress(originText) || "Messina";
  return getLocationPoint(guessedCity, province) || getLocationPoint("Messina", "ME");
}

function haversineDistanceKm(pointA, pointB) {
  if (!pointA || !pointB) return Number.POSITIVE_INFINITY;
  const toRadians = (value) => value * Math.PI / 180;
  const latDelta = toRadians(pointB.lat - pointA.lat);
  const lonDelta = toRadians(pointB.lon - pointA.lon);
  const latA = toRadians(pointA.lat);
  const latB = toRadians(pointB.lat);
  const haversine = Math.sin(latDelta / 2) ** 2
    + Math.cos(latA) * Math.cos(latB) * Math.sin(lonDelta / 2) ** 2;
  return 6371 * 2 * Math.atan2(Math.sqrt(haversine), Math.sqrt(1 - haversine));
}

function optimizeDeliveryStops(stops) {
  if (stops.length <= 1) return stops;
  const geocoded = stops.filter((stop) => !!stop.geoPoint);
  const fallback = stops
    .filter((stop) => !stop.geoPoint)
    .sort((a, b) => normalizeText(`${a.cityLabel} ${a.addressText} ${a.client}`).localeCompare(normalizeText(`${b.cityLabel} ${b.addressText} ${b.client}`)));

  if (!geocoded.length) {
    return fallback;
  }

  const remaining = [...geocoded];
  const ordered = [];
  let currentPoint = getOriginGeoPoint();

  while (remaining.length) {
    let nextIndex = 0;
    let nextScore = Number.POSITIVE_INFINITY;

    remaining.forEach((stop, index) => {
      const baseDistance = haversineDistanceKm(currentPoint, stop.geoPoint);
      const urgencyBonus = stop.urgent ? 5 : 0;
      const actionBonus = stop.pickupCount > 0 ? 2 : 0;
      const score = baseDistance - urgencyBonus - actionBonus;
      if (score < nextScore) {
        nextIndex = index;
        nextScore = score;
      }
    });

    const [nextStop] = remaining.splice(nextIndex, 1);
    ordered.push(nextStop);
    currentPoint = nextStop.geoPoint || currentPoint;
  }

  return [...ordered, ...fallback];
}

function buildDeliveryEmbedUrl() {
  return "";
}

function buildDeliveryMapsUrl(stops) {
  const mappedStops = stops
    .map((stop) => ({
      ...stop,
      mapAddress: clean(stop.addressRecord?.address) || clean(stop.addressText)
    }))
    .filter((stop) => stop.mapAddress && normalizeText(stop.mapAddress) !== "indirizzo non disponibile");
  if (!mappedStops.length) return "";

  const encodedOrigin = encodeURIComponent(state.deliveryOrigin || mappedStops[0].mapAddress);
  if (mappedStops.length === 1) {
    return `https://www.google.com/maps/dir/?api=1&origin=${encodedOrigin}&destination=${encodeURIComponent(mappedStops[0].mapAddress)}&travelmode=driving`;
  }

  const origin = state.deliveryOrigin || mappedStops[0].mapAddress;
  const destination = mappedStops[mappedStops.length - 1].mapAddress;
  const waypointItems = mappedStops.slice(state.deliveryOrigin ? 0 : 1, -1).map((stop) => stop.mapAddress).filter(Boolean);
  const query = [
    "https://www.google.com/maps/dir/?api=1",
    `origin=${encodeURIComponent(origin)}`,
    `destination=${encodeURIComponent(destination)}`,
    waypointItems.length ? `waypoints=${encodeURIComponent(waypointItems.join("|"))}` : "",
    "travelmode=driving"
  ].filter(Boolean).join("&");
  return query;
}

function openDeliveryRouteInMaps() {
  const plannedJobs = getConfirmedDeliveryJobs();
  const selectedJobs = plannedJobs.length ? plannedJobs : getSelectedDeliveryJobs();
  const stops = groupDeliveryStops(selectedJobs);
  const url = buildDeliveryMapsUrl(stops);
  if (!url) {
    window.alert("Non ci sono ancora tappe con indirizzo disponibile da aprire su Google Maps.");
    return;
  }
  window.open(url, "_blank", "noopener,noreferrer");
}

function renderDeliveryPlanner() {
  if (!deliveryKpiGrid) return;
  syncDeliverySelection();
  deliveryOriginInput.value = state.deliveryOrigin || "";
  state.deliveryActionFilter = normalizeDeliveryFilterValue(state.deliveryActionFilter || "all");
  deliveryActionFilterInput.value = state.deliveryActionFilter;
  const eligibleJobs = getDeliveryEligibleJobs();
  const filteredEligibleJobs = getVisibleDeliveryJobs();
  const selectedJobs = getSelectedVisibleDeliveryJobs();
  const confirmedJobs = getConfirmedVisibleDeliveryJobs();
  const jobsForPreview = confirmedJobs.length ? confirmedJobs : selectedJobs;
  const stops = groupDeliveryStops(jobsForPreview);
  confirmDeliveryPlanButton.textContent = confirmedJobs.length ? "Aggiorna lista giro" : "Conferma lista giro";
  printDeliveryPlanButton.disabled = !jobsForPreview.length;
  openDeliveryRouteButton.disabled = !jobsForPreview.length;
  renderDeliveryKpis(stops, jobsForPreview, filteredEligibleJobs);
  renderDeliveryRoutePreview(stops, confirmedJobs.length > 0);
  renderDeliverySelectionSummary(filteredEligibleJobs, selectedJobs, confirmedJobs);
  renderDeliveryConfirmedSummary(confirmedJobs);
  renderDeliveryGroupedList(groupDeliveryStops(filteredEligibleJobs));
  renderDeliveryAddressBook();
  renderManualTaskList();
}

function renderDeliveryKpis(stops, jobs, eligibleJobs) {
  const kpis = [
    { title: "Attivita' candidate", value: eligibleJobs.length },
    { title: "Attivita' nel giro", value: jobs.length },
    { title: "Tappe nel giro", value: stops.length },
    { title: "Da consegnare", value: jobs.filter((job) => normalizeText(job.status) === "pronto").length },
    { title: "Da ritirare in prova", value: jobs.filter((job) => normalizeText(job.status) === "in prova").length },
    { title: "Ritiri su richiesta", value: jobs.filter((job) => !!job.pickupRequested).length },
    { title: "Commissioni manuali", value: jobs.filter((job) => !!job.manualTask).length }
  ];
  deliveryKpiGrid.innerHTML = kpis.map((item) => `<article class="kpi-card"><span>${escapeHtml(item.title)}</span><strong>${escapeHtml(String(item.value))}</strong></article>`).join("");
}

function renderDeliveryRoutePreview(stops, isConfirmed) {
  if (!stops.length) {
    deliveryHint.textContent = "Seleziona almeno un'attivita' per comporre il giro. Dopo la conferma potrai stamparlo e aprirlo su Google Maps.";
    deliveryRoutePreview.innerHTML = `<div class="delivery-empty">Nessuna attivita' selezionata: scegli le tappe da includere nel giro del fattorino.</div>`;
    return;
  }

  const embedUrl = buildDeliveryEmbedUrl(stops);
  const previewItems = stops.map((stop, index) => `
      <div class="delivery-preview-stop">
        <span class="delivery-preview-index">${index + 1}</span>
        <div>
          <strong>${escapeHtml(stop.client)}</strong>
          <p>${escapeHtml(stop.addressText)}</p>
          <div class="delivery-preview-tags">
            ${stop.deliveryCount ? `<span>Consegne ${escapeHtml(String(stop.deliveryCount))}</span>` : ""}
            ${stop.trialCount ? `<span>Ritiri prova ${escapeHtml(String(stop.trialCount))}</span>` : ""}
            ${stop.pickupCount ? `<span>Ingressi ${escapeHtml(String(stop.pickupCount))}</span>` : ""}
            ${stop.manualCount ? `<span>Commissioni ${escapeHtml(String(stop.manualCount))}</span>` : ""}
            ${stop.cityLabel ? `<span>${escapeHtml(stop.cityLabel)}</span>` : ""}
          </div>
        </div>
      </div>
    `).join("");

  const routeUrl = buildDeliveryMapsUrl(stops);
  deliveryHint.textContent = `${isConfirmed ? "Giro confermato" : "Anteprima giro"} con ${formatInteger(stops.length)} tappe raggruppate per studio. L'ordine e' stimato per vicinanza progressiva tra citta' e indirizzi disponibili, cosi' il percorso resta piu' sensato prima dell'apertura su Google Maps.`;
  deliveryRoutePreview.innerHTML = `
      <div class="delivery-route-shell ${isConfirmed ? "confirmed" : ""}">
        <div class="delivery-route-header">
          <strong>${isConfirmed ? "Giro confermato" : "Percorso suggerito"}</strong>
          <span>${escapeHtml(state.deliveryOrigin || "Partenza non impostata")}</span>
        </div>
        <p class="delivery-route-note">Se ci sono indirizzi mancanti, le tappe senza coordinate stimate restano in fondo e puoi correggerle dalla rubrica medici.</p>
        <div class="delivery-preview-list">${previewItems}</div>
        ${embedUrl ? `
          <div class="delivery-map-preview">
            <div class="delivery-map-frame-shell">
              <iframe
                class="delivery-map-frame"
                src="${escapeHtml(embedUrl)}"
                loading="lazy"
                referrerpolicy="no-referrer-when-downgrade"
                title="Anteprima mappa giro consegne"></iframe>
            </div>
          </div>
        ` : ""}
        ${routeUrl ? `<a class="delivery-maps-link" href="${escapeHtml(routeUrl)}" target="_blank" rel="noopener noreferrer">Apri percorso completo su Google Maps</a>` : ""}
      </div>
    `;
}

function renderDeliverySelectionSummary(eligibleJobs, selectedJobs, confirmedJobs) {
  if (!deliverySelectionSummary) return;
  const selectedStops = groupDeliveryStops(selectedJobs);
  const selectedDeliveries = selectedJobs.filter((job) => getDeliveryActionType(job) === "delivery").length;
  const selectedTrials = selectedJobs.filter((job) => getDeliveryActionType(job) === "trial").length;
  const selectedPickups = selectedJobs.filter((job) => getDeliveryActionType(job) === "pickup").length;
  const selectedManual = selectedJobs.filter((job) => getDeliveryActionType(job) === "manual").length;
  const confirmedLabel = confirmedJobs.length
    ? `Giro confermato con ${formatInteger(confirmedJobs.length)} attivita'`
    : "Lista non ancora confermata";
  deliverySelectionSummary.innerHTML = `
    <div class="delivery-selection-card">
      <strong>${escapeHtml(confirmedLabel)}</strong>
      <div class="delivery-selection-stats">
        <span>${formatInteger(selectedJobs.length)} di ${formatInteger(eligibleJobs.length)} attivita' selezionate</span>
        <span>${formatInteger(selectedStops.length)} tappe attive</span>
        <span>Consegne ${formatInteger(selectedDeliveries)}</span>
        <span>Ritiri prova ${formatInteger(selectedTrials)}</span>
        <span>Ingressi ${formatInteger(selectedPickups)}</span>
        <span>Commissioni ${formatInteger(selectedManual)}</span>
      </div>
    </div>
  `;
}

function renderDeliveryConfirmedSummary(confirmedJobs) {
  if (!deliveryConfirmedSummary) return;
  if (!confirmedJobs.length) {
    deliveryConfirmedSummary.innerHTML = "";
    return;
  }
  const confirmedStops = groupDeliveryStops(confirmedJobs);
  const confirmedDeliveries = confirmedJobs.filter((job) => getDeliveryActionType(job) === "delivery").length;
  const confirmedTrials = confirmedJobs.filter((job) => getDeliveryActionType(job) === "trial").length;
  const confirmedPickups = confirmedJobs.filter((job) => getDeliveryActionType(job) === "pickup").length;
  const confirmedManual = confirmedJobs.filter((job) => getDeliveryActionType(job) === "manual").length;
  deliveryConfirmedSummary.innerHTML = `
    <div class="delivery-confirmed-card">
      <div>
        <p class="section-label">Lista confermata</p>
        <h3>${formatInteger(confirmedStops.length)} tappe confermate</h3>
        <p class="subtitle compact">Consegne ${formatInteger(confirmedDeliveries)} - Ritiri prova ${formatInteger(confirmedTrials)} - Ingressi ${formatInteger(confirmedPickups)} - Commissioni ${formatInteger(confirmedManual)}</p>
        <p class="subtitle compact">${formatInteger(confirmedJobs.length)} attivita' incluse${state.deliveryConfirmedAt ? ` - confermata il ${escapeHtml(state.deliveryConfirmedAt)}` : ""}</p>
      </div>
      <div class="delivery-confirmed-actions">
        <button type="button" onclick="window.printDeliveryPlan()">Stampa elenco operatore</button>
        <button type="button" class="secondary" onclick="window.openDeliveryRouteInMaps()">Apri giro su Maps</button>
      </div>
    </div>
  `;
}

function renderDeliveryGroupedList(stops) {
  if (!stops.length) {
    const eligibleJobs = getDeliveryEligibleJobs();
    const deliveries = eligibleJobs.filter((job) => getDeliveryActionType(job) === "delivery").length;
    const trials = eligibleJobs.filter((job) => getDeliveryActionType(job) === "trial").length;
    const pickups = eligibleJobs.filter((job) => getDeliveryActionType(job) === "pickup").length;
    const manuals = eligibleJobs.filter((job) => getDeliveryActionType(job) === "manual").length;
    deliveryGroupedList.innerHTML = `
      <div class="delivery-empty">
        ${eligibleJobs.length
          ? `Con il filtro attuale non risultano tappe. Disponibili nel gestionale: ${formatInteger(deliveries)} consegne, ${formatInteger(trials)} ritiri prova, ${formatInteger(pickups)} ritiri / ingressi, ${formatInteger(manuals)} commissioni.`
          : "Non ci sono ancora lavori nello stato Pronto, In prova o Da ritirare / ingresso, e non risultano commissioni manuali."}
      </div>`;
    return;
  }

  deliveryGroupedList.innerHTML = stops.map((stop) => `
    <article class="delivery-stop-card ${stop.urgent ? "urgent" : ""}">
      <div class="delivery-stop-header">
        <div>
          <label class="delivery-stop-toggle">
            <input type="checkbox" data-delivery-stop="${escapeHtml(stop.client)}" ${stop.jobs.every((job) => !!state.deliverySelection[String(job.id)]) ? "checked" : ""}>
            <span>Seleziona tappa</span>
          </label>
          <h3>${escapeHtml(`${stop.sequence}. ${stop.client}`)}</h3>
          <p class="delivery-stop-address">${escapeHtml(stop.addressText)}</p>
          <p class="delivery-stop-city">${escapeHtml(stop.cityLabel || "Citta' non definita")}</p>
        </div>
        <div class="delivery-stop-badges">
          ${stop.urgent ? `<span class="status-pill status-alert">Urgente</span>` : ""}
          ${stop.manualCount === stop.totalJobs ? "" : `<button class="delivery-stop-action secondary" type="button" data-delivery-action="edit-address" data-delivery-client="${escapeHtml(stop.client)}">Modifica medico</button>`}
          <button class="delivery-stop-action" type="button" data-delivery-action="stop-all" data-delivery-client="${escapeHtml(stop.client)}">Seleziona tutti</button>
        </div>
      </div>
      <div class="delivery-stop-meta">
        <span>Tappa unica per studio</span>
        <strong>${formatInteger(stop.totalJobs)} lavori</strong>
        <span>${formatInteger(stop.totalPieces)} pezzi</span>
        <span>Consegne ${formatInteger(stop.deliveryCount)}</span>
        <span>Ritiri prova ${formatInteger(stop.trialCount)}</span>
        <span>Ingressi ${formatInteger(stop.pickupCount)}</span>
        <span>Commissioni ${formatInteger(stop.manualCount)}</span>
      </div>
      <ul class="delivery-job-list">
        ${stop.jobs.map((job) => {
          const tags = [];
          if (normalizeText(job.status) === "pronto") tags.push("Consegna");
          if (normalizeText(job.status) === "in prova") tags.push("Ritiro prova");
          if (job.pickupRequested) tags.push("Ingresso / ritiro");
          if (job.manualTask) tags.push(job.manualTaskData?.kind || "Commissione");
          return `
            <li class="delivery-job-item ${state.deliverySelection[String(job.id)] ? "selected" : ""}">
              <div>
                <label class="delivery-job-toggle">
                  <input type="checkbox" data-delivery-job="${escapeHtml(String(job.id))}" ${state.deliverySelection[String(job.id)] ? "checked" : ""}>
                  <span>Includi nel giro</span>
                </label>
                <strong>${escapeHtml(job.jobCode)} - ${escapeHtml(job.type)}</strong>
                ${job.timeSlot ? `<p class="delivery-job-extra">Fascia oraria: ${escapeHtml(job.timeSlot)}</p>` : ""}
                ${job.manualTaskData?.address ? `<p class="delivery-job-extra">Destinazione: ${escapeHtml(job.manualTaskData.address)}</p>` : ""}
                ${job.note ? `<p class="delivery-job-extra">Note: ${escapeHtml(job.note)}</p>` : ""}
                <p>${escapeHtml(tags.join(" | ") || job.status || "-")} · ${escapeHtml(job.client)}</p>
              </div>
            </li>
          `;
        }).join("")}
      </ul>
    </article>
  `).join("");
}

function buildPrintableDeliveryMarkup(stops) {
  const routeUrl = buildDeliveryMapsUrl(stops);
  const routeLines = [
    state.deliveryOrigin || "Partenza non impostata",
    ...stops.map((stop, index) => `${index + 1}. ${stop.client} - ${stop.addressText}`)
  ];
  const routeMarkup = `
    <section class="print-route">
      <h2>Percorso suggerito</h2>
      <ol>
        ${routeLines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}
      </ol>
      ${routeUrl
        ? `<p class="route-link">Apri percorso completo su Maps: ${escapeHtml(routeUrl)}</p>`
        : `<p class="route-link">Percorso Maps non disponibile: servono indirizzi completi per le tappe selezionate.</p>`}
    </section>
  `;
  const stopMarkup = stops.map((stop, index) => `
    <section class="print-stop">
      <header>
        <h2>${index + 1}. ${escapeHtml(stop.client)}</h2>
        <p>${escapeHtml(stop.addressText)}</p>
      </header>
      <table>
        <thead>
          <tr><th>ID</th><th>Lavorazione</th><th>Tipo</th><th>Pezzi</th></tr>
        </thead>
        <tbody>
          ${stop.jobs.map((job) => {
            const tags = [];
            if (normalizeText(job.status) === "pronto") tags.push("Consegna");
            if (normalizeText(job.status) === "in prova") tags.push("Ritiro prova");
            if (job.pickupRequested) tags.push("Ingresso / ritiro");
            return `<tr>
              <td>${escapeHtml(job.jobCode)}</td>
              <td>${escapeHtml(job.type)}</td>
              <td>${escapeHtml(tags.join(" / ") || job.status || "-")}</td>
              <td>${escapeHtml(String(formatInteger(job.pieces || job.quantity || 1)))}</td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </section>
  `).join("");

  return `<!doctype html>
  <html lang="it">
  <head>
    <meta charset="utf-8">
    <title>Giro consegne Dental G</title>
    <style>
      body { font-family: Georgia, serif; padding: 24px; color: #1e1b16; }
      h1 { margin: 0 0 8px; font-size: 32px; }
      .meta { margin-bottom: 24px; color: #4c5f82; }
      .print-route { margin-bottom: 24px; border: 1px solid #d9dfec; border-radius: 16px; padding: 16px; background: #f7f9fd; }
      .print-route h2 { margin: 0 0 10px; font-size: 22px; color: #355381; }
      .print-route ol { margin: 0; padding-left: 22px; }
      .print-route li { margin: 0 0 8px; }
      .route-link { margin-top: 14px; color: #4c5f82; font-size: 14px; word-break: break-all; }
      .print-stop { margin-bottom: 24px; border: 1px solid #d9dfec; border-radius: 16px; padding: 16px; }
      table { width: 100%; border-collapse: collapse; margin-top: 12px; }
      th, td { border-bottom: 1px solid #e4e9f3; padding: 10px 8px; text-align: left; }
      th { color: #355381; font-size: 14px; text-transform: uppercase; }
    </style>
  </head>
  <body>
    <h1>Giro consegne Dental G</h1>
    <p class="meta">Partenza: ${escapeHtml(state.deliveryOrigin || "Non impostata")} · Tappe ${escapeHtml(String(stops.length))}${routeUrl ? ` · Percorso Maps disponibile` : ""}</p>
    ${routeMarkup}
    ${stopMarkup}
  </body>
  </html>`;
}

function buildPrintableDeliveryMarkupV2(stops) {
  const summary = {
    deliveries: stops.reduce((sum, stop) => sum + stop.deliveryCount, 0),
    trials: stops.reduce((sum, stop) => sum + stop.trialCount, 0),
    pickups: stops.reduce((sum, stop) => sum + stop.pickupCount, 0),
    manuals: stops.reduce((sum, stop) => sum + stop.manualCount, 0)
  };

  const routeMarkup = `
    <section class="print-route">
      <h2>Ordine tappe</h2>
      <div class="route-origin">Partenza: ${escapeHtml(state.deliveryOrigin || "Partenza non impostata")}</div>
      <div class="route-stops">
        ${stops.map((stop, index) => `
          <div class="route-stop">
            <span class="route-stop-index">${index + 1}</span>
            <div>
              <strong>${escapeHtml(stop.client)}</strong>
              <p>${escapeHtml(stop.addressText)}</p>
            </div>
          </div>
        `).join("")}
      </div>
    </section>
  `;

  const stopMarkup = stops.map((stop, index) => {
    const missionTags = [];
    if (stop.deliveryCount) missionTags.push(`Consegne ${formatInteger(stop.deliveryCount)}`);
    if (stop.trialCount) missionTags.push(`Ritiri prova ${formatInteger(stop.trialCount)}`);
    if (stop.pickupCount) missionTags.push(`Ingressi ${formatInteger(stop.pickupCount)}`);
    if (stop.manualCount) missionTags.push(`Commissioni ${formatInteger(stop.manualCount)}`);

    return `
      <section class="print-stop">
        <header class="print-stop-head">
          <div>
            <h2>${index + 1}. ${escapeHtml(stop.client)}</h2>
            <p>${escapeHtml(stop.addressText)}</p>
            <p class="print-stop-city">${escapeHtml(stop.cityLabel || "Citta' non definita")}</p>
          </div>
          <div class="print-stop-badges">${missionTags.map((tag) => `<span>${escapeHtml(tag)}</span>`).join("")}</div>
        </header>
        <table>
          <thead>
            <tr><th>Fatto</th><th>Codice</th><th>Attivita'</th><th>Tipo</th><th>Pezzi</th><th>Esito / note</th></tr>
          </thead>
          <tbody>
            ${stop.jobs.map((job) => {
              const tags = [];
              if (normalizeText(job.status) === "pronto") tags.push("Consegna");
              if (normalizeText(job.status) === "in prova") tags.push("Ritiro prova");
              if (job.pickupRequested) tags.push("Ingresso / ritiro");
              if (job.manualTask) tags.push(job.manualTaskData?.kind || "Commissione");
              const extraNotes = [
                job.manualTaskData?.scheduledAt ? `Data commissione: ${formatDateForHumans(job.manualTaskData.scheduledAt)}` : "",
                job.timeSlot ? `Fascia oraria: ${job.timeSlot}` : "",
                job.note || ""
              ].filter(Boolean).join(" - ");
              return `<tr>
                <td><span class="check-box"></span></td>
                <td>${escapeHtml(job.jobCode)}</td>
                <td>${escapeHtml(job.type)}</td>
                <td>${escapeHtml(tags.join(" / ") || job.status || "-")}</td>
                <td>${escapeHtml(String(formatInteger(job.pieces || job.quantity || 1)))}</td>
                <td>${escapeHtml(extraNotes)}</td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
        <div class="print-stop-footer">
          <span>Contatto effettuato</span>
          <span>Consegna completata</span>
          <span>Ritiro completato</span>
          <span>Da richiamare</span>
        </div>
      </section>
    `;
  }).join("");

  return `<!doctype html>
  <html lang="it">
  <head>
    <meta charset="utf-8">
    <title>Giro consegne Dental G</title>
    <style>
      @page { size: A4; margin: 14mm; }
      body { font-family: Arial, Helvetica, sans-serif; color: #1f2937; margin: 0; }
      h1, h2, p { margin: 0; }
      .sheet { padding: 8px; }
      .hero { display: flex; justify-content: space-between; gap: 18px; margin-bottom: 20px; align-items: flex-start; }
      .hero h1 { font-size: 30px; margin-bottom: 6px; }
      .hero p { color: #4b5563; }
      .summary-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 10px; margin-bottom: 18px; }
      .summary-card { border: 1px solid #d7deea; border-radius: 14px; padding: 12px; background: #f8fbff; }
      .summary-card span { display: block; color: #526277; font-size: 12px; text-transform: uppercase; letter-spacing: 0.06em; }
      .summary-card strong { display: block; margin-top: 6px; font-size: 20px; }
      .print-route { margin-bottom: 20px; border: 1px solid #d7deea; border-radius: 16px; padding: 14px; background: #fbfdff; }
      .print-route h2 { margin-bottom: 8px; font-size: 20px; color: #1f4b7a; }
      .route-origin { margin-bottom: 12px; color: #4b5563; font-weight: 700; }
      .route-stops { display: grid; gap: 8px; }
      .route-stop { display: grid; grid-template-columns: 32px 1fr; gap: 10px; align-items: start; }
      .route-stop-index { width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; border-radius: 999px; background: #1f4b7a; color: #fff; font-weight: 700; font-size: 13px; }
      .route-stop p { margin-top: 2px; color: #4b5563; font-size: 13px; }
      .print-stop { break-inside: avoid; page-break-inside: avoid; margin-bottom: 18px; border: 1px solid #d7deea; border-radius: 16px; padding: 14px; }
      .print-stop-head { display: flex; justify-content: space-between; gap: 14px; margin-bottom: 12px; }
      .print-stop-head h2 { font-size: 22px; margin-bottom: 4px; }
      .print-stop-head p { color: #4b5563; }
      .print-stop-city { margin-top: 4px; font-weight: 700; color: #1f4b7a; }
      .print-stop-badges { display: flex; gap: 8px; flex-wrap: wrap; align-items: flex-start; justify-content: flex-end; }
      .print-stop-badges span { padding: 6px 10px; border-radius: 999px; background: #edf4ff; color: #1f4b7a; font-size: 12px; font-weight: 700; }
      table { width: 100%; border-collapse: collapse; }
      th, td { border-bottom: 1px solid #e5ebf3; padding: 10px 8px; text-align: left; vertical-align: top; font-size: 13px; }
      th { color: #1f4b7a; text-transform: uppercase; font-size: 11px; letter-spacing: 0.05em; }
      .check-box { display: inline-block; width: 16px; height: 16px; border: 2px solid #9aa8bb; border-radius: 4px; }
      .print-stop-footer { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-top: 12px; color: #526277; font-size: 12px; }
      .print-stop-footer span { border: 1px dashed #cad4e2; border-radius: 10px; padding: 8px; min-height: 32px; }
      @media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
    </style>
  </head>
  <body>
    <div class="sheet">
      <section class="hero">
        <div>
          <h1>Giro consegne Dental G</h1>
          <p>Foglio operativo per il fattorino</p>
          <p>Partenza: ${escapeHtml(state.deliveryOrigin || "Non impostata")}</p>
        </div>
        <div>
          <p>Data stampa: ${escapeHtml(new Date().toLocaleDateString("it-IT"))}</p>
          <p>Tappe: ${escapeHtml(String(stops.length))}</p>
        </div>
      </section>
      <section class="summary-grid">
        <article class="summary-card"><span>Tappe</span><strong>${escapeHtml(String(stops.length))}</strong></article>
        <article class="summary-card"><span>Consegne</span><strong>${escapeHtml(String(summary.deliveries))}</strong></article>
        <article class="summary-card"><span>Ritiri prova</span><strong>${escapeHtml(String(summary.trials))}</strong></article>
        <article class="summary-card"><span>Ingressi</span><strong>${escapeHtml(String(summary.pickups))}</strong></article>
        <article class="summary-card"><span>Commissioni</span><strong>${escapeHtml(String(summary.manuals))}</strong></article>
      </section>
      ${routeMarkup}
      ${stopMarkup}
    </div>
  </body>
  </html>`;
}

function renderDeliveryAddressBook() {
  if (!deliveryAddressBookList) return;
  const search = normalizeText(deliveryAddressSearchInput?.value);
  const filtered = state.addressBook.filter((record) => {
    if (!search) return true;
    return normalizeText(`${record.clientName} ${record.address} ${record.city} ${record.notes}`).includes(search);
  });

  deliveryAddressBookList.innerHTML = filtered.length ? filtered.map((record) => `
    <details class="delivery-address-item ${record.verified ? "" : "unverified"}">
      <summary class="delivery-address-summary">
        <div class="delivery-address-summary-copy">
          <strong>${escapeHtml(record.clientName)}</strong>
          <span>${escapeHtml([record.city, record.province].filter(Boolean).join(" - ") || "Scheda anagrafica da completare")}</span>
        </div>
        <div class="delivery-address-summary-meta">
          <button class="delivery-address-edit-button" type="button" data-address-action="edit" data-address-client="${escapeHtml(record.clientName)}">Modifica</button>
          <button class="delivery-address-delete-button" type="button" data-address-action="delete" data-address-client="${escapeHtml(record.clientName)}">Elimina</button>
          <span class="delivery-address-chevron" aria-hidden="true">+</span>
        </div>
      </summary>
      <div class="delivery-address-body">
        <div class="delivery-address-grid">
          <span>${escapeHtml(record.address || "Indirizzo non disponibile")}</span>
          <span>${escapeHtml([record.city, record.province].filter(Boolean).join(" - ") || "-")}</span>
        </div>
        ${record.notes ? `<p class="delivery-address-note">${escapeHtml(record.notes)}</p>` : ""}
      </div>
    </details>
  `).join("") : `<div class="delivery-empty">Nessuno studio corrisponde alla ricerca anagrafica.</div>`;
}

function handleAddressBookClick(event) {
  const trigger = event.target.closest("[data-address-action][data-address-client]");
  if (!trigger) return;
  event.preventDefault();
  event.stopPropagation();
  const clientName = trigger.getAttribute("data-address-client") || "";
  const action = trigger.getAttribute("data-address-action") || "edit";
  if (action === "delete") {
    deleteAddressRecord(clientName);
    return;
  }
  openAddressModal(clientName);
}

function openAddressModal(clientName) {
  const record = getAddressRecordByClient(clientName);
  if (!record) return;
  activeAddressClientName = record.clientName;
  activeAddressOriginalKey = normalizeText(record.clientName);
  addressModalTitle.textContent = `Modifica indirizzo - ${record.clientName}`;
  if (addressModalSaveButton) addressModalSaveButton.textContent = "Salva indirizzo";
  addressClientInput.value = record.clientName;
  addressClientInput.readOnly = true;
  addressStreetInput.value = record.address || "";
  addressCityInput.value = record.city || "";
  addressProvinceInput.value = record.province || "";
  addressNotesInput.value = record.notes || "";
  addressModal.hidden = false;
  addressModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function openNewAddressModal() {
  activeAddressClientName = "";
  activeAddressOriginalKey = "";
  addressModalTitle.textContent = "Aggiungi medico o studio";
  if (addressModalSaveButton) addressModalSaveButton.textContent = "Aggiungi medico";
  addressForm.reset();
  addressClientInput.readOnly = false;
  addressModal.hidden = false;
  addressModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  addressClientInput.focus();
}

function closeAddressModal() {
  activeAddressClientName = "";
  activeAddressOriginalKey = "";
  addressModal.hidden = true;
  addressModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
  addressForm.reset();
  addressClientInput.readOnly = false;
}

function handleAddressSave(event) {
  event.preventDefault();
  const nextClientName = normalizeImportedText(addressClientInput.value);
  if (!nextClientName) {
    window.alert("Inserisci almeno il nome del medico o dello studio.");
    addressClientInput.focus();
    return;
  }
  const nextAddress = normalizeImportedText(addressStreetInput.value);
  const nextCity = normalizeImportedText(addressCityInput.value);
  const nextProvince = normalizeImportedText(addressProvinceInput.value).slice(0, 2).toUpperCase();
  const nextNotes = normalizeImportedText(addressNotesInput.value);
  const nextRecord = normalizeAddressRecord({
    clientName: nextClientName,
    address: nextAddress,
    city: nextCity,
    province: nextProvince,
    notes: nextNotes,
    verified: !!nextAddress
  });
  const nextKey = normalizeText(nextClientName);

  if (!activeAddressOriginalKey) {
    const alreadyExists = state.addressBook.some((record) => normalizeText(record.clientName) === nextKey);
    if (alreadyExists) {
      window.alert("Esiste gia' una scheda con questo nome. Se vuoi, usa Modifica sulla scheda esistente.");
      addressClientInput.focus();
      return;
    }
    state.addressBook = [...state.addressBook, nextRecord].sort((a, b) => normalizeText(a.clientName).localeCompare(normalizeText(b.clientName)));
  } else {
    state.addressBook = state.addressBook.map((record) => {
      if (normalizeText(record.clientName) !== activeAddressOriginalKey) return record;
      return {
        ...record,
        ...nextRecord
      };
    }).sort((a, b) => normalizeText(a.clientName).localeCompare(normalizeText(b.clientName)));
  }
  persistState();
  renderDeliveryPlanner();
  closeAddressModal();
}

function deleteAddressRecord(clientName) {
  const key = normalizeText(clientName);
  if (!key) return;
  const record = state.addressBook.find((item) => normalizeText(item.clientName) === key);
  if (!record) return;
  const confirmed = window.confirm(`Vuoi eliminare la scheda di "${record.clientName}" dalla rubrica?`);
  if (!confirmed) return;
  state.addressBook = state.addressBook.filter((item) => normalizeText(item.clientName) !== key);
  persistState();
  renderDeliveryPlanner();
}

function resetManualTaskForm() {
  if (!manualTaskForm) return;
  manualTaskForm.reset();
  manualTaskIdHidden.value = "";
  if (manualTaskCreatedDateInput) manualTaskCreatedDateInput.value = formatDateForInput("");
  manualTaskKindInput.value = "Commissione";
  manualTaskUrgentInput.checked = false;
}

function resetManualTaskModalForm() {
  if (!manualTaskModalForm) return;
  manualTaskModalForm.reset();
  if (manualTaskModalIdHidden) manualTaskModalIdHidden.value = "";
  if (manualTaskModalCreatedDateInput) manualTaskModalCreatedDateInput.value = formatDateForInput("");
  if (manualTaskModalKindInput) manualTaskModalKindInput.value = "Commissione";
  if (manualTaskModalUrgentInput) manualTaskModalUrgentInput.checked = false;
}

function openManualTaskModal(taskId) {
  const task = state.manualTasks.find((item) => String(item.id) === String(taskId));
  if (!task || !manualTaskModal) return;
  manualTaskModalIdHidden.value = task.id;
  manualTaskModalTitleInput.value = task.title || "";
  manualTaskModalClientInput.value = task.client || "";
  if (manualTaskModalCreatedDateInput) manualTaskModalCreatedDateInput.value = formatDateForInput(task.scheduledAt);
  manualTaskModalAddressInput.value = task.address || "";
  manualTaskModalCityInput.value = task.city || "";
  manualTaskModalProvinceInput.value = task.province || "";
  manualTaskModalKindInput.value = task.kind || "Commissione";
  manualTaskModalTimeSlotInput.value = task.timeSlot || "";
  manualTaskModalNoteInput.value = task.note || "";
  manualTaskModalUrgentInput.checked = !!task.urgent;
  manualTaskModal.hidden = false;
  manualTaskModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  manualTaskModalTitleInput.focus();
}

function closeManualTaskModal() {
  if (!manualTaskModal) return;
  manualTaskModal.hidden = true;
  manualTaskModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
  resetManualTaskModalForm();
}

function initializeGoogleMapsAddressAutocomplete() {
  if (!GOOGLE_MAPS_API_KEY) {
    if (manualTaskAddressHint) {
      manualTaskAddressHint.textContent = "Suggerimento indirizzo Google non attivo: inserisci la chiave API nel progetto per vedere i suggerimenti automatici.";
    }
    console.info("Google Maps autocomplete disattivato: inserisci la API key nel meta tag google-maps-api-key.");
    return;
  }

  loadGoogleMapsPlacesScript()
    .then(() => {
      if (manualTaskAddressHint) {
        manualTaskAddressHint.textContent = "Suggerimento indirizzo attivo: scegli una proposta per compilare anche citta' e provincia.";
      }
      setupAddressAutocomplete(manualTaskAddressInput, {
        cityInput: manualTaskCityInput,
        provinceInput: manualTaskProvinceInput
      });
    })
    .catch((error) => {
      if (manualTaskAddressHint) {
        manualTaskAddressHint.textContent = "Suggerimento indirizzo non disponibile: controlla la configurazione Google Maps Places.";
      }
      console.error("Impossibile inizializzare Google Maps Places Autocomplete.", error);
    });
}

function loadGoogleMapsPlacesScript() {
  if (window.google?.maps?.places) {
    return Promise.resolve();
  }
  if (googlePlacesLoadPromise) {
    return googlePlacesLoadPromise;
  }

  googlePlacesLoadPromise = new Promise((resolve, reject) => {
    const existingScript = document.getElementById(GOOGLE_MAPS_AUTOCOMPLETE_SCRIPT_ID);
    if (existingScript) {
      existingScript.addEventListener("load", () => resolve(), { once: true });
      existingScript.addEventListener("error", () => reject(new Error("Caricamento Google Maps Places fallito.")), { once: true });
      return;
    }

    window[GOOGLE_MAPS_AUTOCOMPLETE_CALLBACK] = () => resolve();
    const script = document.createElement("script");
    script.id = GOOGLE_MAPS_AUTOCOMPLETE_SCRIPT_ID;
    script.async = true;
    script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(GOOGLE_MAPS_API_KEY)}&loading=async&libraries=places&callback=${GOOGLE_MAPS_AUTOCOMPLETE_CALLBACK}`;
    script.onerror = () => reject(new Error("Caricamento Google Maps Places fallito."));
    document.head.appendChild(script);
  });

  return googlePlacesLoadPromise;
}

function setupAddressAutocomplete(addressInput, { cityInput, provinceInput }) {
  if (!addressInput || !window.google?.maps?.places) return;
  if (addressInput.dataset.autocompleteReady === "true") return;

  const autocomplete = new google.maps.places.Autocomplete(addressInput, {
    fields: ["address_components", "formatted_address"],
    types: ["address"],
    componentRestrictions: { country: "it" }
  });

  autocomplete.addListener("place_changed", () => {
    const place = autocomplete.getPlace();
    if (!place) return;
    if (place.formatted_address) {
      addressInput.value = place.formatted_address;
    }
    applyPlaceAddressComponents(place.address_components || [], cityInput, provinceInput);
  });

  addressInput.dataset.autocompleteReady = "true";
}

function applyPlaceAddressComponents(components, cityInput, provinceInput) {
  if (!Array.isArray(components) || !components.length) return;
  let city = "";
  let province = "";

  components.forEach((component) => {
    const types = Array.isArray(component.types) ? component.types : [];
    if (!city && (types.includes("locality") || types.includes("postal_town") || types.includes("administrative_area_level_3"))) {
      city = component.long_name || "";
    }
    if (!province && types.includes("administrative_area_level_2")) {
      province = component.short_name || "";
    }
  });

  if (cityInput && city) cityInput.value = city;
  if (provinceInput && province) provinceInput.value = province.toUpperCase().slice(0, 2);
}

function formatDateTimeForHumans(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toLocaleString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatDateForHumans(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toLocaleDateString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  });
}

function formatDateForInput(value) {
  const parsed = value ? new Date(value) : new Date();
  if (Number.isNaN(parsed.getTime())) return "";
  const year = String(parsed.getFullYear());
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function combineManualTaskDateTime(dateValue, fallbackValue = "") {
  const safeDate = clean(dateValue);
  if (safeDate) return new Date(`${safeDate}T00:00`).toISOString();
  return clean(fallbackValue) || new Date().toISOString();
}

function buildManualTaskFromValues(values) {
  const existingTask = state.manualTasks.find((task) => String(task.id) === String(values.id));
  return normalizeManualTaskRecord({
    id: values.id || `manual-${Date.now()}`,
    title: values.title,
    client: values.client,
    scheduledAt: combineManualTaskDateTime(
      values.scheduledDate,
      existingTask?.scheduledAt || ""
    ),
    address: values.address,
    city: values.city,
    province: values.province,
    kind: values.kind,
    timeSlot: values.timeSlot,
    note: values.note,
    urgent: values.urgent
  });
}

function saveManualTaskRecord(normalizedTask) {
  const existingIndex = state.manualTasks.findIndex((task) => String(task.id) === String(normalizedTask.id));
  if (existingIndex >= 0) {
    state.manualTasks.splice(existingIndex, 1, normalizedTask);
  } else {
    state.manualTasks.unshift(normalizedTask);
  }

  persistState();
  renderDeliveryPlanner();
}

function handleManualTaskSubmit(event) {
  event.preventDefault();
  const normalizedTask = buildManualTaskFromValues({
    id: manualTaskIdHidden.value,
    title: manualTaskTitleInput.value,
    client: manualTaskClientInput.value,
    scheduledDate: manualTaskCreatedDateInput?.value,
    address: manualTaskAddressInput.value,
    city: manualTaskCityInput.value,
    province: manualTaskProvinceInput.value,
    kind: manualTaskKindInput.value,
    timeSlot: manualTaskTimeSlotInput.value,
    note: manualTaskNoteInput.value,
    urgent: manualTaskUrgentInput.checked
  });
  saveManualTaskRecord(normalizedTask);
  resetManualTaskForm();
}

function handleManualTaskModalSubmit(event) {
  event.preventDefault();
  const normalizedTask = buildManualTaskFromValues({
    id: manualTaskModalIdHidden.value,
    title: manualTaskModalTitleInput.value,
    client: manualTaskModalClientInput.value,
    scheduledDate: manualTaskModalCreatedDateInput?.value,
    address: manualTaskModalAddressInput.value,
    city: manualTaskModalCityInput.value,
    province: manualTaskModalProvinceInput.value,
    kind: manualTaskModalKindInput.value,
    timeSlot: manualTaskModalTimeSlotInput.value,
    note: manualTaskModalNoteInput.value,
    urgent: manualTaskModalUrgentInput.checked
  });
  saveManualTaskRecord(normalizedTask);
  closeManualTaskModal();
}

function populateManualTaskForm(taskId) {
  const task = state.manualTasks.find((item) => String(item.id) === String(taskId));
  if (!task) return;
  manualTaskIdHidden.value = task.id;
  manualTaskTitleInput.value = task.title || "";
  manualTaskClientInput.value = task.client || "";
  if (manualTaskCreatedDateInput) manualTaskCreatedDateInput.value = formatDateForInput(task.scheduledAt);
  manualTaskAddressInput.value = task.address || "";
  manualTaskCityInput.value = task.city || "";
  manualTaskProvinceInput.value = task.province || "";
  manualTaskKindInput.value = task.kind || "Commissione";
  manualTaskTimeSlotInput.value = task.timeSlot || "";
  manualTaskNoteInput.value = task.note || "";
  manualTaskUrgentInput.checked = !!task.urgent;
}

function handleManualTaskListClick(event) {
  const actionTrigger = event.target.closest("[data-manual-task-action]");
  if (!actionTrigger) return;
  const taskId = actionTrigger.getAttribute("data-manual-task-id") || "";
  const action = actionTrigger.getAttribute("data-manual-task-action");
  if (action === "edit") {
    openManualTaskModal(taskId);
    return;
  }
  if (action === "delete") {
    state.manualTasks = state.manualTasks.filter((task) => String(task.id) !== String(taskId));
    delete state.deliverySelection[String(taskId)];
    state.deliveryConfirmedJobIds = state.deliveryConfirmedJobIds.filter((id) => String(id) !== String(taskId));
    persistState();
    resetManualTaskForm();
    renderDeliveryPlanner();
  }
}

function renderManualTaskList() {
  if (!manualTaskList) return;
  if (!state.manualTasks.length) {
    manualTaskList.innerHTML = `<div class="delivery-empty">Nessuna commissione manuale inserita. Aggiungile qui per farle entrare nel giro insieme ai lavori.</div>`;
    return;
  }

  manualTaskList.innerHTML = state.manualTasks.map((task) => `
    <article class="manual-task-card ${task.urgent ? "urgent" : ""}">
      <div>
        <strong>${escapeHtml(task.title || "Commissione")}</strong>
        <p>${escapeHtml(task.client || "Destinatario non indicato")}</p>
        ${task.scheduledAt ? `<p class="delivery-job-extra">Data commissione: ${escapeHtml(formatDateForHumans(task.scheduledAt))}</p>` : ""}
        ${task.address ? `<p>${escapeHtml(task.address)}</p>` : ""}
        <div class="manual-task-tags">
          <span>${escapeHtml(task.kind || "Commissione")}</span>
          ${task.timeSlot ? `<span>${escapeHtml(task.timeSlot)}</span>` : ""}
          ${task.urgent ? `<span>Urgente</span>` : ""}
        </div>
        ${task.note ? `<p class="delivery-job-extra">Note: ${escapeHtml(task.note)}</p>` : ""}
      </div>
      <div class="manual-task-actions">
        <button class="secondary" type="button" data-manual-task-action="edit" data-manual-task-id="${escapeHtml(task.id)}">Modifica</button>
        <button class="secondary" type="button" data-manual-task-action="delete" data-manual-task-id="${escapeHtml(task.id)}">Elimina</button>
      </div>
    </article>
  `).join("");
}

jobCodeInput.addEventListener("input", () => {
  jobCodeInput.value = clean(jobCodeInput.value).replace(/\D/g, "").slice(0, 8);
});

function initializeParentFrameReporting(frameName) {
  if (window.top === window.self) {
    return;
  }

  const reportHeight = () => {
    const height = Math.max(
      document.body.scrollHeight,
      document.body.offsetHeight,
      document.documentElement.scrollHeight,
      document.documentElement.offsetHeight
    );

    window.parent.postMessage({ type: "dentalg:frameHeight", frameName, height }, "*");
  };

  const scheduleReport = () => {
    reportHeight();
    [120, 320, 700, 1400].forEach((delay) => window.setTimeout(reportHeight, delay));
  };

  window.addEventListener("load", scheduleReport);
  window.addEventListener("resize", scheduleReport);

  const observer = new MutationObserver(scheduleReport);
  observer.observe(document.body, { childList: true, subtree: true, attributes: true, characterData: true });

  scheduleReport();
}

window.printDeliveryPlan = printDeliveryPlan;
window.openDeliveryRouteInMaps = openDeliveryRouteInMaps;
