import templateMapping from "../config/templateMapping.js";
import {
  calculateGrandTotal,
  exportedTripRowCount,
  generateClaimWorkbook,
  safeFileName
} from "./lib/xlsxTemplatePatcher.js";

const STORAGE_KEYS = {
  draft: "travelClaimDraft.v1",
  history: "travelClaimHistory.v1",
  settings: "travelClaimSettings.v1"
};

const emptyTrip = () => ({
  id: crypto.randomUUID(),
  date: "",
  from: "",
  to: "",
  timeFrom: "",
  timeTo: "",
  client: "",
  projectCode: "",
  transportType: "MTR",
  amount: "",
  deduction: "",
  remarks: ""
});

const emptyClaim = () => ({
  id: crypto.randomUUID(),
  employeeName: "",
  employeeNumber: "",
  address: "",
  weekStart: "",
  weekEnding: "",
  claimDate: new Date().toISOString().slice(0, 10),
  remarks: "",
  trips: [emptyTrip()],
  updatedAt: new Date().toISOString()
});

let state = {
  view: "home",
  claim: readJson(STORAGE_KEYS.draft, emptyClaim()),
  history: readJson(STORAGE_KEYS.history, []),
  settings: readJson(STORAGE_KEYS.settings, { suggestions: { from: [], to: [], client: [], projectCode: [] } }),
  editingTripId: null,
  notice: ""
};

function readJson(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function setState(patch) {
  state = { ...state, ...patch };
  saveDraft();
  render();
}

function saveDraft() {
  state.claim.updatedAt = new Date().toISOString();
  saveJson(STORAGE_KEYS.draft, state.claim);
}

function persistClaimInput() {
  saveDraft();
  updateLiveTotals();
}

function saveHistory(claim) {
  const snapshot = { ...claim, id: crypto.randomUUID(), savedAt: new Date().toISOString() };
  const history = [snapshot, ...state.history.filter((item) => item.id !== claim.id)].slice(0, 50);
  state.history = history;
  saveJson(STORAGE_KEYS.history, history);
}

function updateSuggestions() {
  const suggestions = { ...state.settings.suggestions };
  for (const trip of state.claim.trips) {
    for (const key of ["from", "to", "client", "projectCode"]) {
      const value = (trip[key] || "").trim();
      if (value && !suggestions[key].includes(value)) suggestions[key] = [value, ...suggestions[key]].slice(0, 30);
    }
  }
  state.settings = { ...state.settings, suggestions };
  saveJson(STORAGE_KEYS.settings, state.settings);
}

function money(value) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? n : 0;
}

function tripNet(trip) {
  return money(trip.amount) - money(trip.deduction);
}

function updateLiveTotals() {
  document.querySelectorAll("[data-grand-total]").forEach((node) => {
    node.textContent = `Grand Total HK$ ${calculateGrandTotal(state.claim).toFixed(2)}`;
  });
}

function updateTripSummary(trip) {
  const label = document.querySelector(`[data-trip-label="${trip.id}"]`);
  const total = document.querySelector(`[data-trip-total="${trip.id}"]`);
  const index = state.claim.trips.findIndex((item) => item.id === trip.id);
  if (label) label.textContent = `${index + 1}. ${trip.from || "From"} -> ${trip.to || "To"}`;
  if (total) total.textContent = `HK$ ${tripNet(trip).toFixed(2)}`;
  updateLiveTotals();
}

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "class") node.className = value;
    else if (key === "dataset") Object.assign(node.dataset, value);
    else if (key.startsWith("on")) node.addEventListener(key.slice(2).toLowerCase(), value);
    else if (value !== false && value !== null && value !== undefined) node.setAttribute(key, value === true ? "" : value);
  }
  for (const child of Array.isArray(children) ? children : [children]) {
    if (child === null || child === undefined) continue;
    node.append(child.nodeType ? child : document.createTextNode(String(child)));
  }
  return node;
}

function field(label, key, type = "text", options = {}) {
  const input = el("input", {
    id: key,
    type,
    value: state.claim[key] || "",
    inputmode: options.inputmode,
    list: options.list,
    placeholder: options.placeholder || "",
    oninput: (event) => {
      state.claim[key] = event.target.value;
      persistClaimInput();
    }
  });
  return el("label", { class: "field" }, [el("span", {}, label), input]);
}

function tripField(trip, label, key, type = "text", options = {}) {
  const input = el("input", {
    type,
    value: trip[key] || "",
    inputmode: options.inputmode,
    list: options.list,
    placeholder: options.placeholder || "",
    oninput: (event) => {
      trip[key] = event.target.value;
      persistClaimInput();
      updateTripSummary(trip);
    }
  });
  return el("label", { class: "field" }, [el("span", {}, label), input]);
}

function datalist(id, values) {
  return el("datalist", { id }, values.map((value) => el("option", { value })));
}

function topBar(title) {
  return el("header", { class: "topbar" }, [
    state.view === "home" ? el("div", { class: "brand" }, "Travelling Claim") : button("Back", () => setState({ view: "home", editingTripId: null }), "secondary"),
    el("h1", {}, title)
  ]);
}

function button(text, onClick, variant = "primary", attrs = {}) {
  return el("button", { class: `btn ${variant}`, onclick: onClick, ...attrs }, text);
}

function renderHome() {
  return [
    topBar(""),
    el("main", { class: "screen home" }, [
      el("section", { class: "home-actions" }, [
        button("New Claim", () => setState({ claim: emptyClaim(), view: "claim", notice: "New draft started." })),
        button("Draft", () => setState({ view: "claim" }), "secondary"),
        button("History", () => setState({ view: "history" }), "secondary"),
        button("Settings", () => setState({ view: "settings" }), "secondary")
      ]),
      el("section", { class: "status-panel" }, [
        el("strong", {}, `${state.claim.trips.length}/${templateMapping.maxTrips} trip rows`),
        el("span", {}, `Current total HK$ ${calculateGrandTotal(state.claim).toFixed(2)}`),
        el("span", {}, "All claim data stays in this browser.")
      ])
    ])
  ];
}

function renderClaim() {
  const overLimit = state.claim.trips.length > templateMapping.maxTrips;
  return [
    topBar("Claim"),
    el("main", { class: "screen" }, [
      notice(),
      el("section", { class: "panel" }, [
        el("h2", {}, "Basic Details"),
        field("Employee Name", "employeeName"),
        field("Employee Number", "employeeNumber"),
        field("Address", "address"),
        el("div", { class: "two" }, [field("Week Start", "weekStart", "date"), field("Week Ending", "weekEnding", "date")]),
        field("Claim Date", "claimDate", "date"),
        field("Remarks", "remarks")
      ]),
      el("section", { class: "panel" }, [
        el("div", { class: "section-title" }, [
          el("h2", {}, "Trips"),
          el("span", { class: overLimit ? "limit bad" : "limit" }, `${state.claim.trips.length}/${templateMapping.maxTrips}`)
        ]),
        overLimit ? el("p", { class: "error" }, `Template only has ${templateMapping.maxTrips} trip rows. Delete or split trips before export.`) : null,
        ...state.claim.trips.map((trip, index) => renderTripCard(trip, index)),
        button("Add Trip", () => addTrip(), "secondary", { disabled: state.claim.trips.length >= templateMapping.maxTrips + 1 })
      ]),
      el("section", { class: "sticky-total" }, [
        el("span", { dataset: { grandTotal: "true" } }, `Grand Total HK$ ${calculateGrandTotal(state.claim).toFixed(2)}`),
        button("Preview", () => setState({ view: "preview" }), overLimit ? "disabled" : "primary", { disabled: overLimit })
      ]),
      renderSuggestionLists()
    ])
  ];
}

function renderTripCard(trip, index) {
  const expanded = state.editingTripId === trip.id;
  return el("article", { class: "trip-card" }, [
    el("button", { class: "trip-summary", onclick: () => setState({ editingTripId: expanded ? null : trip.id }) }, [
      el("span", { dataset: { tripLabel: trip.id } }, `${index + 1}. ${trip.from || "From"} -> ${trip.to || "To"}`),
      el("strong", { dataset: { tripTotal: trip.id } }, `HK$ ${tripNet(trip).toFixed(2)}`)
    ]),
    expanded
      ? el("div", { class: "trip-editor" }, [
          el("div", { class: "two" }, [tripField(trip, "Date", "date", "date"), tripField(trip, "Amount", "amount", "number", { inputmode: "decimal" })]),
          el("div", { class: "two" }, [tripField(trip, "From", "from", "text", { list: "from-list" }), tripField(trip, "To", "to", "text", { list: "to-list" })]),
          el("div", { class: "two" }, [tripField(trip, "Time From", "timeFrom", "time"), tripField(trip, "Time To", "timeTo", "time")]),
          tripField(trip, "Client", "client", "text", { list: "client-list" }),
          tripField(trip, "Project Code", "projectCode", "text", { list: "project-list" }),
          el("label", { class: "field" }, [
            el("span", {}, "Transport Type"),
            el("select", {
              onchange: (event) => {
                trip.transportType = event.target.value;
                persistClaimInput();
                updateTripSummary(trip);
              }
            }, ["MTR", "Bus", "Van", "Taxi", "Other"].map((item) => el("option", { value: item, selected: trip.transportType === item }, item)))
          ]),
          tripField(trip, "Deduction", "deduction", "number", { inputmode: "decimal" }),
          tripField(trip, "Remarks", "remarks"),
          el("div", { class: "trip-actions" }, [
            button("Duplicate", () => duplicateTrip(trip.id), "secondary"),
            button("Up", () => moveTrip(index, -1), "secondary", { disabled: index === 0 }),
            button("Down", () => moveTrip(index, 1), "secondary", { disabled: index === state.claim.trips.length - 1 }),
            button("Delete", () => deleteTrip(trip.id), "danger")
          ])
        ])
      : null
  ]);
}

function renderPreview() {
  const rows = exportedTripRowCount(state.claim, templateMapping);
  return [
    topBar("Preview"),
    el("main", { class: "screen" }, [
      el("section", { class: "panel" }, [
        el("h2", {}, "Basic Details"),
        previewLine("Employee", `${state.claim.employeeName || "-"} ${state.claim.employeeNumber ? `(${state.claim.employeeNumber})` : ""}`),
        previewLine("Address", state.claim.address || "-"),
        previewLine("Week Ending", state.claim.weekEnding || "-"),
        previewLine("Claim Date", state.claim.claimDate || "-"),
        previewLine("Remarks", state.claim.remarks || "-")
      ]),
      el("section", { class: "panel" }, [
        el("h2", {}, "Trips"),
        ...state.claim.trips.map((trip, index) =>
          el("div", { class: "preview-trip" }, [
            el("strong", {}, `${index + 1}. ${trip.date || "-"} ${trip.from || "-"} → ${trip.to || "-"}`),
            el("span", {}, `${trip.transportType} HK$ ${money(trip.amount).toFixed(2)} | Deduction HK$ ${money(trip.deduction).toFixed(2)}`),
            el("span", {}, `${trip.client || "-"} / ${trip.projectCode || "-"}`)
          ])
        )
      ]),
      el("section", { class: "panel totals" }, [
        previewLine("Excel rows to fill", String(rows)),
        previewLine("Grand Total", `HK$ ${calculateGrandTotal(state.claim).toFixed(2)}`)
      ]),
      button("Confirm and Generate Excel", generateExcel)
    ])
  ];
}

function previewLine(label, value) {
  return el("div", { class: "preview-line" }, [el("span", {}, label), el("strong", {}, value)]);
}

function renderHistory() {
  return [
    topBar("History"),
    el("main", { class: "screen" }, [
      state.history.length ? null : el("p", { class: "empty" }, "No saved claims yet."),
      ...state.history.map((claim) =>
        el("article", { class: "history-item" }, [
          el("strong", {}, claim.employeeName || "Unnamed Claim"),
          el("span", {}, `${claim.claimDate || "-"} | ${claim.trips.length} trips | HK$ ${calculateGrandTotal(claim).toFixed(2)}`),
          el("div", { class: "trip-actions" }, [
            button("Open", () => setState({ claim: structuredClone(claim), view: "claim" }), "secondary"),
            button("Duplicate Previous Claim", () => {
              const copy = structuredClone(claim);
              copy.id = crypto.randomUUID();
              copy.claimDate = new Date().toISOString().slice(0, 10);
              copy.trips = copy.trips.map((trip) => ({ ...trip, id: crypto.randomUUID() }));
              setState({ claim: copy, view: "claim", notice: "Previous claim duplicated." });
            })
          ])
        ])
      )
    ])
  ];
}

function renderSettings() {
  return [
    topBar("Settings"),
    el("main", { class: "screen" }, [
      el("section", { class: "panel" }, [
        el("h2", {}, "Local Data"),
        el("p", {}, "Drafts, history and suggestions are stored only in this browser."),
        button("Clear All Local Data", () => {
          localStorage.removeItem(STORAGE_KEYS.draft);
          localStorage.removeItem(STORAGE_KEYS.history);
          localStorage.removeItem(STORAGE_KEYS.settings);
          state = { ...state, claim: emptyClaim(), history: [], settings: { suggestions: { from: [], to: [], client: [], projectCode: [] } } };
          render();
        }, "danger")
      ]),
      el("section", { class: "panel" }, [
        el("h2", {}, "Template"),
        previewLine("Sheets", templateMapping.workbook.sheets.join(", ")),
        previewLine("Trip row capacity", String(templateMapping.maxTrips)),
        previewLine("Hash", templateMapping.templateSha256.slice(0, 16) + "...")
      ])
    ])
  ];
}

function renderSuggestionLists() {
  const s = state.settings.suggestions;
  return el("div", {}, [
    datalist("from-list", s.from),
    datalist("to-list", s.to),
    datalist("client-list", s.client),
    datalist("project-list", s.projectCode)
  ]);
}

function notice() {
  return state.notice ? el("p", { class: "notice" }, state.notice) : null;
}

function addTrip() {
  state.claim.trips.push(emptyTrip());
  setState({ claim: state.claim, editingTripId: state.claim.trips.at(-1).id, notice: "" });
}

function duplicateTrip(id) {
  const index = state.claim.trips.findIndex((trip) => trip.id === id);
  const copy = { ...state.claim.trips[index], id: crypto.randomUUID() };
  state.claim.trips.splice(index + 1, 0, copy);
  setState({ claim: state.claim, editingTripId: copy.id });
}

function deleteTrip(id) {
  state.claim.trips = state.claim.trips.filter((trip) => trip.id !== id);
  if (!state.claim.trips.length) state.claim.trips.push(emptyTrip());
  setState({ claim: state.claim, editingTripId: null });
}

function moveTrip(index, delta) {
  const next = index + delta;
  if (next < 0 || next >= state.claim.trips.length) return;
  const [trip] = state.claim.trips.splice(index, 1);
  state.claim.trips.splice(next, 0, trip);
  setState({ claim: state.claim });
}

async function generateExcel() {
  try {
    if (state.claim.trips.length > templateMapping.maxTrips) throw new Error(`Template supports ${templateMapping.maxTrips} trips only.`);
    updateSuggestions();
    const blob = await generateClaimWorkbook(state.claim, templateMapping);
    const url = URL.createObjectURL(blob);
    const link = el("a", { href: url, download: safeFileName(state.claim) });
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    saveHistory(state.claim);
    setState({ view: "history", notice: "Excel generated and claim saved to history." });
  } catch (error) {
    setState({ view: "claim", notice: error.message });
  }
}

function render() {
  const root = document.getElementById("app");
  root.innerHTML = "";
  const views = {
    home: renderHome,
    claim: renderClaim,
    preview: renderPreview,
    history: renderHistory,
    settings: renderSettings
  };
  root.append(...views[state.view]());
}

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("./service-worker.js").catch(() => {});
}

render();
