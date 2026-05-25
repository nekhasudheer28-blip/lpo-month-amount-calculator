const API_URL = window.LPO_API_URL || "http://localhost:8000";

const state = {
  file: null,
  jobNumbers: [],
  selectedJobs: new Set(),
  downloadUrl: "",
  highlightedDetails: [],
};

const fileInput             = document.getElementById("fileInput");
const fileName              = document.getElementById("fileName");
const uploadZone            = document.getElementById("uploadZone");
const taskAButton           = document.getElementById("taskAButton");
const taskBButton           = document.getElementById("taskBButton");
const downloadButton        = document.getElementById("downloadButton");
const downloadRow           = document.getElementById("downloadRow");
const errorBox              = document.getElementById("errorBox");
const summaryGrid           = document.getElementById("summaryGrid");
const selectorPanel         = document.getElementById("selectorPanel");
const searchInput           = document.getElementById("searchInput");
const jobList               = document.getElementById("jobList");
const selectedCount         = document.getElementById("selectedCount");
const highlightedPanel      = document.getElementById("highlightedPanel");
const highlightedList       = document.getElementById("highlightedList");
const unresolvedSection     = document.getElementById("unresolvedSection");
const unresolvedContent     = document.getElementById("unresolvedContent");
const taskBModalOverlay     = document.getElementById("taskBModalOverlay");
const modalBody             = document.getElementById("modalBody");
const modalCloseBtn         = document.getElementById("modalCloseBtn");
const modalDoneBtn          = document.getElementById("modalDoneBtn");

// ── File handling ────────────────────────────────────────────────
fileInput.addEventListener("change", () => setFile(fileInput.files[0]));

uploadZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadZone.classList.add("drag-over");
});

uploadZone.addEventListener("dragleave", () => uploadZone.classList.remove("drag-over"));

uploadZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadZone.classList.remove("drag-over");
  setFile(e.dataTransfer.files[0]);
});

function setFile(file) {
  if (!file) return;
  state.file = file;
  state.jobNumbers = [];
  state.selectedJobs.clear();
  state.highlightedDetails = [];

  fileName.textContent = file.name;
  taskAButton.disabled = false;
  taskBButton.disabled = true;

  summaryGrid.classList.add("hidden");
  highlightedPanel.classList.add("hidden");
  unresolvedSection.classList.add("hidden");
  selectorPanel.classList.add("hidden");
  downloadRow.classList.add("hidden");
  taskBModalOverlay.classList.add("hidden");
  hideError();
}

// ── Button wiring ────────────────────────────────────────────────
taskAButton.addEventListener("click", () => processWorkbook([], false));
taskBButton.addEventListener("click", () => processWorkbook([...state.selectedJobs], true));
searchInput.addEventListener("input", renderJobList);

modalCloseBtn.addEventListener("click", closeModal);
modalDoneBtn.addEventListener("click", closeModal);

// Close modal when clicking outside
taskBModalOverlay.addEventListener("click", (e) => {
  if (e.target === taskBModalOverlay) closeModal();
});

function closeModal() {
  taskBModalOverlay.classList.add("hidden");
  downloadRow.classList.remove("hidden");
  downloadRow.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ── Core fetch ───────────────────────────────────────────────────
async function processWorkbook(selectedJobs, isTaskB) {
  if (!state.file) { showError("Upload an Excel file first."); return; }

  setBusy(true);
  hideError();
  downloadRow.classList.add("hidden");
  taskBModalOverlay.classList.add("hidden");

  const formData = new FormData();
  formData.append("file", state.file);
  formData.append("selected_job_numbers", JSON.stringify(selectedJobs));

  try {
    const response = await fetch(`${API_URL}/process`, { method: "POST", body: formData });

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.detail || "Processing failed.");
    }

    const blob = await response.blob();
    if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = URL.createObjectURL(blob);

    const disposition = response.headers.get("content-disposition") || "";
    const fnMatch = disposition.match(/filename="([^"]+)"/);
    downloadButton.href = state.downloadUrl;
    downloadButton.download = fnMatch ? fnMatch[1] : "processed-lpo.xlsx";

    const jobsHeader    = response.headers.get("x-job-numbers");
    const summaryHeader = response.headers.get("x-summary");

    if (jobsHeader) state.jobNumbers = JSON.parse(jobsHeader);

    let summary = {};
    if (summaryHeader) {
      summary = JSON.parse(summaryHeader);
      renderSummary(summary);
    }

    renderJobList();
    selectorPanel.classList.remove("hidden");

    if (isTaskB && summary.taskBJobTotals && summary.taskBJobTotals.length > 0) {
      renderTaskBModal(summary.taskBJobTotals);
      // download revealed only after user closes the modal
    } else {
      downloadRow.classList.remove("hidden");
    }
  } catch (err) {
    showError(err.message);
  } finally {
    setBusy(false);
  }
}

// ── Summary (7 cards) ────────────────────────────────────────────
function renderSummary(summary) {
  state.highlightedDetails = summary.highlightedRowDetails || [];
  const hlCount = summary.taskAHighlightedRows || 0;

  const metrics = [
    { label: "Total Job No. Rows",           value: summary.totalRows            ?? 0 },
    { label: "Open Rows Processed",          value: summary.openRows             ?? 0 },
    { label: "Closed Rows",                  value: summary.closedRows           ?? 0 },
    { label: "Rows Highlighted",             value: hlCount,
      extra: hlCount > 0
        ? `<button class="expand-btn" onclick="toggleHighlightedPanel()">&#9660; View ${hlCount} rows</button>`
        : "" },
    { label: "Delivery Cells Concluded",     value: summary.deliveryCellsWritten ?? 0 },
    { label: "Payment Due Cells Concluded",  value: summary.paymentCellsWritten  ?? 0 },
    { label: "Totals Written",               value: summary.taskBTotalsWritten   ?? 0 },
  ];

  summaryGrid.innerHTML = metrics
    .map(({ label, value, extra }) => `
      <div class="metric">
        <span>${label}</span>
        <strong>${value}</strong>
        ${extra || ""}
      </div>`)
    .join("");

  summaryGrid.classList.remove("hidden");

  renderHighlightedPanel(state.highlightedDetails);
  renderUnresolvedRows(state.highlightedDetails);
}

// ── Highlighted panel (collapsible compact list) ─────────────────
function renderHighlightedPanel(details) {
  if (details.length === 0) {
    highlightedPanel.classList.add("hidden");
    return;
  }

  highlightedList.innerHTML = details
    .map((d) => {
      let badge;
      if (!d.deliveryOk && !d.paymentOk) {
        badge = '<span class="badge badge-red">Both &#10007;</span>';
      } else if (!d.deliveryOk) {
        badge = '<span class="badge badge-orange">Delivery &#10007;</span>';
      } else {
        badge = '<span class="badge badge-orange">Payment Due &#10007;</span>';
      }
      return `
        <div class="hl-row">
          <span class="hl-rownum">Row ${d.rowNumber}</span>
          <span class="hl-job">${escapeHtml(d.jobNumber)}</span>
          ${badge}
        </div>`;
    })
    .join("");
  // panel stays hidden until user clicks the toggle button
}

function toggleHighlightedPanel() {
  const isHidden = highlightedPanel.classList.toggle("hidden");
  const btn = summaryGrid.querySelector(".expand-btn");
  if (btn) {
    const count = state.highlightedDetails.length;
    btn.innerHTML = isHidden
      ? `&#9660; View ${count} rows`
      : `&#9650; Hide rows`;
  }
}

function closeHighlightedPanel() {
  highlightedPanel.classList.add("hidden");
  const btn = summaryGrid.querySelector(".expand-btn");
  if (btn) {
    const count = state.highlightedDetails.length;
    btn.innerHTML = `&#9660; View ${count} rows`;
  }
}

// ── Unresolved rows section (full table) ─────────────────────────
function renderUnresolvedRows(details) {
  if (details.length === 0) {
    unresolvedContent.innerHTML =
      '<div class="all-resolved">&#10003; All rows resolved — no manual review needed.</div>';
    unresolvedSection.classList.remove("hidden");
    return;
  }

  const rows = details
    .map((d) => {
      let issueText, issueCls;
      if (!d.deliveryOk && !d.paymentOk) {
        issueText = "Both unresolved";  issueCls = "issue-both";
      } else if (!d.deliveryOk) {
        issueText = "Delivery unresolved";  issueCls = "issue-one";
      } else {
        issueText = "Payment unresolved";  issueCls = "issue-one";
      }
      return `
        <tr>
          <td>${d.rowNumber}</td>
          <td class="job-number">${escapeHtml(d.jobNumber)}</td>
          <td class="term-cell">${escapeHtml(d.deliveryTerm || "—")}</td>
          <td class="term-cell">${escapeHtml(d.paymentTerm  || "—")}</td>
          <td><span class="issue-badge ${issueCls}">${issueText}</span></td>
        </tr>`;
    })
    .join("");

  unresolvedContent.innerHTML = `
    <div class="table-scroll">
      <table class="unresolved-table">
        <thead>
          <tr>
            <th>Row No.</th>
            <th>Job Number</th>
            <th>Delivery Term</th>
            <th>Payment Term</th>
            <th>Issue</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  unresolvedSection.classList.remove("hidden");
}

// ── Task B modal ─────────────────────────────────────────────────
function renderTaskBModal(jobTotals) {
  const rows = jobTotals
    .map((item) => `
      <tr>
        <td class="job-number">${escapeHtml(item.jobNumber)}</td>
        <td class="row-list">${(item.rows || []).join(", ")}</td>
        <td class="amount">${escapeHtml(item.total)}</td>
      </tr>`)
    .join("");

  modalBody.innerHTML = `
    <div class="table-scroll">
      <table class="totals-table">
        <thead>
          <tr>
            <th>Job Number</th>
            <th>Row(s)</th>
            <th>Total LPO Amount</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  taskBModalOverlay.classList.remove("hidden");
}

// ── Job list ─────────────────────────────────────────────────────
function renderJobList() {
  const query = searchInput.value.trim().toLowerCase();
  const filtered = query
    ? state.jobNumbers.filter((j) => j.toLowerCase().includes(query))
    : state.jobNumbers;

  jobList.innerHTML = filtered
    .map((job) => `
      <label class="job-option">
        <input type="checkbox" value="${escapeHtml(job)}" ${state.selectedJobs.has(job) ? "checked" : ""} />
        <span>${escapeHtml(job)}</span>
      </label>`)
    .join("");

  jobList.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", () => {
      input.checked ? state.selectedJobs.add(input.value) : state.selectedJobs.delete(input.value);
      updateSelectedCount();
    });
  });

  updateSelectedCount();
}

function updateSelectedCount() {
  const count = state.selectedJobs.size;
  selectedCount.textContent = `${count} selected`;
  taskBButton.disabled = !state.file || count === 0;
}

// ── Busy state ───────────────────────────────────────────────────
function setBusy(isBusy) {
  taskAButton.disabled = isBusy || !state.file;
  taskBButton.disabled = isBusy || !state.file || state.selectedJobs.size === 0;
  taskAButton.classList.toggle("busy", isBusy);
  taskBButton.classList.toggle("busy", isBusy);
}

// ── Error helpers ────────────────────────────────────────────────
function showError(msg) {
  errorBox.textContent = msg;
  errorBox.classList.remove("hidden");
}
function hideError() {
  errorBox.textContent = "";
  errorBox.classList.add("hidden");
}

// ── XSS guard ────────────────────────────────────────────────────
function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
