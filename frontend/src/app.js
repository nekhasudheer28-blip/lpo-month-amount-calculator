const API_URL = window.LPO_API_URL || "http://localhost:8000";

const state = {
  file: null,
  jobNumbers: [],
  selectedJobs: new Set(),
  downloadUrl: "",
};

const fileInput      = document.getElementById("fileInput");
const fileName       = document.getElementById("fileName");
const uploadZone     = document.getElementById("uploadZone");
const taskAButton    = document.getElementById("taskAButton");
const taskBButton    = document.getElementById("taskBButton");
const downloadButton = document.getElementById("downloadButton");
const downloadRow    = document.getElementById("downloadRow");
const errorBox       = document.getElementById("errorBox");
const summaryGrid    = document.getElementById("summaryGrid");
const selectorPanel  = document.getElementById("selectorPanel");
const searchInput    = document.getElementById("searchInput");
const jobList        = document.getElementById("jobList");
const selectedCount  = document.getElementById("selectedCount");
const taskBResults   = document.getElementById("taskBResults");
const taskBTable     = document.getElementById("taskBTable");

fileInput.addEventListener("change", () => setFile(fileInput.files[0]));

uploadZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadZone.style.borderColor = "#1a56db";
});

uploadZone.addEventListener("dragleave", () => {
  uploadZone.style.borderColor = "";
});

uploadZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadZone.style.borderColor = "";
  setFile(e.dataTransfer.files[0]);
});

taskAButton.addEventListener("click", () => processWorkbook([], false));
taskBButton.addEventListener("click", () => processWorkbook([...state.selectedJobs], true));
searchInput.addEventListener("input", renderJobList);

function setFile(file) {
  if (!file) return;

  state.file = file;
  state.jobNumbers = [];
  state.selectedJobs.clear();

  fileName.textContent = file.name;
  taskAButton.disabled = false;
  taskBButton.disabled = true;

  selectorPanel.classList.add("hidden");
  summaryGrid.classList.add("hidden");
  downloadRow.classList.add("hidden");
  taskBResults.classList.add("hidden");

  hideError();
}

async function processWorkbook(selectedJobs, isTaskB) {
  if (!state.file) {
    showError("Upload an Excel file first.");
    return;
  }

  setBusy(true);
  hideError();

  const formData = new FormData();
  formData.append("file", state.file);
  formData.append("selected_job_numbers", JSON.stringify(selectedJobs));

  try {
    const response = await fetch(`${API_URL}/process`, {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.detail || "Processing failed.");
    }

    const blob = await response.blob();

    if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = URL.createObjectURL(blob);

    const disposition = response.headers.get("content-disposition") || "";
    const filenameMatch = disposition.match(/filename="([^"]+)"/);
    const downloadName = filenameMatch ? filenameMatch[1] : "processed-lpo.xlsx";

    downloadButton.href = state.downloadUrl;
    downloadButton.download = downloadName;
    downloadRow.classList.remove("hidden");

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
      renderTaskBResults(summary.taskBJobTotals);
    } else {
      taskBResults.classList.add("hidden");
    }
  } catch (error) {
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

function renderSummary(summary) {
  const metrics = [
    ["Open rows processed", summary.taskAProcessedOpenRows],
    ["Rows highlighted",    summary.taskAHighlightedRows],
    ["Delivery cells",      summary.deliveryCellsWritten],
    ["Payment cells",       summary.paymentCellsWritten],
    ["Totals written",      summary.taskBTotalsWritten],
  ];

  summaryGrid.innerHTML = metrics
    .map(([label, value]) => `
      <div class="metric">
        <span>${label}</span>
        <strong>${value ?? 0}</strong>
      </div>
    `)
    .join("");

  summaryGrid.classList.remove("hidden");
}

function renderTaskBResults(jobTotals) {
  const rows = jobTotals
    .map((item) => `
      <tr>
        <td class="job-number">${escapeHtml(item.jobNumber)}</td>
        <td class="amount">${escapeHtml(item.total)}</td>
      </tr>
    `)
    .join("");

  taskBTable.innerHTML = `
    <table class="totals-table">
      <thead>
        <tr>
          <th>Job Number</th>
          <th>Total LPO Amount</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  taskBResults.classList.remove("hidden");
}

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
      </label>
    `)
    .join("");

  jobList.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) {
        state.selectedJobs.add(input.value);
      } else {
        state.selectedJobs.delete(input.value);
      }
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

function setBusy(isBusy) {
  taskAButton.disabled = isBusy || !state.file;
  taskBButton.disabled = isBusy || !state.file || state.selectedJobs.size === 0;
  taskAButton.classList.toggle("busy", isBusy);
  taskBButton.classList.toggle("busy", isBusy);
}

function showError(msg) {
  errorBox.textContent = msg;
  errorBox.classList.remove("hidden");
}

function hideError() {
  errorBox.textContent = "";
  errorBox.classList.add("hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
