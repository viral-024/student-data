/* app.js - Full functionality: upload, filters, sort, pagination, select, export, EmailJS send */

let rawData = []; // Array of objects (each has a __rowId field)
let headers = []; // Column header names (in order)
let selectedIds = new Set();
let currentSort = { col: null, asc: true };

let currentPage = 1;
let rowsPerPage =
  parseInt(document.getElementById("rows-per-page").value, 10) || 10;

const fileInput = document.getElementById("file-input");
const searchInput = document.getElementById("search");
const filterBranch = document.getElementById("filter-branch");
const filterYear = document.getElementById("filter-year");
const filterInterest = document.getElementById("filter-interest");
const exportBtn = document.getElementById("export-csv");
const sendEmailBtn = document.getElementById("send-email");
const tableBody = document.getElementById("table-body");
const tableHeadersRow = document.getElementById("table-headers");

const prevBtn = document.getElementById("prev-btn");
const nextBtn = document.getElementById("next-btn");
const pageInfo = document.getElementById("page-info");
const rowsPerPageSelect = document.getElementById("rows-per-page");
const statusEl = document.getElementById("status");

fileInput.addEventListener("change", handleFile, false);
searchInput.addEventListener("input", () => {
  currentPage = 1;
  renderFilteredTable();
});
filterBranch.addEventListener("change", () => {
  currentPage = 1;
  renderFilteredTable();
});
filterYear.addEventListener("change", () => {
  currentPage = 1;
  renderFilteredTable();
});
filterInterest.addEventListener("change", () => {
  currentPage = 1;
  renderFilteredTable();
});
exportBtn.addEventListener("click", exportCSV);
sendEmailBtn.addEventListener("click", sendEmailViaEmailJS);
rowsPerPageSelect.addEventListener("change", () => {
  rowsPerPage = parseInt(rowsPerPageSelect.value, 10);
  currentPage = 1;
  renderFilteredTable();
});

prevBtn.addEventListener("click", () => {
  if (currentPage > 1) {
    currentPage--;
    renderFilteredTable();
  }
});
nextBtn.addEventListener("click", () => {
  const total = Math.ceil(getFilteredData().length / rowsPerPage) || 1;
  if (currentPage < total) {
    currentPage++;
    renderFilteredTable();
  }
});

/* ------------ File reading and processing ------------ */
function handleFile(e) {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  const name = f.name.toLowerCase();

  reader.onload = (ev) => {
    const data = ev.target.result;

    if (name.endsWith(".csv")) {
      // For CSV read as text
      const text = new TextDecoder("utf-8").decode(data);
      const workbook = XLSX.read(text, { type: "string" });
      processWorkbook(workbook);
    } else {
      // xlsx, xls
      const workbook = XLSX.read(data, { type: "array" });
      processWorkbook(workbook);
    }
  };

  reader.readAsArrayBuffer(f);
}

function processWorkbook(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[firstSheetName];

  // Use header:1 to preserve header order
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!aoa || aoa.length === 0) {
    alert("No data found in sheet.");
    return;
  }

  const headerRow = aoa[0].map((h, i) =>
    h === null || h === undefined || String(h).trim() === ""
      ? `Column${i + 1}`
      : String(h).trim()
  );
  const dataRows = aoa.slice(1);

  rawData = dataRows.map((rowArr, idx) => {
    const obj = {};
    headerRow.forEach(
      (h, i) =>
        (obj[h] =
          rowArr[i] !== undefined && rowArr[i] !== null ? rowArr[i] : "")
    );
    obj.__rowId = idx + 1; // unique id
    return normalizeRow(obj);
  });

  // headers in original order (excludes __rowId)
  headers = Object.keys(rawData[0] || {}).filter((k) => k !== "__rowId");

  buildTableHeader();
  populateFilters();
  selectedIds.clear();
  currentPage = 1;
  renderFilteredTable();
}

/* ------------ Normalize values (year mapping etc) ------------ */
function normalizeRow(row) {
  const out = {};
  for (const k of Object.keys(row)) {
    const cleanKey = k.toString().trim();
    let v = row[k];
    if (typeof v === "string") v = v.trim();
    if (/^year$/i.test(cleanKey)) {
      v = normalizeYearValue(v);
    }
    out[cleanKey] = v;
  }
  // keep original __rowId if any
  if (row.__rowId) out.__rowId = row.__rowId;
  return out;
}

function normalizeYearValue(val) {
  if (typeof val === "number") return val;
  const mapping = {
    first: 1,
    second: 2,
    third: 3,
    fourth: 4,
    "1st": 1,
    "2nd": 2,
    "3rd": 3,
    "4th": 4,
  };
  const str = String(val).toLowerCase().trim();
  return mapping[str] !== undefined ? mapping[str] : parseInt(str, 10) || "";
}

/* ------------ Header building & sorting ------------ */
function buildTableHeader() {
  tableHeadersRow.innerHTML = "";

  // Select-all header (selects all filtered rows)
  const thCheck = document.createElement("th");
  thCheck.innerHTML = `<input id="select-all" type="checkbox" />`;
  tableHeadersRow.appendChild(thCheck);
  document
    .getElementById("select-all")
    .addEventListener("change", toggleSelectAllGlobal);

  for (const h of headers) {
    const th = document.createElement("th");
    th.textContent = h;
    th.style.cursor = "pointer";
    th.addEventListener("click", () => {
      // toggle sort for this column
      if (currentSort.col === h) currentSort.asc = !currentSort.asc;
      else {
        currentSort.col = h;
        currentSort.asc = true;
      }
      currentPage = 1;
      renderFilteredTable();
    });
    tableHeadersRow.appendChild(th);
  }
}

/* ------------ Filtering & Searching ------------ */
function getFilteredData() {
  const q = (searchInput.value || "").toLowerCase();
  const branch = filterBranch.value;
  const year = filterYear.value;
  const interest = filterInterest.value;

  let filtered = rawData.filter((row) => {
    if (
      branch &&
      String(row["Branch"] || "").toLowerCase() !== branch.toLowerCase()
    )
      return false;
    if (year && String(row["Year"] || "") !== year) return false;
    if (
      interest &&
      !String(row["Interests"] || "")
        .toLowerCase()
        .includes(interest.toLowerCase())
    )
      return false;
    if (
      q &&
      !Object.values(row).some((v) => String(v).toLowerCase().includes(q))
    )
      return false;
    return true;
  });

  // Sorting
  if (currentSort.col) {
    const col = currentSort.col;
    filtered.sort((a, b) => {
      const va = String(a[col] ?? "").toLowerCase();
      const vb = String(b[col] ?? "").toLowerCase();
      if (va < vb) return currentSort.asc ? -1 : 1;
      if (va > vb) return currentSort.asc ? 1 : -1;
      return 0;
    });
  }

  return filtered;
}

/* ------------ Render table (with pagination) ------------ */
function renderFilteredTable() {
  const filtered = getFilteredData();
  const totalPages = Math.max(1, Math.ceil(filtered.length / rowsPerPage));
  if (currentPage > totalPages) currentPage = totalPages;

  const start = (currentPage - 1) * rowsPerPage;
  const end = start + rowsPerPage;
  const pageData = filtered.slice(start, end);

  renderTableRows(pageData);

  // Update pagination UI
  pageInfo.textContent = `Page ${currentPage} of ${totalPages} (${filtered.length} rows)`;
  prevBtn.disabled = currentPage <= 1;
  nextBtn.disabled = currentPage >= totalPages;

  // Update select-all checkbox state based on filtered selection
  updateSelectAllCheckbox(filtered);
}

function renderTableRows(dataArr) {
  tableBody.innerHTML = "";
  for (const row of dataArr) {
    const tr = document.createElement("tr");
    tr.dataset.rowid = row.__rowId;

    // Checkbox cell
    const tdCheck = document.createElement("td");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "row-checkbox";
    cb.dataset.rowid = row.__rowId;
    cb.checked = selectedIds.has(row.__rowId);
    cb.addEventListener("change", (e) => {
      const id = Number(e.target.dataset.rowid);
      if (e.target.checked) selectedIds.add(id);
      else selectedIds.delete(id);
      // keep header select-all in sync:
      const filtered = getFilteredData();
      updateSelectAllCheckbox(filtered);
    });
    tdCheck.appendChild(cb);
    tr.appendChild(tdCheck);

    // Data columns
    for (const h of headers) {
      const td = document.createElement("td");
      td.textContent = row[h] ?? "";
      tr.appendChild(td);
    }
    tableBody.appendChild(tr);
  }
}

/* ------------ Filters population ------------ */
function populateFilters() {
  const branches = new Set();
  const years = new Set();
  const interests = new Set();

  rawData.forEach((r) => {
    if (r["Branch"]) branches.add(String(r["Branch"]));
    if (r["Year"]) years.add(String(r["Year"]));
    if (r["Interests"]) {
      String(r["Interests"])
        .split(/[,;|]+/)
        .map((s) => s.trim())
        .forEach((p) => {
          if (p) interests.add(p);
        });
    }
  });

  fillSelect(filterBranch, Array.from(branches).sort());
  fillSelect(filterYear, Array.from(years).sort());
  fillSelect(filterInterest, Array.from(interests).sort());
}

function fillSelect(selectEl, items) {
  while (selectEl.options.length > 1) selectEl.remove(1);
  items.forEach((it) => {
    const o = document.createElement("option");
    o.value = it;
    o.textContent = it;
    selectEl.appendChild(o);
  });
}

/* ------------ Select-all behavior (global across filtered results) ------------ */
function toggleSelectAllGlobal(e) {
  const checked = e.target.checked;
  const filtered = getFilteredData();
  if (checked) {
    filtered.forEach((r) => selectedIds.add(r.__rowId));
  } else {
    filtered.forEach((r) => selectedIds.delete(r.__rowId));
  }
  // re-render current page so visible checkboxes update
  renderFilteredTable();
}

function updateSelectAllCheckbox(filtered) {
  const selectAllEl = document.getElementById("select-all");
  if (!selectAllEl) return;
  const totalFiltered = filtered.length || 0;
  if (totalFiltered === 0) {
    selectAllEl.checked = false;
    selectAllEl.indeterminate = false;
    return;
  }
  // count how many filtered are selected
  let selectedCount = 0;
  filtered.forEach((r) => {
    if (selectedIds.has(r.__rowId)) selectedCount++;
  });

  if (selectedCount === 0) {
    selectAllEl.checked = false;
    selectAllEl.indeterminate = false;
  } else if (selectedCount === totalFiltered) {
    selectAllEl.checked = true;
    selectAllEl.indeterminate = false;
  } else {
    selectAllEl.checked = false;
    selectAllEl.indeterminate = true;
  }
}

/* ------------ Export CSV (selected or all filtered) ------------ */
function exportCSV() {
  const filtered = getFilteredData();

  // If some rows selected => export those; else export all filtered
  let rowsToExport = [];
  if (selectedIds.size > 0) {
    rowsToExport = rawData.filter((r) => selectedIds.has(r.__rowId));
  } else {
    rowsToExport = filtered;
  }

  if (!rowsToExport.length) {
    alert("No rows to export.");
    return;
  }

  const csvRows = [];
  csvRows.push(headers.map((h) => `"${h.replace(/"/g, '""')}"`));

  rowsToExport.forEach((r) => {
    const line = headers.map(
      (h) => `"${String(r[h] ?? "").replace(/"/g, '""')}"`
    );
    csvRows.push(line);
  });

  const csvContent = csvRows.map((r) => r.join(",")).join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "students_export.csv";
  a.click();
  URL.revokeObjectURL(url);
}

/* ------------ Send Email via EmailJS to selected candidates ------------ */
function sendEmailViaEmailJS() {
  // Ensure EmailJS is initialized in index.html via emailjs.init("PUBLIC_KEY")
  const emailHeader = headers.find((h) => /email/i.test(h));
  if (!emailHeader) {
    alert(
      'No "Email" column found in the file (column name should include "email").'
    );
    return;
  }

  // Collect selected rows
  const selectedRows = rawData.filter((r) => selectedIds.has(r.__rowId));
  if (!selectedRows.length) {
    alert("Select at least one row to send emails.");
    return;
  }

  // Build promises for each selected row
  const serviceId = "service_q79vzij"; // <-- replace
  const templateId = "template_vy7usfj"; // <-- replace

  // Make sure the template in EmailJS has variables like {{to_email}}, {{subject}}, {{message}}, etc.
  const promises = selectedRows.map((row) => {
    const toEmail = row[emailHeader];
    if (!toEmail || String(toEmail).trim() === "") {
      return Promise.resolve({ status: "skipped", email: toEmail });
    }

    // Customize payload as per your EmailJS template variables
    const payload = {
      to_email: String(toEmail).trim(),
      subject: `Hello from Student Dashboard`,
      message: `Hello ${
        row["Name"] ?? ""
      },\n\nThis is a test email from the Student Dashboard.`,
      // you can add other template variables here (like name, branch, year, etc.)
      name: row["Name"] ?? "",
      branch: row["Branch"] ?? "",
      year: row["Year"] ?? "",
    };

    return emailjs
      .send(serviceId, templateId, payload)
      .then(() => ({ status: "ok", email: toEmail }))
      .catch((err) => ({ status: "error", email: toEmail, err }));
  });

  statusEl.textContent = "Sending emails…";
  Promise.all(promises).then((results) => {
    const ok = results.filter((r) => r.status === "ok").length;
    const skipped = results.filter((r) => r.status === "skipped").length;
    const err = results.filter((r) => r.status === "error");
    const errCount = err.length;

    statusEl.textContent = `Done: sent ${ok}, skipped ${skipped}, failed ${errCount}. See console for details.`;
    console.log("Email send results:", results);
    alert(
      `Email send finished.\nSent: ${ok}\nSkipped: ${skipped}\nFailed: ${errCount}`
    );
  });
}
