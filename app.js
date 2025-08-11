let rawData = [];
let headers = [];

const fileInput = document.getElementById("file-input");
const searchInput = document.getElementById("search");
const filterBranch = document.getElementById("filter-branch");
const filterYear = document.getElementById("filter-year");
const filterInterest = document.getElementById("filter-interest");
const exportBtn = document.getElementById("export-csv");
const sendEmailBtn = document.getElementById("send-email");
const tableBody = document.getElementById("table-body");
const tableHeadersRow = document.getElementById("table-headers");

fileInput.addEventListener("change", handleFile, false);
searchInput.addEventListener("input", renderFilteredTable);
filterBranch.addEventListener("change", renderFilteredTable);
filterYear.addEventListener("change", renderFilteredTable);
filterInterest.addEventListener("change", renderFilteredTable);
exportBtn.addEventListener("click", exportCSV);
sendEmailBtn.addEventListener("click", sendEmailViaEmailJS);

function handleFile(e) {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  const name = f.name.toLowerCase();

  reader.onload = (ev) => {
    const data = ev.target.result;

    if (name.endsWith(".csv")) {
      const text = new TextDecoder("utf-8").decode(data);
      const workbook = XLSX.read(text, { type: "string" });
      processWorkbook(workbook);
    } else {
      const workbook = XLSX.read(data, { type: "array" });
      processWorkbook(workbook);
    }
  };

  reader.readAsArrayBuffer(f);
}

function processWorkbook(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[firstSheetName];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  rawData = json.map(normalizeRow);
  headers = Object.keys(rawData[0] || {});
  buildTableHeader();
  populateFilters();
  renderFilteredTable();
}

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
  return mapping[str] !== undefined ? mapping[str] : parseInt(str) || "";
}

function buildTableHeader() {
  tableHeadersRow.innerHTML = "";
  const thCheck = document.createElement("th");
  thCheck.innerHTML = `<input id="select-all" type="checkbox" />`;
  tableHeadersRow.appendChild(thCheck);
  document
    .getElementById("select-all")
    .addEventListener("change", toggleSelectAll);

  for (const h of headers) {
    const th = document.createElement("th");
    th.textContent = h;
    th.style.cursor = "pointer";
    th.addEventListener("click", () => sortByColumn(h));
    tableHeadersRow.appendChild(th);
  }
}

function renderFilteredTable() {
  const q = (searchInput.value || "").toLowerCase();
  const branch = filterBranch.value;
  const year = filterYear.value;
  const interest = filterInterest.value;

  const filtered = rawData.filter((row) => {
    if (branch && String(row["Branch"]).toLowerCase() !== branch.toLowerCase())
      return false;
    if (year && String(row["Year"]) !== year) return false;
    if (
      interest &&
      !String(row["Interests"]).toLowerCase().includes(interest.toLowerCase())
    )
      return false;
    if (
      q &&
      !Object.values(row).some((v) => String(v).toLowerCase().includes(q))
    )
      return false;
    return true;
  });

  renderTableRows(filtered);
}

function renderTableRows(dataArr) {
  tableBody.innerHTML = "";
  for (const row of dataArr) {
    const tr = document.createElement("tr");
    const tdCheck = document.createElement("td");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "row-checkbox";
    tdCheck.appendChild(cb);
    tr.appendChild(tdCheck);

    for (const h of headers) {
      const td = document.createElement("td");
      td.textContent = row[h];
      tr.appendChild(td);
    }
    tableBody.appendChild(tr);
  }
}

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
        .forEach((p) => interests.add(p));
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

let currentSort = { col: null, asc: true };
function sortByColumn(col) {
  rawData.sort((a, b) => {
    const va = String(a[col]).toLowerCase();
    const vb = String(b[col]).toLowerCase();
    if (va < vb) return currentSort.asc ? -1 : 1;
    if (va > vb) return currentSort.asc ? 1 : -1;
    return 0;
  });
  currentSort.asc = currentSort.col === col ? !currentSort.asc : true;
  currentSort.col = col;
  renderFilteredTable();
}

function exportCSV() {
  const rows = [headers];
  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const cells = Array.from(tr.querySelectorAll("td")).slice(1);
    rows.push(cells.map((td) => `"${td.textContent.replace(/"/g, '""')}"`));
  });
  const csvContent = rows.map((r) => r.join(",")).join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "students_export.csv";
  a.click();
  URL.revokeObjectURL(url);
}
function sendEmailViaEmailJS() {
  const emailHeader = headers.find((h) => /email/i.test(h));
  if (!emailHeader) {
    alert('No "Email" column found in your file.');
    return;
  }

  const selectedEmails = [];
  const emailColIndex = headers.indexOf(emailHeader);

  document.querySelectorAll("#table-body tr").forEach((tr) => {
    const cb = tr.querySelector(".row-checkbox");
    if (cb && cb.checked) {
      const cells = tr.querySelectorAll("td");
      const email = cells[emailColIndex]?.textContent.trim();
      if (email) selectedEmails.push(email);
    }
  });

  if (!selectedEmails.length) {
    alert("Please select at least one row.");
    return;
  }

  // Initialize EmailJS (replace with your real user/public key)
  emailjs.init("ECd6WqdgS-YdP-66S");

  // Send emails one-by-one
  selectedEmails.forEach((email) => {
    emailjs
      .send("service_q79vzij", "template_vy7usfj", {
        to_email: email, // Dynamic recipient
        subject: "Hello from Student Dashboard",
        message:
          "This is a test email sent via EmailJS from the uploaded Excel sheet.",
      })
      .then(() => {
        console.log(`✅ Email sent to ${email}`);
      })
      .catch((err) => {
        console.error(`❌ Failed to send to ${email}`, err);
      });
  });
}

function toggleSelectAll(e) {
  const checked = e.target.checked;
  document
    .querySelectorAll(".row-checkbox")
    .forEach((cb) => (cb.checked = checked));
}
