const DB_NAME = "gjt-product-database";
const DB_VERSION = 1;
const LOCAL_EXCEL_URL = "./%E6%88%90%E6%9C%AC%E8%A8%88%E7%AE%97%EF%BC%BF%E5%85%A7%E9%83%A8%E6%AA%94%E6%A1%88%20202604.xlsx";

const costKeys = ["exw", "fob", "freight40", "cifFactory", "cifTwd"];

const dataColumns = [
  { key: "code", label: "GY編號", aliases: ["GY編號", "編號", "序號", "產品編號"] },
  { key: "type", label: "型態", aliases: ["型態", "產品型態", "常溫/冷凍", "溫層", "類型"] },
  { key: "supplier", label: "供應商", aliases: ["供應商", "供應商（工廠）", "廠商", "Supplier"] },
  { key: "brand", label: "品牌", aliases: ["品牌", "Brand"] },
  { key: "nameEn", label: "產品品名", aliases: ["產品品名", "英文品名", "英文名稱", "English", "品名(英)"] },
  { key: "nameZh", label: "中文品名", aliases: ["中文品名", "品名", "產品名稱", "商品名稱"] },
  { key: "spec", label: "規格", aliases: ["規格", "包裝", "容量", "重量"] },
  { key: "ean", label: "國際條碼", aliases: ["國際條碼", "EAN", "EAN條碼", "條碼", "Barcode"] },
  { key: "perBox", label: "入/箱", numeric: true, aliases: ["入/箱", "每箱", "每箱入數", "箱入數"] },
  { key: "boxPerPallet", label: "箱/板", numeric: true, aliases: ["箱/板", "每板箱數", "棧板箱數"] },
  { key: "container20", label: '20"箱', numeric: true, aliases: ['20"箱', "20呎箱", "20尺箱"] },
  { key: "container40", label: '40"箱', numeric: true, aliases: ['40"箱', "40呎箱", "40尺箱"] },
  { key: "shelfLife", label: "保存期限(月)", numeric: true, aliases: ["保存期限(月）", "保存期限(月)", "BBD", "效期"] },
  { key: "origin", label: "產地", aliases: ["產地", "原產地", "國家"] },
  { key: "currency", label: "幣別", aliases: ["幣別", "Currency"] },
  { key: "exw", label: "EXW報價", numeric: true, aliases: ["EXW報價", "EXW 報價"] },
  { key: "fob", label: "FOB報價", numeric: true, aliases: ["FOB 報價", "FOB報價"] },
  { key: "freight40", label: '40"運費', numeric: true, aliases: ['40"運費', "40呎運費"] },
  { key: "cifFactory", label: "CIF原廠報價", numeric: true, aliases: ["CIF 原廠報價", "CIF原廠報價"] },
  { key: "cifTwd", label: "CIF台幣報價", numeric: true, aliases: ["CIF 台幣報價", "CIF台幣報價"] },
  { key: "term", label: "交易方式", aliases: ["高玉轉手／直採", "合作模式", "交易方式"] },
  { key: "source", label: "狀態", aliases: ["來源", "狀態", "工作表"] },
];

const quoteDefaultKeys = ["code", "nameZh", "nameEn", "spec", "ean", "origin", "currency", "cifFactory", "cifTwd", "supplier", "term"];

const state = {
  db: null,
  datasets: [],
  pdfPages: [],
  activeDatasetId: null,
  activeRows: [],
  filteredRows: [],
  selectedIds: new Set(),
  editMode: false,
  pendingEdits: new Map(),
  sortKey: null,
  sortAsc: true,
};

const els = {
  sourceLine: document.querySelector("#sourceLine"),
  importer: document.querySelector("#importer"),
  openImporter: document.querySelector("#openImporter"),
  loadLocalExcel: document.querySelector("#loadLocalExcel"),
  clearDb: document.querySelector("#clearDb"),
  printCurrent: document.querySelector("#printCurrent"),
  fileInput: document.querySelector("#fileInput"),
  dropZone: document.querySelector("#dropZone"),
  log: document.querySelector("#log"),
  totalRows: document.querySelector("#totalRows"),
  supplierCount: document.querySelector("#supplierCount"),
  originCount: document.querySelector("#originCount"),
  pdfPageCount: document.querySelector("#pdfPageCount"),
  datasetSelect: document.querySelector("#datasetSelect"),
  search: document.querySelector("#search"),
  filters: {
    type: document.querySelector("#fType"),
    supplier: document.querySelector("#fSupplier"),
    origin: document.querySelector("#fOrigin"),
    biz: document.querySelector("#fBiz"),
    source: document.querySelector("#fSource"),
  },
  visibleRows: document.querySelector("#visibleRows"),
  datasetRows: document.querySelector("#datasetRows"),
  tableHead: document.querySelector("#tableHead"),
  tableBody: document.querySelector("#tableBody"),
  tableFoot: document.querySelector("#tableFoot"),
  toggleEdit: document.querySelector("#toggleEdit"),
  saveEdits: document.querySelector("#saveEdits"),
  cancelEdits: document.querySelector("#cancelEdits"),
  editStatus: document.querySelector("#editStatus"),
  quoteIncludeCost: document.querySelector("#quoteIncludeCost"),
  fieldPicker: document.querySelector("#fieldPicker"),
  quoteTableHead: document.querySelector("#quoteTableHead"),
  quoteTableBody: document.querySelector("#quoteTableBody"),
  quoteTableFoot: document.querySelector("#quoteTableFoot"),
  quoteCustomer: document.querySelector("#quoteCustomer"),
  quoteTerms: document.querySelector("#quoteTerms"),
  quoteValidUntil: document.querySelector("#quoteValidUntil"),
  printQuote: document.querySelector("#printQuote"),
  selectFiltered: document.querySelector("#selectFiltered"),
  clearSelection: document.querySelector("#clearSelection"),
  exportSelectedPdf: document.querySelector("#exportSelectedPdf"),
  selectedCount: document.querySelector("#selectedCount"),
  pdfSearch: document.querySelector("#pdfSearch"),
  pdfResults: document.querySelector("#pdfResults"),
};

if (window.pdfjsLib?.GlobalWorkerOptions) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
}

boot();

async function boot() {
  state.db = await openDb();
  bindEvents();
  renderDataHeader();
  renderQuoteFieldPicker();
  await refreshFromDb();
  if (!state.datasets.length) {
    if (window.GJT_EMBEDDED_DATA) await loadEmbeddedData();
    else await loadLocalExcel(true);
  }
  writeLog("資料庫已啟動。預設資料來源為成本計算＿內部檔案 202604.xlsx。");
}

function bindEvents() {
  els.openImporter.addEventListener("click", () => els.importer.classList.toggle("open"));
  els.loadLocalExcel.addEventListener("click", () => loadLocalExcel(false));
  els.clearDb.addEventListener("click", clearDb);
  els.printCurrent.addEventListener("click", () => window.print());
  els.printQuote.addEventListener("click", () => {
    switchView("quoteView");
    window.print();
  });
  els.exportSelectedPdf.addEventListener("click", exportSelectedPdf);
  els.toggleEdit.addEventListener("click", toggleEditMode);
  els.saveEdits.addEventListener("click", saveEdits);
  els.cancelEdits.addEventListener("click", cancelEdits);
  els.selectFiltered.addEventListener("click", () => {
    state.filteredRows.forEach((row) => state.selectedIds.add(row._id));
    renderTables();
  });
  els.clearSelection.addEventListener("click", () => {
    state.selectedIds.clear();
    renderTables();
  });
  els.fileInput.addEventListener("change", async (event) => {
    await importFiles([...event.target.files]);
    els.fileInput.value = "";
  });
  els.datasetSelect.addEventListener("change", () => {
    state.activeDatasetId = Number(els.datasetSelect.value) || null;
    state.selectedIds.clear();
    selectActiveRows();
    renderAll();
  });
  els.search.addEventListener("input", renderTables);
  Object.values(els.filters).forEach((el) => el.addEventListener("change", renderTables));
  els.quoteIncludeCost.addEventListener("change", () => {
    renderQuoteFieldPicker();
    renderQuoteTable();
  });
  els.fieldPicker.addEventListener("change", renderQuoteTable);
  els.pdfSearch.addEventListener("input", renderPdfResults);

  els.tableBody.addEventListener("change", handleSelectionChange);
  els.tableBody.addEventListener("input", handleCellEdit);
  els.quoteTableBody.addEventListener("change", handleSelectionChange);

  document.querySelectorAll(".view-tabs button").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.view));
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("dragover");
    });
  });
  els.dropZone.addEventListener("drop", async (event) => importFiles([...event.dataTransfer.files]));
}

function handleSelectionChange(event) {
  if (!event.target.matches(".row-select")) return;
  if (event.target.checked) state.selectedIds.add(event.target.value);
  else state.selectedIds.delete(event.target.value);
  renderTables();
}

function handleCellEdit(event) {
  const editor = event.target.closest(".cell-editor");
  if (!editor || !state.editMode) return;
  const rowId = editor.dataset.rowId;
  const key = editor.dataset.key;
  const editKey = `${rowId}::${key}`;
  const value = editor.value;
  state.pendingEdits.set(editKey, { rowId, key, value });
  editor.classList.add("edited");
  updateEditStatus();
}

function toggleEditMode() {
  state.editMode = !state.editMode;
  if (!state.editMode && state.pendingEdits.size) {
    if (!confirm("尚有未儲存變更，確定離開編輯模式？")) {
      state.editMode = true;
      return;
    }
  }
  if (!state.editMode) state.pendingEdits.clear();
  renderTables();
  updateEditStatus();
}

async function saveEdits() {
  if (!state.pendingEdits.size) return;
  const dataset = state.datasets.find((item) => item.id === state.activeDatasetId);
  if (!dataset) return;

  const edits = [...state.pendingEdits.values()];
  edits.forEach(({ rowId, key, value }) => {
    const row = state.activeRows.find((item) => item._id === rowId);
    if (!row) return;
    row[key] = value;
    if (row._raw) row._raw[key] = value;
  });

  dataset.rows = state.activeRows.map(({ _id, ...row }) => row);
  dataset.rowCount = dataset.rows.length;
  dataset.editedAt = new Date().toISOString();
  await put("datasets", dataset);
  state.pendingEdits.clear();
  state.activeDatasetId = dataset.id;
  await refreshFromDb();
  updateEditStatus();
  writeLog(`已儲存 ${edits.length} 個儲存格變更。`);
}

function cancelEdits() {
  state.pendingEdits.clear();
  renderTables();
  updateEditStatus();
}

function updateEditStatus() {
  els.toggleEdit.textContent = state.editMode ? "關閉編輯" : "開啟編輯";
  els.saveEdits.disabled = !state.pendingEdits.size;
  els.cancelEdits.disabled = !state.pendingEdits.size;
  els.editStatus.textContent = state.editMode
    ? `編輯模式，未儲存 ${state.pendingEdits.size} 格`
    : "檢視模式";
}

function switchView(viewId) {
  document.querySelectorAll(".view-tabs button").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === viewId);
  });
  document.querySelectorAll(".view-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === viewId);
  });
}

function renderDataHeader() {
  els.tableHead.innerHTML = `<tr><th class="select-col">選取</th>${dataColumns.map((col) => `<th data-key="${col.key}">${escapeHtml(col.label)} ↕</th>`).join("")}</tr>`;
  els.tableHead.querySelectorAll("th[data-key]").forEach((th) => {
    th.addEventListener("click", () => {
      state.sortAsc = state.sortKey === th.dataset.key ? !state.sortAsc : true;
      state.sortKey = th.dataset.key;
      renderTables();
    });
  });
}

function renderQuoteFieldPicker() {
  const includeCost = els.quoteIncludeCost.checked;
  const cols = dataColumns.filter((col) => includeCost || !costKeys.includes(col.key));
  els.fieldPicker.innerHTML = cols.map((col) => {
    const checked = quoteDefaultKeys.includes(col.key) ? "checked" : "";
    return `<label><input type="checkbox" value="${col.key}" ${checked}> ${escapeHtml(col.label)}</label>`;
  }).join("");
}

async function loadLocalExcel(silent) {
  try {
    const response = await fetch(LOCAL_EXCEL_URL);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const blob = await response.blob();
    const file = new File([blob], "成本計算＿內部檔案 202604.xlsx", { type: blob.type });
    await importFiles([file]);
  } catch (error) {
    if (!silent) writeLog(`無法自動載入成本 Excel：${error.message}。請使用 http://localhost 或直接上傳檔案。`);
  }
}

async function loadEmbeddedData() {
  const data = window.GJT_EMBEDDED_DATA;
  if (!data?.rows?.length) return;
  await put("datasets", {
    fileName: data.fileName || "成本計算＿內部檔案 202604.xlsx",
    importedAt: data.importedAt || new Date().toISOString(),
    sheetNames: data.sheetNames || [],
    rowCount: data.rows.length,
    rows: data.rows,
    embedded: true,
  });
  await refreshFromDb();
  writeLog(`已載入內嵌資料：${data.rows.length} 筆。`);
}

async function clearDb() {
  if (!confirm("確定清空本機瀏覽器資料庫？原始 Excel 檔案不會被刪除。")) return;
  await clearStore("datasets");
  await clearStore("pdfPages");
  state.selectedIds.clear();
  await refreshFromDb();
  writeLog("已清空本機資料庫。");
}

async function importFiles(files) {
  for (const file of files) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (["xlsx", "xls", "csv"].includes(ext)) await importWorkbook(file);
    else if (ext === "pdf") await importPdf(file);
    else writeLog(`略過不支援格式：${file.name}`);
  }
  await refreshFromDb();
}

async function importWorkbook(file) {
  if (!window.XLSX) {
    alert("Excel 解析套件尚未載入。若要上傳 Excel，請確認網路可載入 SheetJS，或改用已內嵌的預設資料。");
    return;
  }
  writeLog(`解析 Excel：${file.name}`);
  const workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const rows = [];
  const sheetNames = [];

  workbook.SheetNames.forEach((sheetName) => {
    const rawRows = rowsFromSheet(workbook.Sheets[sheetName]);
    if (!rawRows.length) return;
    sheetNames.push(sheetName);
    rawRows.forEach((raw) => {
      const row = normalizeRow(raw, sheetName);
      if (!hasMeaningfulData(row)) return;
      row.code = row.code || `GY-${String(rows.length + 1).padStart(3, "0")}`;
      row.source = row.source || inferSource(sheetName);
      row.biz = row.term;
      row._raw = raw;
      rows.push(row);
    });
  });

  await put("datasets", {
    fileName: file.name,
    importedAt: new Date().toISOString(),
    sheetNames,
    rowCount: rows.length,
    rows,
  });
  writeLog(`已匯入 ${rows.length} 筆產品資料。`);
}

function rowsFromSheet(sheet) {
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  const headerIndex = rows.findIndex((row) => {
    const text = row.map(clean).join("|");
    return text.includes("產品型態") && text.includes("供應商") && (text.includes("中文品名") || text.includes("產品品名"));
  });
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((value) => String(value || "").trim());
  return rows.slice(headerIndex + 1).map((cells) => {
    const record = {};
    headers.forEach((header, index) => {
      if (header) record[header] = cells[index] ?? "";
    });
    return record;
  }).filter((record) => !isHeaderRecord(record));
}

function isHeaderRecord(record) {
  const values = Object.values(record).map(clean).filter(Boolean);
  if (!values.length) return true;
  const text = values.join("|");
  const headerHits = [
    "產品型態",
    "供應商",
    "產品品名",
    "中文品名",
    "規格",
    "國際條碼",
    "入/箱",
  ].filter((label) => text.includes(clean(label))).length;
  return headerHits >= 4;
}

async function importPdf(file) {
  if (!window.pdfjsLib) {
    alert("PDF 解析套件尚未載入。若要上傳 PDF，請確認網路可載入 PDF.js。");
    return;
  }
  writeLog(`解析 PDF：${file.name}`);
  const pdf = await window.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const text = textContent.items.map((item) => item.str).join(" ").replace(/\s+/g, " ").trim();
    await put("pdfPages", { fileName: file.name, pageNumber, text, importedAt: new Date().toISOString() });
  }
  writeLog(`已建立 PDF 索引：${file.name}，${pdf.numPages} 頁。`);
}

function normalizeRow(raw, sheetName) {
  const row = {};
  dataColumns.forEach((col) => {
    row[col.key] = findValue(raw, col.aliases);
  });
  row.source = row.source || inferSource(sheetName);
  return row;
}

function findValue(raw, aliases) {
  const entries = Object.entries(raw);
  for (const alias of aliases) {
    const direct = entries.find(([key]) => clean(key) === clean(alias));
    if (direct) return direct[1];
  }
  for (const alias of aliases) {
    const fuzzy = entries.find(([key]) => clean(key).includes(clean(alias)) || clean(alias).includes(clean(key)));
    if (fuzzy) return fuzzy[1];
  }
  return "";
}

function inferSource(sheetName) {
  if (String(sheetName).includes("已進口")) return "已進口";
  if (String(sheetName).includes("討論")) return "討論中";
  return sheetName;
}

function hasMeaningfulData(row) {
  return ["nameZh", "nameEn", "supplier", "ean", "cifFactory", "cifTwd"].some((key) => hasValue(row[key]));
}

async function refreshFromDb() {
  state.datasets = await getAll("datasets");
  state.pdfPages = await getAll("pdfPages");
  state.activeDatasetId = state.activeDatasetId || state.datasets.at(-1)?.id || null;
  selectActiveRows();
  renderAll();
}

function selectActiveRows() {
  const active = state.datasets.find((item) => item.id === state.activeDatasetId) || state.datasets.at(-1);
  state.activeDatasetId = active?.id || null;
  state.activeRows = (active?.rows || []).map((row, index) => ({
    ...row,
    _id: `${state.activeDatasetId || "dataset"}-${row.code || index}-${index}`,
  }));
  state.selectedIds = new Set([...state.selectedIds].filter((id) => state.activeRows.some((row) => row._id === id)));
}

function renderAll() {
  renderDatasetSelect();
  renderFilters();
  renderSummary();
  renderTables();
  renderPdfResults();
  updateEditStatus();
}

function renderDatasetSelect() {
  els.datasetSelect.innerHTML = "";
  if (!state.datasets.length) {
    els.datasetSelect.append(new Option("尚無資料集", ""));
    els.sourceLine.textContent = "資料來源：成本計算＿內部檔案 202604.xlsx";
    return;
  }
  state.datasets.forEach((dataset) => {
    const date = new Date(dataset.importedAt).toLocaleString("zh-TW", { hour12: false });
    els.datasetSelect.append(new Option(`${dataset.fileName}｜${dataset.rowCount} 筆｜${date}`, dataset.id));
  });
  els.datasetSelect.value = state.activeDatasetId || "";
  const active = state.datasets.find((item) => item.id === state.activeDatasetId);
  els.sourceLine.textContent = active
    ? `資料來源：${active.fileName}｜更新時間：${new Date(active.importedAt).toLocaleString("zh-TW", { hour12: false })}｜共 ${active.rowCount} 筆`
    : "資料來源：成本計算＿內部檔案 202604.xlsx";
}

function renderFilters() {
  setOptions(els.filters.type, unique("type"));
  setOptions(els.filters.supplier, unique("supplier"));
  setOptions(els.filters.origin, unique("origin"));
  setOptions(els.filters.biz, unique("biz"));
  setOptions(els.filters.source, unique("source"));
}

function setOptions(select, values) {
  const current = select.value;
  select.innerHTML = '<option value="">全部</option>';
  values.forEach((value) => select.append(new Option(value, value)));
  select.value = values.includes(current) ? current : "";
}

function unique(key) {
  return [...new Set(state.activeRows.map((row) => String(row[key] || "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "zh-Hant"));
}

function renderSummary() {
  els.totalRows.textContent = state.activeRows.length.toLocaleString();
  els.supplierCount.textContent = unique("supplier").length.toLocaleString();
  els.originCount.textContent = unique("origin").length.toLocaleString();
  els.pdfPageCount.textContent = state.pdfPages.length.toLocaleString();
  els.datasetRows.textContent = state.activeRows.length.toLocaleString();
}

function renderTables() {
  state.filteredRows = getFilteredRows();
  if (state.sortKey) {
    const col = dataColumns.find((item) => item.key === state.sortKey);
    state.filteredRows.sort((a, b) => compareValue(a[state.sortKey], b[state.sortKey], col?.numeric));
    if (!state.sortAsc) state.filteredRows.reverse();
  }
  renderDataTable();
  renderQuoteTable();
  renderSelectionStatus();
}

function getFilteredRows() {
  const query = clean(els.search.value);
  return state.activeRows.filter((row) => {
    const textMatch = !query || clean(Object.values(row._raw || row).join(" ")).includes(query);
    return textMatch
      && filterMatch(row.type, els.filters.type.value)
      && filterMatch(row.supplier, els.filters.supplier.value)
      && filterMatch(row.origin, els.filters.origin.value)
      && filterMatch(row.biz, els.filters.biz.value)
      && filterMatch(row.source, els.filters.source.value);
  });
}

function renderDataTable() {
  els.visibleRows.textContent = state.filteredRows.length.toLocaleString();
  els.tableBody.innerHTML = state.filteredRows.slice(0, 800).map((row) => rowHtml(row, dataColumns)).join("");
  els.tableFoot.colSpan = dataColumns.length + 1;
  els.tableFoot.textContent = state.activeRows.length
    ? `合計：${state.filteredRows.length} / ${state.activeRows.length} 筆產品，${unique("supplier").length} 家供應商，涵蓋 ${unique("origin").length} 個產地`
    : "尚未載入資料。";
  if (!state.filteredRows.length) {
    els.tableBody.innerHTML = `<tr><td class="empty" colspan="${dataColumns.length + 1}">沒有符合條件的資料。</td></tr>`;
  }
}

function renderQuoteTable() {
  const rows = getSelectedRows();
  const quoteColumns = getVisibleQuoteColumns(rows);
  els.quoteTableHead.innerHTML = `<tr><th class="select-col">選取</th><th class="num">#</th>${quoteColumns.map((col) => `<th>${escapeHtml(col.label)}</th>`).join("")}</tr>`;
  els.quoteTableBody.innerHTML = rows.map((row, index) => {
    const cells = quoteColumns.map((col) => `<td class="${col.numeric ? "num" : ""}">${escapeHtml(displayValue(row[col.key]))}</td>`).join("");
    return `<tr>${selectCell(row)}<td class="num">${index + 1}</td>${cells}</tr>`;
  }).join("");
  els.quoteTableFoot.colSpan = quoteColumns.length + 2;
  els.quoteTableFoot.textContent = `報價表格：共 ${rows.length} 筆已勾選商品`;
  if (!rows.length) {
    els.quoteTableBody.innerHTML = `<tr><td class="empty" colspan="${quoteColumns.length + 2}">請先勾選要輸出的商品。</td></tr>`;
  }
}

function getSelectedRows() {
  return state.filteredRows.filter((row) => state.selectedIds.has(row._id));
}

function getVisibleQuoteColumns(rows) {
  const selectedKeys = [...els.fieldPicker.querySelectorAll("input:checked")].map((input) => input.value);
  return dataColumns
    .filter((col) => selectedKeys.includes(col.key))
    .filter((col) => rows.length === 0 || rows.some((row) => hasValue(row[col.key])));
}

function rowHtml(row, columns) {
  const cls = row.source === "已進口" ? "imported" : row.source === "討論中" ? "discussing" : "";
  return `<tr class="${cls}">${selectCell(row)}${columns.map((col) => cellHtml(row, col)).join("")}</tr>`;
}

function selectCell(row) {
  const checked = state.selectedIds.has(row._id) ? "checked" : "";
  return `<td class="select-col"><input class="row-select" type="checkbox" value="${escapeHtml(row._id)}" ${checked}></td>`;
}

function cellHtml(row, col) {
  const value = displayValue(row[col.key]);
  const title = escapeHtml(value);
  const editKey = `${row._id}::${col.key}`;
  const edited = state.pendingEdits.has(editKey);
  const currentValue = edited ? state.pendingEdits.get(editKey).value : value;
  const safeValue = escapeHtml(currentValue);
  const classes = [col.numeric ? "num" : "", col.key === "code" || col.key === "ean" ? "code" : "", state.editMode ? "editable" : "", edited ? "edited" : ""].filter(Boolean).join(" ");
  if (state.editMode) {
    return `<td class="${classes}" title="${safeValue}"><input class="cell-editor" type="text" value="${safeValue}" data-row-id="${escapeHtml(row._id)}" data-key="${escapeHtml(col.key)}"></td>`;
  }
  if (col.key === "type") return `<td><span class="badge ${String(value).includes("冷") ? "badge-frozen" : "badge-normal"}">${title}</span></td>`;
  if (col.key === "source") return `<td><span class="badge ${value === "已進口" ? "badge-src-imp" : "badge-src-dis"}">${title}</span></td>`;
  return `<td class="${classes}" title="${safeValue}">${safeValue}</td>`;
}

function renderSelectionStatus() {
  els.selectedCount.textContent = state.selectedIds.size.toLocaleString();
}

function exportSelectedPdf() {
  const rows = getSelectedRows();
  if (!rows.length) {
    alert("請先勾選要輸出的商品。");
    return;
  }
  switchView("quoteView");
  window.print();
}

function renderPdfResults() {
  const query = clean(els.pdfSearch.value);
  const pages = query ? state.pdfPages.filter((page) => clean(page.text).includes(query)) : state.pdfPages.slice(-20);
  if (!pages.length) {
    els.pdfResults.innerHTML = '<div class="empty">尚無 PDF 索引，或沒有符合搜尋條件。</div>';
    return;
  }
  els.pdfResults.innerHTML = pages.slice(0, 80).map((page) => {
    const text = page.text.length > 320 ? `${page.text.slice(0, 320)}...` : page.text;
    return `<article class="pdf-hit"><strong>${escapeHtml(page.fileName)}｜第 ${page.pageNumber} 頁</strong><br>${escapeHtml(text)}</article>`;
  }).join("");
}

function filterMatch(value, filter) {
  return !filter || String(value || "") === filter;
}

function compareValue(a, b, numeric) {
  if (numeric) {
    const an = parseNumber(a);
    const bn = parseNumber(b);
    if (!Number.isNaN(an) && !Number.isNaN(bn)) return an - bn;
  }
  return String(a || "").localeCompare(String(b || ""), "zh-Hant");
}

function parseNumber(value) {
  return Number(String(value || "").replace(/[,%]/g, ""));
}

function clean(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, "");
}

function displayValue(value) {
  return hasValue(value) ? String(value).trim() : "";
}

function hasValue(value) {
  return String(value ?? "").trim() !== "";
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("datasets")) db.createObjectStore("datasets", { keyPath: "id", autoIncrement: true });
      if (!db.objectStoreNames.contains("pdfPages")) db.createObjectStore("pdfPages", { keyPath: "id", autoIncrement: true });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function put(storeName, value) {
  return new Promise((resolve, reject) => {
    const tx = state.db.transaction(storeName, "readwrite");
    tx.objectStore(storeName).put(value);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

function getAll(storeName) {
  return new Promise((resolve, reject) => {
    const tx = state.db.transaction(storeName, "readonly");
    const request = tx.objectStore(storeName).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

function clearStore(storeName) {
  return new Promise((resolve, reject) => {
    const tx = state.db.transaction(storeName, "readwrite");
    tx.objectStore(storeName).clear();
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

function writeLog(message) {
  const time = new Date().toLocaleTimeString("zh-TW", { hour12: false });
  els.log.textContent = `[${time}] ${message}\n${els.log.textContent}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
