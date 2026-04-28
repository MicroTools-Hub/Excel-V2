const state = {
  messages: [],
  parsedRows: [],
  isExcelHost: false,
  isBusy: false,
  lastSheetSnapshot: null,
  latestImageContext: null,
};

const OCR_SERVER_TIMEOUT_MS = 12000;
const OCR_BROWSER_TIMEOUT_MS = 45000;
const OCR_MAX_IMAGE_EDGE = 1800;
const SHEET_SNAPSHOT_MAX_ROWS = 120;
const SHEET_SNAPSHOT_MAX_COLS = 24;
const SHEET_SELECTION_MAX_ROWS = 30;
const SHEET_SELECTION_MAX_COLS = 12;
const IMAGE_CONTEXT_MAX_ROWS = 80;
const IMAGE_CONTEXT_MAX_COLS = 16;
const IMAGE_CONTEXT_TEXT_MAX_CHARS = 6000;

const ui = {
  messages: document.getElementById("messages"),
  chatInput: document.getElementById("chatInput"),
  chatImageBtn: document.getElementById("chatImageBtn"),
  sendBtn: document.getElementById("sendBtn"),
  clearChatBtn: document.getElementById("clearChatBtn"),
  officeStatus: document.getElementById("officeStatus"),
  filepathInput: document.getElementById("filepathInput"),
  sheetInput: document.getElementById("sheetInput"),
  startCellInput: document.getElementById("startCellInput"),
  pasteInput: document.getElementById("pasteInput"),
  previewBtn: document.getElementById("previewBtn"),
  pasteExcelBtn: document.getElementById("pasteExcelBtn"),
  pasteServerBtn: document.getElementById("pasteServerBtn"),
  imageInput: document.getElementById("imageInput"),
  ocrBtn: document.getElementById("ocrBtn"),
  ocrToExcelBtn: document.getElementById("ocrToExcelBtn"),
  ocrToServerBtn: document.getElementById("ocrToServerBtn"),
  ocrStatus: document.getElementById("ocrStatus"),
  actionStatus: document.getElementById("actionStatus"),
  previewGrid: document.getElementById("previewGrid"),
  quickButtons: Array.from(document.querySelectorAll("[data-prompt]")),
};

function updateStatus(element, message, isError = false) {
  element.textContent = message;
  element.classList.toggle("error", isError);
}

function appendMessage(role, content) {
  const item = document.createElement("div");
  item.className = `message ${role}`;
  item.textContent = content;
  ui.messages.appendChild(item);
  ui.messages.scrollTop = ui.messages.scrollHeight;
}

function setBusy(isBusy, statusMessage = "") {
  state.isBusy = Boolean(isBusy);

  ui.chatInput.disabled = state.isBusy;
  ui.chatImageBtn.disabled = state.isBusy;
  ui.sendBtn.disabled = state.isBusy;
  ui.clearChatBtn.disabled = state.isBusy;
  ui.filepathInput.disabled = state.isBusy;
  ui.sheetInput.disabled = state.isBusy;
  ui.startCellInput.disabled = state.isBusy;
  ui.pasteInput.disabled = state.isBusy;
  ui.previewBtn.disabled = state.isBusy;
  ui.pasteExcelBtn.disabled = state.isBusy;
  ui.pasteServerBtn.disabled = state.isBusy;
  ui.ocrBtn.disabled = state.isBusy;
  ui.ocrToExcelBtn.disabled = state.isBusy;
  ui.ocrToServerBtn.disabled = state.isBusy;

  ui.quickButtons.forEach((button) => {
    button.disabled = state.isBusy;
  });

  ui.sendBtn.textContent = state.isBusy ? "Working..." : "Send";

  if (statusMessage) {
    updateStatus(ui.actionStatus, statusMessage, false);
  }
}

async function runBusyTask(statusMessage, task) {
  if (state.isBusy) {
    return null;
  }

  setBusy(true, statusMessage);
  try {
    return await task();
  } finally {
    setBusy(false);
  }
}

function normalizeRows(rows) {
  if (!rows.length) {
    return [];
  }
  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => {
    const padded = [...row];
    while (padded.length < width) {
      padded.push("");
    }
    return padded.map((cell) => {
      const token = String(cell).trim();
      if (token === "") {
        return "";
      }
      if (/^-?\d+$/.test(token)) {
        const intVal = Number.parseInt(token, 10);
        if (Number.isFinite(intVal)) {
          return intVal;
        }
      }
      if (/^-?\d+\.\d+$/.test(token)) {
        const floatVal = Number.parseFloat(token);
        if (Number.isFinite(floatVal)) {
          return floatVal;
        }
      }
      if (token.toLowerCase() === "true") {
        return true;
      }
      if (token.toLowerCase() === "false") {
        return false;
      }
      return token;
    });
  });
}

function parseTabularText(raw) {
  const text = (raw || "").replace(/\r\n/g, "\n").trim();
  if (!text) {
    return [];
  }

  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  const markdownRule = /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/;
  if (lines.length && lines.filter((line) => line.includes("|")).length >= Math.max(1, lines.length - 1)) {
    const rows = lines
      .filter((line) => !markdownRule.test(line))
      .map((line) => line.replace(/^\|/, "").replace(/\|$/, "").split("|").map((cell) => cell.trim()));
    const normalized = normalizeRows(rows);
    if (normalized.length) {
      return normalized;
    }
  }

  if (text.includes("\t")) {
    return normalizeRows(lines.map((line) => line.split("\t").map((cell) => cell.trim())));
  }

  const csvRows = lines.map((line) => line.split(/\s*,\s*/));
  if (csvRows.some((row) => row.length > 1)) {
    return normalizeRows(csvRows);
  }

  const spacesRows = lines.map((line) => line.split(/\s{2,}/));
  if (spacesRows.some((row) => row.length > 1)) {
    return normalizeRows(spacesRows);
  }

  return normalizeRows(lines.map((line) => [line]));
}

function renderPreview(rows) {
  ui.previewGrid.innerHTML = "";
  if (!rows.length) {
    updateStatus(ui.actionStatus, "No tabular rows detected.", true);
    return;
  }

  const table = document.createElement("table");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = String(cell);
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  ui.previewGrid.appendChild(table);
  updateStatus(ui.actionStatus, `Preview ready: ${rows.length} rows x ${rows[0].length} cols`);
}

function getWorkbookConfig() {
  return {
    filepath: ui.filepathInput.value.trim(),
    sheet_name: ui.sheetInput.value.trim(),
    start_cell: ui.startCellInput.value.trim() || "A1",
  };
}

function toRectangularData(rawData) {
  if (!Array.isArray(rawData)) {
    return [];
  }

  if (!rawData.length) {
    return [];
  }

  const rows = rawData.map((row) => (Array.isArray(row) ? [...row] : [row]));
  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => {
    const padded = [...row];
    while (padded.length < width) {
      padded.push("");
    }
    return padded;
  });
}

function stripSheetQualifier(address) {
  const raw = String(address || "").trim();
  if (!raw) {
    return "";
  }
  const withoutSheet = raw.includes("!") ? raw.split("!").pop() : raw;
  return String(withoutSheet || "").replace(/\$/g, "").trim();
}

function resolveRangeAddress(args) {
  if (typeof args.range === "string" && args.range.trim()) {
    return stripSheetQualifier(args.range);
  }

  const start = typeof args.start_cell === "string" && args.start_cell.trim()
    ? stripSheetQualifier(args.start_cell)
    : "A1";
  const end = typeof args.end_cell === "string" && args.end_cell.trim()
    ? stripSheetQualifier(args.end_cell)
    : "";
  return end ? `${start}:${end}` : start;
}

function getRangeStartCell(rangeAddress) {
  if (typeof rangeAddress !== "string" || !rangeAddress.trim()) {
    return "A1";
  }
  const rawStart = rangeAddress.split(":")[0].trim();
  const normalized = stripSheetQualifier(rawStart);
  return normalized || "A1";
}

function normalizeHexColor(value) {
  if (typeof value !== "string" || !value.trim()) {
    return "";
  }
  const trimmed = value.trim();
  if (trimmed.startsWith("#")) {
    return trimmed;
  }
  if (/^[0-9a-fA-F]{6}$/.test(trimmed)) {
    return `#${trimmed}`;
  }
  return trimmed;
}

function splitCellAddress(cellAddress) {
  const match = String(cellAddress || "A1").replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    return { column: "A", row: 1 };
  }
  return { column: match[1].toUpperCase(), row: Number.parseInt(match[2], 10) };
}

function columnNameToNumber(columnName) {
  return String(columnName || "A")
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
}

function columnNumberToName(columnNumber) {
  let value = Math.max(1, Number.parseInt(columnNumber, 10) || 1);
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function offsetCellAddress(cellAddress, rowOffset, columnOffset) {
  const parsed = splitCellAddress(cellAddress);
  const col = columnNameToNumber(parsed.column) + Number(columnOffset || 0);
  const row = parsed.row + Number(rowOffset || 0);
  return `${columnNumberToName(col)}${Math.max(1, row)}`;
}

function normalizeChartType(chartType) {
  const key = String(chartType || "bar").trim().toLowerCase();
  const map = {
    area: "Area",
    bar: "BarClustered",
    column: "ColumnClustered",
    line: "Line",
    pie: "Pie",
    scatter: "XYScatter",
  };
  return map[key] || map.bar;
}

function makeTableName(rawName) {
  const cleaned = String(rawName || `SmartTable${Date.now()}`)
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^[^A-Za-z_]+/, "");
  return cleaned || `SmartTable${Date.now()}`;
}

function makeSheetName(rawName, fallbackPrefix = "Smart") {
  const cleaned = String(rawName || `${fallbackPrefix}_${Date.now()}`)
    .replace(/[\\/?*\[\]:]/g, " ")
    .trim()
    .slice(0, 31);
  return cleaned || `${fallbackPrefix}_${Date.now()}`.slice(0, 31);
}

function normalizeHeaderText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ");
}

function promptIncludesAll(prompt, tokens) {
  return tokens.every((token) => prompt.includes(token));
}

function isNonEmptyCell(value) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function matchesStudentSampleIntent(prompt) {
  return (
    prompt.includes("student")
    && (prompt.includes("sample data") || prompt.includes("dataset") || prompt.includes("table"))
    && (prompt.includes("roll") || prompt.includes("rollnumber") || prompt.includes("roll number"))
    && prompt.includes("mark")
  );
}

function matchesStudentFormulaIntent(prompt) {
  return (
    promptIncludesAll(prompt, ["useful", "formula"])
    || promptIncludesAll(prompt, ["add", "formula"])
    || promptIncludesAll(prompt, ["total", "average"])
    || promptIncludesAll(prompt, ["percentage", "grade"])
  );
}

function getSampleStudentRows() {
  return [
    ["Name", "Roll Number", "Math", "Science", "English", "History", "Total", "Average", "Percentage", "Grade"],
    ["John Smith", 101, 85, 92, 78, 88, "", "", "", ""],
    ["Emma Johnson", 102, 92, 88, 95, 90, "", "", "", ""],
    ["Michael Brown", 103, 78, 85, 82, 79, "", "", "", ""],
    ["Sarah Davis", 104, 95, 90, 88, 92, "", "", "", ""],
    ["David Wilson", 105, 82, 79, 85, 80, "", "", "", ""],
    ["Lisa Miller", 106, 90, 95, 92, 88, "", "", "", ""],
    ["Robert Taylor", 107, 75, 82, 78, 85, "", "", "", ""],
    ["Jennifer Lee", 108, 88, 85, 90, 87, "", "", "", ""],
    ["William Clark", 109, 92, 88, 85, 91, "", "", "", ""],
    ["Amanda White", 110, 85, 90, 92, 89, "", "", "", ""],
  ];
}

function findHeaderIndex(headers, aliases) {
  const normalizedAliases = aliases.map((alias) => normalizeHeaderText(alias));
  return headers.findIndex((header) => normalizedAliases.includes(normalizeHeaderText(header)));
}

function findStudentMarksLayout(values) {
  const rows = toRectangularData(values);
  if (rows.length < 2) {
    return null;
  }

  const headers = rows[0].map((value) => String(value ?? "").trim());
  const nameIndex = findHeaderIndex(headers, ["name", "student name"]);
  const rollIndex = findHeaderIndex(headers, ["roll number", "rollnumber", "roll no", "roll"]);
  if (nameIndex < 0 || rollIndex < 0) {
    return null;
  }

  const metricNames = new Set(["total", "average", "avg", "percentage", "percent", "grade"]);
  const subjectIndices = headers
    .map((header, index) => ({ header, index }))
    .filter(({ header, index }) => {
      const normalized = normalizeHeaderText(header);
      if (!normalized || metricNames.has(normalized)) {
        return false;
      }
      return index !== nameIndex && index !== rollIndex;
    })
    .map(({ index }) => index);

  if (subjectIndices.length < 2) {
    return null;
  }

  let dataRowStart = 1;
  let dataRowEnd = 0;
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const hasName = isNonEmptyCell(row[nameIndex]);
    const hasRoll = isNonEmptyCell(row[rollIndex]);
    const hasSubjectData = subjectIndices.some((index) => isNonEmptyCell(row[index]));

    if (hasName || hasRoll || hasSubjectData) {
      dataRowEnd = rowIndex;
      continue;
    }

    if (dataRowEnd >= dataRowStart) {
      break;
    }
  }

  if (dataRowEnd < dataRowStart) {
    return null;
  }

  return {
    headers,
    nameIndex,
    rollIndex,
    subjectIndices,
    dataRowStart,
    dataRowEnd,
    summaryLabelRow: rows.findIndex((row, index) => index > dataRowEnd && normalizeHeaderText(row[0]) === "summary statistics"),
    metrics: {
      total: findHeaderIndex(headers, ["total"]),
      average: findHeaderIndex(headers, ["average", "avg"]),
      percentage: findHeaderIndex(headers, ["percentage", "percent"]),
      grade: findHeaderIndex(headers, ["grade"]),
    },
  };
}

async function createSampleStudentDataset() {
  const rows = getSampleStudentRows();

  return Excel.run(async (context) => {
    const sheet = getActiveWorksheet(context);
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("isNullObject");
    await context.sync();

    if (!usedRange.isNullObject) {
      usedRange.clear();
    }

    const target = sheet.getRange("A1").getResizedRange(rows.length - 1, rows[0].length - 1);
    target.values = rows;

    const headerRange = sheet.getRange(`A1:${columnNumberToName(rows[0].length)}1`);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#1F3330";
    headerRange.format.font.color = "#F3F7F5";

    sheet.getRange(`B2:B${rows.length}`).numberFormat = Array.from({ length: rows.length - 1 }, () => ["0"]);
    target.format.autofitColumns();
    target.format.autofitRows();
    await context.sync();

    return {
      reply: "Created a clean 10-row student marks table with Name, Roll Number, four subjects, and placeholder columns for Total, Average, Percentage, and Grade.",
      status: "Sample student table written to the active sheet.",
    };
  });
}

async function addStudentDerivedFormulas() {
  return Excel.run(async (context) => {
    const sheet = getActiveWorksheet(context);
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("isNullObject,values,rowCount,columnCount");
    await context.sync();

    if (usedRange.isNullObject) {
      return null;
    }

    const layout = findStudentMarksLayout(usedRange.values);
    if (!layout) {
      return null;
    }

    const headerRow = 1;
    const firstDataRow = layout.dataRowStart + 1;
    const lastDataRow = layout.dataRowEnd + 1;
    const subjectStartCol = columnNumberToName(layout.subjectIndices[0] + 1);
    const subjectEndCol = columnNumberToName(layout.subjectIndices[layout.subjectIndices.length - 1] + 1);

    const metricIndexes = { ...layout.metrics };
    let nextColumnIndex = Math.max(...layout.subjectIndices, layout.rollIndex, layout.nameIndex) + 1;
    ["total", "average", "percentage", "grade"].forEach((metricName) => {
      if (metricIndexes[metricName] < 0) {
        metricIndexes[metricName] = nextColumnIndex;
        nextColumnIndex += 1;
      }
    });

    const metricHeaders = {
      total: "Total",
      average: "Average",
      percentage: "Percentage",
      grade: "Grade",
    };
    Object.entries(metricIndexes).forEach(([metricName, columnIndex]) => {
      sheet.getCell(headerRow - 1, columnIndex).values = [[metricHeaders[metricName]]];
    });

    const totalCol = columnNumberToName(metricIndexes.total + 1);
    const averageCol = columnNumberToName(metricIndexes.average + 1);
    const percentageCol = columnNumberToName(metricIndexes.percentage + 1);
    const gradeCol = columnNumberToName(metricIndexes.grade + 1);
    const subjectCount = layout.subjectIndices.length;
    const formulaRowCount = lastDataRow - firstDataRow + 1;

    const totalFormulas = [];
    const averageFormulas = [];
    const percentageFormulas = [];
    const gradeFormulas = [];

    for (let excelRow = firstDataRow; excelRow <= lastDataRow; excelRow += 1) {
      totalFormulas.push([`=SUM(${subjectStartCol}${excelRow}:${subjectEndCol}${excelRow})`]);
      averageFormulas.push([`=ROUND(AVERAGE(${subjectStartCol}${excelRow}:${subjectEndCol}${excelRow}),2)`]);
      percentageFormulas.push([`=ROUND(${totalCol}${excelRow}/(${subjectCount}*100)*100,2)`]);
      gradeFormulas.push([`=IF(${percentageCol}${excelRow}>=90,"A+",IF(${percentageCol}${excelRow}>=80,"A",IF(${percentageCol}${excelRow}>=70,"B",IF(${percentageCol}${excelRow}>=60,"C",IF(${percentageCol}${excelRow}>=50,"D","F")))))`]);
    }

    sheet.getRange(`${totalCol}${firstDataRow}:${totalCol}${lastDataRow}`).formulas = totalFormulas;
    sheet.getRange(`${averageCol}${firstDataRow}:${averageCol}${lastDataRow}`).formulas = averageFormulas;
    sheet.getRange(`${percentageCol}${firstDataRow}:${percentageCol}${lastDataRow}`).formulas = percentageFormulas;
    sheet.getRange(`${gradeCol}${firstDataRow}:${gradeCol}${lastDataRow}`).formulas = gradeFormulas;
    sheet.getRange(`${percentageCol}${firstDataRow}:${percentageCol}${lastDataRow}`).numberFormat = Array.from(
      { length: formulaRowCount },
      () => ["0.00"],
    );

    const headerEndCol = columnNumberToName(Math.max(...Object.values(metricIndexes)) + 1);
    const headerRange = sheet.getRange(`A1:${headerEndCol}1`);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#1F3330";
    headerRange.format.font.color = "#F3F7F5";

    let cleanedSummary = false;
    if (layout.summaryLabelRow >= 0) {
      const summaryStartRow = layout.summaryLabelRow + 1;
      const summaryEndRow = Math.min(summaryStartRow + 8, usedRange.rowCount);
      sheet.getRange(`A${summaryStartRow}:${headerEndCol}${summaryEndRow}`).clear();
      cleanedSummary = true;
    }

    sheet.getRange(`A1:${headerEndCol}${lastDataRow}`).format.autofitColumns();
    sheet.getRange(`A1:${headerEndCol}${lastDataRow}`).format.autofitRows();
    await context.sync();

    return {
      reply: cleanedSummary
        ? `Filled Total, Average, Percentage, and Grade for each student, and removed the earlier summary block that was cluttering the sheet.`
        : `Filled Total, Average, Percentage, and Grade for each student in the next logical columns.`,
      status: `Added formulas to ${totalCol}:${gradeCol} for rows ${firstDataRow}:${lastDataRow}.`,
    };
  });
}

async function tryHandleLocalCopilotIntent(prompt) {
  if (!state.isExcelHost) {
    return null;
  }

  const normalizedPrompt = String(prompt || "").trim().toLowerCase();
  const shouldCreateStudentData = matchesStudentSampleIntent(normalizedPrompt);
  const shouldAddStudentFormulas = matchesStudentFormulaIntent(normalizedPrompt);

  if (!shouldCreateStudentData && !shouldAddStudentFormulas) {
    return null;
  }

  const replies = [];
  let lastStatus = "Applied changes to active workbook.";

  if (shouldCreateStudentData) {
    const datasetResult = await createSampleStudentDataset();
    if (datasetResult) {
      replies.push(datasetResult.reply);
      lastStatus = datasetResult.status || lastStatus;
    }
  }

  if (shouldAddStudentFormulas) {
    const formulaResult = await addStudentDerivedFormulas();
    if (formulaResult) {
      replies.push(formulaResult.reply);
      lastStatus = formulaResult.status || lastStatus;
    }
  }

  if (!replies.length) {
    return null;
  }

  return {
    reply: replies.join("\n\n"),
    status: lastStatus,
  };
}

function getColumnIndex(headers, fieldName) {
  const target = String(fieldName || "").trim().toLowerCase();
  return headers.findIndex((header) => String(header || "").trim().toLowerCase() === target);
}

function aggregateValues(values, aggFunc) {
  const numeric = values
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
  const mode = String(aggFunc || "sum").toLowerCase();
  if (mode === "count") {
    return values.filter((value) => value !== null && value !== undefined && String(value).trim() !== "").length;
  }
  if (!numeric.length) {
    return 0;
  }
  if (mode === "average" || mode === "avg" || mode === "mean") {
    return numeric.reduce((sum, value) => sum + value, 0) / numeric.length;
  }
  if (mode === "min") {
    return Math.min(...numeric);
  }
  if (mode === "max") {
    return Math.max(...numeric);
  }
  return numeric.reduce((sum, value) => sum + value, 0);
}

async function getOrCreateWorksheet(context, sheetName) {
  const name = makeSheetName(sheetName);
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load("isNullObject");
  await context.sync();
  if (!sheet.isNullObject) {
    return sheet;
  }
  return context.workbook.worksheets.add(name);
}

async function writePivotSummary(context, sourceSheet, args) {
  const sourceRange = sourceSheet.getRange(args.data_range || args.range || "A1");
  sourceRange.load("values");
  await context.sync();

  const sourceValues = toRectangularData(sourceRange.values);
  if (sourceValues.length < 2) {
    throw new Error("Pivot summary needs a header row and at least one data row.");
  }

  const headers = sourceValues[0].map((header) => String(header || "").trim());
  const rowFields = args.rows || args.row_fields || [];
  const valueFields = args.values || args.value_fields || [];
  const aggFunc = args.agg_func || "sum";
  const rowIndexes = rowFields.map((field) => getColumnIndex(headers, field));
  const valueIndexes = valueFields.map((field) => getColumnIndex(headers, field));

  if (rowIndexes.some((index) => index < 0) || valueIndexes.some((index) => index < 0)) {
    throw new Error("Pivot summary field names must match source headers.");
  }

  const buckets = new Map();
  sourceValues.slice(1).forEach((row) => {
    const keyParts = rowIndexes.map((index) => row[index] ?? "");
    const key = JSON.stringify(keyParts);
    if (!buckets.has(key)) {
      buckets.set(key, { keyParts, values: valueIndexes.map(() => []) });
    }
    const bucket = buckets.get(key);
    valueIndexes.forEach((index, valueIndex) => {
      bucket.values[valueIndex].push(row[index]);
    });
  });

  const output = [
    [
      ...rowFields,
      ...valueFields.map((field) => `${field} (${aggFunc})`),
    ],
  ];
  [...buckets.values()].forEach((bucket) => {
    output.push([
      ...bucket.keyParts,
      ...bucket.values.map((items) => aggregateValues(items, aggFunc)),
    ]);
  });

  const outputSheet = await getOrCreateWorksheet(context, args.output_sheet_name || "Smart Summary");
  const target = outputSheet.getRange(args.output_start_cell || "A1").getResizedRange(output.length - 1, output[0].length - 1);
  target.values = output;
  target.format.autofitColumns();
  target.format.autofitRows();
  target.getRow(0).format.font.bold = true;
  return output.length - 1;
}

function getActiveWorksheet(context) {
  return context.workbook.worksheets.getActiveWorksheet();
}

async function resolveWorksheet(context, sheetName) {
  if (!sheetName) {
    return getActiveWorksheet(context);
  }

  const candidate = context.workbook.worksheets.getItemOrNullObject(String(sheetName));
  candidate.load("isNullObject");
  await context.sync();

  if (candidate.isNullObject) {
    return getActiveWorksheet(context);
  }

  return candidate;
}

async function collectExcelSheetSnapshot() {
  if (!state.isExcelHost || !window.Excel || !window.Office) {
    return null;
  }

  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeSheet = workbook.worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();
    const usedRange = activeSheet.getUsedRangeOrNullObject(true);

    activeSheet.load("name");
    worksheets.load("items/name");
    selectedRange.load("address,rowCount,columnCount");
    usedRange.load("isNullObject,address,rowCount,columnCount");
    await context.sync();

    const selectionAddress = selectedRange.address || "A1";
    const selectionStartCell = getRangeStartCell(selectionAddress);
    const selectionPreviewRows = Math.max(1, Math.min(Number(selectedRange.rowCount || 1), SHEET_SELECTION_MAX_ROWS));
    const selectionPreviewCols = Math.max(1, Math.min(Number(selectedRange.columnCount || 1), SHEET_SELECTION_MAX_COLS));
    const selectionPreviewRange = selectedRange
      .getCell(0, 0)
      .getResizedRange(selectionPreviewRows - 1, selectionPreviewCols - 1);

    selectionPreviewRange.load("address,rowCount,columnCount,values,formulas");

    let usedPreviewRange = null;
    if (!usedRange.isNullObject) {
      const previewRows = Math.max(1, Math.min(Number(usedRange.rowCount || 1), SHEET_SNAPSHOT_MAX_ROWS));
      const previewCols = Math.max(1, Math.min(Number(usedRange.columnCount || 1), SHEET_SNAPSHOT_MAX_COLS));
      usedPreviewRange = usedRange
        .getCell(0, 0)
        .getResizedRange(previewRows - 1, previewCols - 1);
      usedPreviewRange.load("address,rowCount,columnCount,values,formulas");
    }

    await context.sync();

    const sheetNames = Array.isArray(worksheets.items)
      ? worksheets.items.map((sheet) => String(sheet.name || "")).filter(Boolean)
      : [];

    if (usedRange.isNullObject) {
      return {
        sheet_name: activeSheet.name,
        sheet_names: sheetNames,
        selection_address: selectionAddress,
        selection_start_cell: selectionStartCell,
        selection_preview_address: selectionPreviewRange.address,
        selection_preview_row_count: selectionPreviewRange.rowCount,
        selection_preview_column_count: selectionPreviewRange.columnCount,
        selection_values: selectionPreviewRange.values,
        selection_formulas: selectionPreviewRange.formulas,
        used_range_address: null,
        preview_address: null,
        snapshot_mode: "empty",
        row_count: 0,
        column_count: 0,
        values: [],
        formulas: [],
      };
    }

    return {
      sheet_name: activeSheet.name,
      sheet_names: sheetNames,
      selection_address: selectionAddress,
      selection_start_cell: selectionStartCell,
      selection_preview_address: selectionPreviewRange.address,
      selection_preview_row_count: selectionPreviewRange.rowCount,
      selection_preview_column_count: selectionPreviewRange.columnCount,
      selection_values: selectionPreviewRange.values,
      selection_formulas: selectionPreviewRange.formulas,
      used_range_address: usedRange.address,
      preview_address: usedPreviewRange ? usedPreviewRange.address : usedRange.address,
      snapshot_mode: (
        Number(usedRange.rowCount || 0) > SHEET_SNAPSHOT_MAX_ROWS
        || Number(usedRange.columnCount || 0) > SHEET_SNAPSHOT_MAX_COLS
      )
        ? "bounded"
        : "full",
      row_count: usedRange.rowCount,
      column_count: usedRange.columnCount,
      values: usedPreviewRange ? usedPreviewRange.values : [],
      formulas: usedPreviewRange ? usedPreviewRange.formulas : [],
    };
  });
}

async function executeOperationsInExcel(operations) {
  if (!state.isExcelHost || !window.Excel || !window.Office) {
    return {
      executed: 0,
      warnings: ["Excel host is unavailable for direct operation execution."],
    };
  }

  const safeOps = Array.isArray(operations) ? operations : [];
  const warnings = [];
  let executed = 0;

  await Excel.run(async (context) => {
    for (const op of safeOps) {
      const opName = String(op && op.name ? op.name : "").trim().toLowerCase();
      const args = op && typeof op.args === "object" && op.args !== null ? op.args : {};
      const sheet = await resolveWorksheet(context, args.sheet_name);

      if (opName === "write_data") {
        const rows = toRectangularData(args.data);
        if (!rows.length) {
          sheet.getRange(resolveRangeAddress(args)).clear();
          executed += 1;
          continue;
        }

        const rangeAddress = typeof args.range === "string" ? args.range : "";
        const startCell = typeof args.start_cell === "string" && args.start_cell.trim()
          ? args.start_cell.trim()
          : getRangeStartCell(rangeAddress);
        const target = sheet
          .getRange(startCell)
          .getResizedRange(rows.length - 1, rows[0].length - 1);
        target.values = rows;
        target.format.autofitColumns();
        target.format.autofitRows();
        executed += 1;
        continue;
      }

      if (opName === "read_data") {
        const rangeAddress = resolveRangeAddress(args);
        const range = sheet.getRange(rangeAddress);
        range.load("values");
        await context.sync();
        executed += 1;
        continue;
      }

      if (opName === "get_workbook_metadata") {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        const used = activeSheet.getUsedRangeOrNullObject(true);
        used.load("isNullObject,address,rowCount,columnCount");
        await context.sync();
        executed += 1;
        continue;
      }

      if (opName === "apply_formula") {
        const formula = typeof args.formula === "string" ? args.formula : "";
        if (!formula) {
          warnings.push("Skipped apply_formula because formula is missing.");
          continue;
        }
        const cell = typeof args.cell === "string" && args.cell.trim()
          ? args.cell.trim()
          : (typeof args.start_cell === "string" && args.start_cell.trim() ? args.start_cell.trim() : "A1");
        sheet.getRange(cell).formulas = [[formula]];
        executed += 1;
        continue;
      }

      if (opName === "format_range") {
        const range = sheet.getRange(resolveRangeAddress(args));

        if (typeof args.bold === "boolean") {
          range.format.font.bold = args.bold;
        }
        if (typeof args.italic === "boolean") {
          range.format.font.italic = args.italic;
        }
        if (typeof args.underline === "boolean") {
          range.format.font.underline = args.underline ? "Single" : "None";
        }
        if (Number.isFinite(Number(args.font_size))) {
          range.format.font.size = Number(args.font_size);
        }
        if (typeof args.font_color === "string" && args.font_color.trim()) {
          range.format.font.color = normalizeHexColor(args.font_color);
        }
        if (typeof args.bg_color === "string" && args.bg_color.trim()) {
          range.format.fill.color = normalizeHexColor(args.bg_color);
        }
        if (typeof args.alignment === "string" && args.alignment.trim()) {
          const alignment = args.alignment.trim().toLowerCase();
          const map = {
            center: "Center",
            left: "Left",
            right: "Right",
            justify: "Justify",
          };
          range.format.horizontalAlignment = map[alignment] || args.alignment;
        }
        if (typeof args.wrap_text === "boolean") {
          range.format.wrapText = args.wrap_text;
        }
        if (typeof args.number_format === "string" && args.number_format.trim()) {
          range.load("rowCount,columnCount");
          await context.sync();
          range.numberFormat = Array.from({ length: range.rowCount }, () =>
            Array.from({ length: range.columnCount }, () => args.number_format)
          );
        }
        range.format.autofitColumns();
        range.format.autofitRows();
        executed += 1;
        continue;
      }

      if (opName === "clear_range") {
        sheet.getRange(resolveRangeAddress(args)).clear();
        executed += 1;
        continue;
      }

      if (opName === "create_worksheet") {
        const name = typeof args.sheet_name === "string" && args.sheet_name.trim()
          ? args.sheet_name.trim()
          : `Sheet${Date.now()}`;
        context.workbook.worksheets.add(name);
        executed += 1;
        continue;
      }

      if (opName === "rename_worksheet") {
        if (!args.old_name || !args.new_name) {
          warnings.push("Skipped rename_worksheet because old_name/new_name is missing.");
          continue;
        }
        const oldSheet = context.workbook.worksheets.getItemOrNullObject(String(args.old_name));
        oldSheet.load("isNullObject");
        await context.sync();
        if (oldSheet.isNullObject) {
          warnings.push(`Skipped rename_worksheet: sheet ${args.old_name} was not found.`);
          continue;
        }
        oldSheet.name = String(args.new_name);
        executed += 1;
        continue;
      }

      if (opName === "delete_worksheet") {
        if (!args.sheet_name) {
          warnings.push("Skipped delete_worksheet because sheet_name is missing.");
          continue;
        }
        const toDelete = context.workbook.worksheets.getItemOrNullObject(String(args.sheet_name));
        toDelete.load("isNullObject");
        await context.sync();
        if (toDelete.isNullObject) {
          warnings.push(`Skipped delete_worksheet: sheet ${args.sheet_name} was not found.`);
          continue;
        }
        toDelete.delete();
        executed += 1;
        continue;
      }

      if (opName === "insert_rows") {
        const startRow = Math.max(1, Number.parseInt(args.start_row || 1, 10));
        const count = Math.max(1, Number.parseInt(args.count || 1, 10));
        sheet.getRangeByIndexes(startRow - 1, 0, count, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
        executed += 1;
        continue;
      }

      if (opName === "insert_columns") {
        const startCol = Math.max(1, Number.parseInt(args.start_col || 1, 10));
        const count = Math.max(1, Number.parseInt(args.count || 1, 10));
        sheet.getRangeByIndexes(0, startCol - 1, 1, count).getEntireColumn().insert(Excel.InsertShiftDirection.right);
        executed += 1;
        continue;
      }

      if (opName === "delete_sheet_rows") {
        const startRow = Math.max(1, Number.parseInt(args.start_row || 1, 10));
        const count = Math.max(1, Number.parseInt(args.count || 1, 10));
        sheet.getRangeByIndexes(startRow - 1, 0, count, 1).getEntireRow().delete(Excel.DeleteShiftDirection.up);
        executed += 1;
        continue;
      }

      if (opName === "delete_sheet_columns") {
        const startCol = Math.max(1, Number.parseInt(args.start_col || 1, 10));
        const count = Math.max(1, Number.parseInt(args.count || 1, 10));
        sheet.getRangeByIndexes(0, startCol - 1, 1, count).getEntireColumn().delete(Excel.DeleteShiftDirection.left);
        executed += 1;
        continue;
      }

      if (opName === "merge_cells") {
        sheet.getRange(resolveRangeAddress(args)).merge(false);
        executed += 1;
        continue;
      }

      if (opName === "unmerge_cells") {
        sheet.getRange(resolveRangeAddress(args)).unmerge();
        executed += 1;
        continue;
      }

      if (opName === "copy_range") {
        const sourceAddress = args.source_end ? `${args.source_start}:${args.source_end}` : args.source_start;
        const source = sheet.getRange(sourceAddress);
        const targetSheet = await resolveWorksheet(context, args.target_sheet || args.sheet_name);
        const target = targetSheet.getRange(args.target_start || "A1");
        target.copyFrom(source, Excel.RangeCopyType.all, false, false);
        executed += 1;
        continue;
      }

      if (opName === "delete_range") {
        const shift = String(args.shift_direction || "up").toLowerCase() === "left"
          ? Excel.DeleteShiftDirection.left
          : Excel.DeleteShiftDirection.up;
        sheet.getRange(resolveRangeAddress(args)).delete(shift);
        executed += 1;
        continue;
      }

      if (opName === "create_table") {
        const rangeAddress = args.data_range || args.range || resolveRangeAddress(args);
        const table = context.workbook.tables.add(sheet.getRange(rangeAddress), true);
        table.name = makeTableName(args.table_name);
        table.style = args.table_style || "TableStyleMedium9";
        sheet.getRange(rangeAddress).format.autofitColumns();
        executed += 1;
        continue;
      }

      if (opName === "create_chart") {
        const rangeAddress = args.data_range || args.range || resolveRangeAddress(args);
        const dataRange = sheet.getRange(rangeAddress);
        const chartTypeMap = {
          area: Excel.ChartType.area,
          bar: Excel.ChartType.barClustered,
          column: Excel.ChartType.columnClustered,
          line: Excel.ChartType.line,
          pie: Excel.ChartType.pie,
          scatter: Excel.ChartType.xyScatter,
        };
        const chartKey = String(args.chart_type || "bar").toLowerCase();
        const chart = sheet.charts.add(chartTypeMap[chartKey] || chartTypeMap.bar || normalizeChartType(args.chart_type), dataRange, Excel.ChartSeriesBy.auto);
        if (args.title) {
          chart.title.text = String(args.title);
          chart.title.visible = true;
        }
        if (args.x_axis && chart.axes && chart.axes.categoryAxis) {
          chart.axes.categoryAxis.title.text = String(args.x_axis);
        }
        if (args.y_axis && chart.axes && chart.axes.valueAxis) {
          chart.axes.valueAxis.title.text = String(args.y_axis);
        }
        const topLeft = args.target_cell || "H2";
        chart.setPosition(topLeft, offsetCellAddress(topLeft, 16, 7));
        executed += 1;
        continue;
      }

      if (opName === "create_pivot_table") {
        const rowsWritten = await writePivotSummary(context, sheet, args);
        warnings.push(`Created summary sheet with ${rowsWritten} grouped row(s).`);
        executed += 1;
        continue;
      }

      warnings.push(`Skipped unsupported operation: ${opName || "unknown"}.`);
    }

    await context.sync();
  });

  return { executed, warnings };
}

async function sendChat() {
  const prompt = ui.chatInput.value.trim();
  if (!prompt || state.isBusy) {
    return;
  }

  return runBusyTask("Reading the active sheet...", async () => {
    appendMessage("user", prompt);
    state.messages.push({ role: "user", content: prompt });
    ui.chatInput.value = "";

    try {
      const config = getWorkbookConfig();
      let sheetSnapshot = null;

      if (state.isExcelHost) {
        updateStatus(ui.actionStatus, "Reading active sheet context...");
        try {
          sheetSnapshot = await collectExcelSheetSnapshot();
          state.lastSheetSnapshot = sheetSnapshot;
        } catch (error) {
          throw new Error(`Could not read the active sheet before sending the command: ${error.message}`);
        }
      }

      const localResult = await tryHandleLocalCopilotIntent(prompt);
      if (localResult) {
        appendMessage("assistant", localResult.reply);
        state.messages.push({ role: "assistant", content: localResult.reply });
        updateStatus(ui.actionStatus, localResult.status || "Applied changes to active workbook.");
        return;
      }

      const effectiveSheetName =
        sheetSnapshot && sheetSnapshot.sheet_name
          ? String(sheetSnapshot.sheet_name)
          : (config.sheet_name || null);

      const effectiveStartCell =
        sheetSnapshot && sheetSnapshot.selection_start_cell
          ? String(sheetSnapshot.selection_start_cell)
          : config.start_cell;

      const payload = {
        messages: state.messages,
        filepath: config.filepath || null,
        sheet_name: effectiveSheetName,
        start_cell: effectiveStartCell,
        sheet_snapshot: sheetSnapshot,
        image_context: state.latestImageContext,
        auto_execute: !state.isExcelHost,
      };

      updateStatus(ui.actionStatus, "Thinking with the latest sheet context...");
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        let errorText = "";
        try {
          const errorPayload = await response.json();
          errorText = errorPayload && errorPayload.detail ? String(errorPayload.detail) : JSON.stringify(errorPayload);
        } catch (_err) {
          errorText = await response.text();
        }
        throw new Error(errorText || `Chat request failed (${response.status})`);
      }

      const data = await response.json();
      let executionSuffix = "";
      if (Array.isArray(data.operations) && data.operations.length) {
        if (state.isExcelHost) {
          const localExecution = await executeOperationsInExcel(data.operations);
          executionSuffix = `\n\nApplied ${localExecution.executed} operation(s) to active workbook.`;
          if (localExecution.warnings.length) {
            executionSuffix += `\n${localExecution.warnings.slice(0, 2).join(" ")}`;
          }
        } else if (Array.isArray(data.operation_results) && data.operation_results.length) {
          executionSuffix = `\n\nExecuted ${data.operation_results.length} server operation(s).`;
        }
      }

      if (Array.isArray(data.operation_validation_warnings) && data.operation_validation_warnings.length) {
        executionSuffix += `\nValidation warnings: ${data.operation_validation_warnings.slice(0, 2).join(" ")}`;
      }

      if (data.sheet_context_mode === "compact-fallback") {
        executionSuffix += "\nSheet context fallback: full-sheet snapshot timed out, compact context used.";
      }

      appendMessage("assistant", `${data.reply || "No response"}${executionSuffix}`);
      state.messages.push({ role: "assistant", content: data.reply || "" });

      updateStatus(ui.actionStatus, state.isExcelHost ? "Applied changes to active workbook." : `Model: ${data.model}`);
    } catch (error) {
      appendMessage("system", `Error: ${error.message}`);
      updateStatus(ui.actionStatus, `Chat error: ${error.message}`, true);
    }
  });
}

async function pasteToActiveSelection(rows) {
  if (!rows.length) {
    throw new Error("No parsed rows to paste");
  }

  if (!window.Excel || !window.Office) {
    throw new Error("Office.js is unavailable. Open this UI inside Excel to paste directly.");
  }

  await Excel.run(async (context) => {
    const anchor = context.workbook.getSelectedRange();
    const target = anchor.getResizedRange(rows.length - 1, rows[0].length - 1);
    target.values = rows;
    target.format.autofitColumns();
    target.format.autofitRows();
    await context.sync();
  });
}

async function pasteToServerWorkbook(rows) {
  const config = getWorkbookConfig();
  if (!config.filepath || !config.sheet_name) {
    throw new Error("Workbook file and sheet name are required for server paste");
  }

  const normalizedRows = Array.isArray(rows) && rows.length ? rows : parseTabularText(ui.pasteInput.value);
  const text = rowsToTsv(normalizedRows) || ui.pasteInput.value;
  const response = await fetch("/api/paste-text", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      filepath: config.filepath,
      sheet_name: config.sheet_name,
      start_cell: config.start_cell,
      text,
    }),
  });

  if (!response.ok) {
    const payload = await response.text();
    throw new Error(payload || "Server paste failed");
  }

  return response.json();
}

function pickImageFile() {
  return new Promise((resolve, reject) => {
    const handler = () => {
      ui.imageInput.removeEventListener("change", handler);
      const [file] = ui.imageInput.files || [];
      if (!file) {
        reject(new Error("No image selected"));
        return;
      }
      resolve(file);
    };
    ui.imageInput.value = "";
    ui.imageInput.addEventListener("change", handler, { once: true });
    ui.imageInput.click();
  });
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Failed to read selected image"));
    reader.readAsDataURL(file);
  });
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to load image for OCR processing."));
    image.src = dataUrl;
  });
}

async function optimizeImageForOcr(dataUrl) {
  const image = await loadImageFromDataUrl(dataUrl);
  const width = Number(image.naturalWidth || image.width || 0);
  const height = Number(image.naturalHeight || image.height || 0);

  if (!width || !height) {
    return {
      imageBase64: dataUrl,
      width: 0,
      height: 0,
      scaled: false,
    };
  }

  const maxEdge = Math.max(width, height);
  if (maxEdge <= OCR_MAX_IMAGE_EDGE) {
    return {
      imageBase64: dataUrl,
      width,
      height,
      scaled: false,
    };
  }

  const scale = OCR_MAX_IMAGE_EDGE / maxEdge;
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));

  const canvas = document.createElement("canvas");
  canvas.width = targetWidth;
  canvas.height = targetHeight;

  const context = canvas.getContext("2d", { alpha: false });
  if (!context) {
    return {
      imageBase64: dataUrl,
      width,
      height,
      scaled: false,
    };
  }

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, targetWidth, targetHeight);
  context.drawImage(image, 0, 0, targetWidth, targetHeight);

  return {
    imageBase64: canvas.toDataURL("image/png"),
    width: targetWidth,
    height: targetHeight,
    scaled: true,
    originalWidth: width,
    originalHeight: height,
  };
}

async function fetchJsonWithTimeout(url, init, timeoutMs) {
  const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
  const timeoutHandle = controller
    ? setTimeout(() => controller.abort(), timeoutMs)
    : null;

  try {
    const response = await fetch(url, {
      ...init,
      signal: controller ? controller.signal : undefined,
    });

    if (!response.ok) {
      const payload = await response.text();
      throw new Error(payload || `Request failed (${response.status})`);
    }

    return response.json();
  } catch (error) {
    if (error && error.name === "AbortError") {
      throw new Error(`Request timed out after ${Math.round(timeoutMs / 1000)}s`);
    }
    throw error;
  } finally {
    if (timeoutHandle) {
      clearTimeout(timeoutHandle);
    }
  }
}

function rowsToTsv(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return "";
  }

  return rows
    .map((row) => {
      const cells = Array.isArray(row) ? row : [row];
      return cells.map((cell) => String(cell ?? "")).join("\t");
    })
    .join("\n");
}

function trimRowsForContext(rows, maxRows = IMAGE_CONTEXT_MAX_ROWS, maxCols = IMAGE_CONTEXT_MAX_COLS) {
  const rectangular = toRectangularData(rows);
  return rectangular
    .slice(0, maxRows)
    .map((row) => row.slice(0, maxCols).map((cell) => String(cell ?? "").trim()));
}

function truncateText(value, maxChars = IMAGE_CONTEXT_TEXT_MAX_CHARS) {
  const text = String(value ?? "").trim();
  if (text.length <= maxChars) {
    return text;
  }
  return `${text.slice(0, maxChars)}...`;
}

function setLatestImageContext(fileName, imageResult) {
  const rows = trimRowsForContext(imageResult.rows || []);
  const payload = imageResult.payload || {};
  const text = truncateText(payload.text || "");

  state.latestImageContext = {
    file_name: String(fileName || "image"),
    extracted_text: text,
    rows,
    row_count: Array.isArray(imageResult.rows) ? imageResult.rows.length : 0,
    column_count: Array.isArray(imageResult.rows) && imageResult.rows.length ? imageResult.rows[0].length : 0,
    ocr_engine: payload.ocr_engine || "ocr",
    layout_source: payload.layout_source || "text",
  };
}

function normalizeOcrCellText(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function median(numbers, fallback = 0) {
  const values = numbers
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value))
    .sort((a, b) => a - b);
  if (!values.length) {
    return fallback;
  }
  const middle = Math.floor(values.length / 2);
  return values.length % 2 === 0 ? (values[middle - 1] + values[middle]) / 2 : values[middle];
}

function groupBrowserOcrWordsIntoRows(words) {
  if (!Array.isArray(words) || !words.length) {
    return [];
  }

  const sorted = words
    .map((word) => {
      const text = normalizeOcrCellText(word && word.text);
      const bbox = word && typeof word === "object" ? word.bbox || {} : {};
      const left = Number(bbox.x0);
      const top = Number(bbox.y0);
      const right = Number(bbox.x1);
      const bottom = Number(bbox.y1);

      if (!text || !Number.isFinite(left) || !Number.isFinite(top) || !Number.isFinite(right) || !Number.isFinite(bottom)) {
        return null;
      }

      return {
        text,
        left,
        top,
        right,
        bottom,
        width: Math.max(1, right - left),
        height: Math.max(1, bottom - top),
        centerY: top + ((bottom - top) / 2),
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.top - b.top) || (a.left - b.left));

  if (!sorted.length) {
    return [];
  }

  const tolerance = Math.max(10, median(sorted.map((word) => word.height), 16) * 0.7);
  const rows = [];

  sorted.forEach((word) => {
    const lastRow = rows[rows.length - 1];
    if (!lastRow || Math.abs(lastRow.centerY - word.centerY) > tolerance) {
      rows.push({
        centerY: word.centerY,
        words: [word],
      });
      return;
    }

    lastRow.words.push(word);
    lastRow.centerY = (lastRow.centerY + word.centerY) / 2;
  });

  return rows.map((row) => row.words.sort((a, b) => a.left - b.left));
}

function buildRowsFromBrowserOcr(words) {
  const groupedRows = groupBrowserOcrWordsIntoRows(words);
  if (!groupedRows.length) {
    return [];
  }

  const gapThreshold = Math.max(18, median(
    groupedRows.flatMap((row) => row.map((word) => word.width)),
    28,
  ) * 1.35);

  const rows = groupedRows.map((rowWords) => {
    const cells = [];
    let currentCell = null;

    rowWords.forEach((word) => {
      if (!currentCell) {
        currentCell = { text: word.text, right: word.right };
        return;
      }

      const gap = word.left - currentCell.right;
      if (gap > gapThreshold) {
        cells.push(currentCell.text);
        currentCell = { text: word.text, right: word.right };
        return;
      }

      currentCell.text = `${currentCell.text} ${word.text}`.trim();
      currentCell.right = word.right;
    });

    if (currentCell) {
      cells.push(currentCell.text);
    }

    return cells;
  });

  return normalizeRows(rows);
}

function rowsFromOcrPayload(payload) {
  const candidateRows = Array.isArray(payload && payload.rows) ? payload.rows : [];
  if (candidateRows.length) {
    return normalizeRows(
      candidateRows.map((row) => (
        Array.isArray(row)
          ? row.map((cell) => String(cell ?? ""))
          : [String(row ?? "")]
      ))
    );
  }

  return parseTabularText(String((payload && payload.text) || ""));
}

async function requestServerOcr(imageBase64) {
  return fetchJsonWithTimeout(
    "/api/ocr-text",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        image_base64: imageBase64,
        use_ai_layout: false,
      }),
    },
    OCR_SERVER_TIMEOUT_MS,
  );
}

async function requestBrowserOcr(imageBase64) {
  if (!window.Tesseract || typeof window.Tesseract.recognize !== "function") {
    throw new Error("Browser OCR fallback is unavailable because Tesseract.js did not load.");
  }

  let timeoutHandle = null;
  try {
    const result = await Promise.race([
      window.Tesseract.recognize(imageBase64, "eng", {
        logger(message) {
          if (!message || !message.status) {
            return;
          }
          const progress = Number.isFinite(message.progress)
            ? ` ${Math.round(message.progress * 100)}%`
            : "";
          updateStatus(ui.ocrStatus, `Browser OCR: ${message.status}${progress}`);
        },
      }),
      new Promise((_, reject) => {
        timeoutHandle = setTimeout(() => {
          reject(new Error(`Browser OCR timed out after ${Math.round(OCR_BROWSER_TIMEOUT_MS / 1000)}s`));
        }, OCR_BROWSER_TIMEOUT_MS);
      }),
    ]);

    const text = String((result && result.data && result.data.text) || "").trim();
    const words = Array.isArray(result && result.data && result.data.words) ? result.data.words : [];
    const geometryRows = buildRowsFromBrowserOcr(words);
    const textRows = parseTabularText(text);
    const rows = geometryRows.length ? geometryRows : textRows;

    return {
      text,
      rows,
      line_count: text ? text.split(/\r?\n/).filter((line) => line.trim()).length : 0,
      layout_source: geometryRows.length ? "browser-geometry" : "browser-text",
      ocr_engine: "tesseract.js",
      word_count: words.length,
    };
  } finally {
    if (timeoutHandle) {
      clearTimeout(timeoutHandle);
    }
  }
}

async function requestStructuredOcr(imageBase64) {
  try {
    updateStatus(ui.ocrStatus, "Trying fast server OCR...");
    const serverPayload = await requestServerOcr(imageBase64);
    if (rowsFromOcrPayload(serverPayload).length) {
      return serverPayload;
    }

    updateStatus(ui.ocrStatus, "Server OCR returned weak output. Switching to browser OCR...");
    const fallbackPayload = await requestBrowserOcr(imageBase64);
    fallbackPayload.fallback_reason = "Server OCR returned no usable table rows.";
    return fallbackPayload;
  } catch (error) {
    updateStatus(ui.ocrStatus, `Server OCR was slow or failed. Switching to browser OCR...`);
    const fallbackPayload = await requestBrowserOcr(imageBase64);
    fallbackPayload.fallback_reason = error.message;
    return fallbackPayload;
  }
}

async function extractImageContextFromFile(file) {
  const resolvedFile = file || await pickImageFile();
  updateStatus(ui.ocrStatus, "Preparing image for OCR...");
  const originalImageBase64 = await fileToDataUrl(resolvedFile);
  const optimizedImage = await optimizeImageForOcr(originalImageBase64);
  const imageBase64 = optimizedImage.imageBase64;

  if (optimizedImage.scaled) {
    updateStatus(
      ui.ocrStatus,
      `Optimized image for OCR (${optimizedImage.originalWidth}x${optimizedImage.originalHeight} -> ${optimizedImage.width}x${optimizedImage.height}).`
    );
  }

  updateStatus(ui.ocrStatus, "Analyzing image layout...");
  const payload = await requestStructuredOcr(imageBase64);
  const rows = rowsFromOcrPayload(payload);

  setLatestImageContext(resolvedFile.name, { rows, payload });

  return {
    fileName: resolvedFile.name,
    imageBase64,
    originalImageBase64,
    payload,
    rows,
  };
}

async function pasteOcrImageToServerWorkbook(imageBase64) {
  const config = getWorkbookConfig();
  if (!config.filepath || !config.sheet_name) {
    throw new Error("Workbook file and sheet name are required for server OCR paste");
  }

  return fetchJsonWithTimeout(
    "/api/paste-ocr",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        filepath: config.filepath,
        sheet_name: config.sheet_name,
        start_cell: config.start_cell,
        image_base64: imageBase64,
        use_ai_layout: false,
      }),
    },
    OCR_SERVER_TIMEOUT_MS,
  );
}

async function runOcrFlow(file = null) {
  const imageResult = await extractImageContextFromFile(file);
  const { fileName, imageBase64, payload, rows } = imageResult;

  if (!rows.length) {
    throw new Error("OCR finished, but no table-like data could be detected from the image.");
  }

  state.parsedRows = rows;
  ui.pasteInput.value = rows.length ? rowsToTsv(rows) : String(payload.text || "");
  renderPreview(rows);

  const source = payload.layout_source ? String(payload.layout_source) : "text";
  const engine = payload.ocr_engine ? String(payload.ocr_engine) : "ocr";
  const cols = rows.length ? rows[0].length : 0;
  const lines = Number.isFinite(payload.line_count) ? Number(payload.line_count) : 0;
  const fallbackNote = payload.fallback_reason
    ? ` Server OCR failed, so browser OCR was used instead.`
    : "";
  updateStatus(ui.ocrStatus, `OCR complete (${engine}, ${source}). ${rows.length} rows x ${cols} cols from ${lines} lines.${fallbackNote}`);

  return {
    fileName,
    rows,
    imageBase64,
    payload,
  };
}

function bindEvents() {
  document.querySelectorAll("[data-prompt]").forEach((button) => {
    button.addEventListener("click", () => {
      ui.chatInput.value = button.getAttribute("data-prompt") || "";
      ui.chatInput.focus();
      void sendChat();
    });
  });

  ui.sendBtn.addEventListener("click", () => {
    void sendChat();
  });

  ui.chatImageBtn.addEventListener("click", async () => {
    await runBusyTask("Reading image content for AI...", async () => {
      try {
        const file = await pickImageFile();
        const imageResult = await extractImageContextFromFile(file);
        const hasRows = Array.isArray(imageResult.rows) && imageResult.rows.length > 0;
        const extractedText = String((imageResult.payload && imageResult.payload.text) || "").trim();

        if (!hasRows && !extractedText) {
          throw new Error("The image was read, but no usable text or table data was detected.");
        }

        if (hasRows) {
          state.parsedRows = imageResult.rows;
          ui.pasteInput.value = rowsToTsv(imageResult.rows);
          renderPreview(imageResult.rows);
        }

        appendMessage(
          "system",
          hasRows
            ? `Image attached for AI context: ${imageResult.fileName} (${imageResult.rows.length} row(s) detected).`
            : `Image attached for AI context: ${imageResult.fileName} (text extracted).`
        );
        updateStatus(ui.actionStatus, `Image context ready from ${imageResult.fileName}. Send your prompt.`);
      } catch (error) {
        updateStatus(ui.ocrStatus, error.message, true);
        updateStatus(ui.actionStatus, error.message, true);
      }
    });
  });

  ui.chatInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void sendChat();
    }
  });

  ui.clearChatBtn.addEventListener("click", () => {
    if (state.isBusy) {
      return;
    }
    state.messages = [];
    state.parsedRows = [];
    state.lastSheetSnapshot = null;
    state.latestImageContext = null;
    ui.messages.innerHTML = "";
    appendMessage("system", "Conversation cleared.");
  });

  ui.previewBtn.addEventListener("click", () => {
    if (state.isBusy) {
      return;
    }
    const rows = parseTabularText(ui.pasteInput.value);
    state.parsedRows = rows;
    renderPreview(rows);
  });

  ui.pasteExcelBtn.addEventListener("click", async () => {
    await runBusyTask("Pasting parsed rows into Excel...", async () => {
      try {
        const rows = state.parsedRows.length ? state.parsedRows : parseTabularText(ui.pasteInput.value);
        state.parsedRows = rows;
        await pasteToActiveSelection(rows);
        updateStatus(ui.actionStatus, "Pasted to active Excel selection.");
      } catch (error) {
        updateStatus(ui.actionStatus, error.message, true);
      }
    });
  });

  ui.pasteServerBtn.addEventListener("click", async () => {
    await runBusyTask("Pasting parsed rows into the server workbook...", async () => {
      try {
        const rows = state.parsedRows.length ? state.parsedRows : parseTabularText(ui.pasteInput.value);
        state.parsedRows = rows;
        const payload = await pasteToServerWorkbook(rows);
        updateStatus(ui.actionStatus, payload.message || "Data pasted to server workbook.");
      } catch (error) {
        updateStatus(ui.actionStatus, error.message, true);
      }
    });
  });

  ui.ocrBtn.addEventListener("click", async () => {
    await runBusyTask("Running OCR on the selected image...", async () => {
      try {
        await runOcrFlow();
      } catch (error) {
        updateStatus(ui.ocrStatus, error.message, true);
      }
    });
  });

  ui.ocrToExcelBtn.addEventListener("click", async () => {
    await runBusyTask("Extracting the image and pasting it into Excel...", async () => {
      try {
        const ocrResult = await runOcrFlow();
        await pasteToActiveSelection(ocrResult.rows);
        updateStatus(ui.actionStatus, "OCR table pasted to active Excel selection.");
      } catch (error) {
        updateStatus(ui.actionStatus, error.message, true);
      }
    });
  });

  ui.ocrToServerBtn.addEventListener("click", async () => {
    await runBusyTask("Extracting the image and writing it to the server workbook...", async () => {
      try {
        const ocrResult = await runOcrFlow();
        const payload = await pasteToServerWorkbook(ocrResult.rows);
        updateStatus(ui.actionStatus, payload.message || "OCR table pasted to server workbook.");
      } catch (error) {
        updateStatus(ui.actionStatus, error.message, true);
      }
    });
  });
}

function initOfficeStatus() {
  if (!window.Office) {
    state.isExcelHost = false;
    updateStatus(ui.officeStatus, "Office.js unavailable. Web mode only.", true);
    appendMessage("system", "Running in browser mode. Direct Excel paste needs Excel add-in host.");
    return;
  }

  Office.onReady((info) => {
    const inExcel = info && info.host && info.host.toString().toLowerCase().includes("excel");
    state.isExcelHost = Boolean(inExcel);
    if (inExcel) {
      updateStatus(ui.officeStatus, "Excel connected. Direct paste is enabled.");
      appendMessage("system", "Excel host detected. Chat operations now apply directly to the active workbook.");
    } else {
      updateStatus(ui.officeStatus, "Office loaded but Excel host not detected.", true);
      appendMessage("system", "Office loaded. Open in Excel for direct sheet writes.");
    }
  });
}

function bootstrap() {
  bindEvents();
  initOfficeStatus();
  appendMessage(
    "assistant",
    "I am ready. Ask me to build formulas, summarize sheets, generate table layouts, or paste typed and OCR data into Excel."
  );
}

bootstrap();
