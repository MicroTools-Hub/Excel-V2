const state = {
  messages: [],
  parsedRows: [],
  isExcelHost: false,
};

const ui = {
  messages: document.getElementById("messages"),
  chatInput: document.getElementById("chatInput"),
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

function resolveRangeAddress(args) {
  if (typeof args.range === "string" && args.range.trim()) {
    return args.range.trim();
  }

  const start = typeof args.start_cell === "string" && args.start_cell.trim() ? args.start_cell.trim() : "A1";
  const end = typeof args.end_cell === "string" && args.end_cell.trim() ? args.end_cell.trim() : "";
  return end ? `${start}:${end}` : start;
}

function getRangeStartCell(rangeAddress) {
  if (typeof rangeAddress !== "string" || !rangeAddress.trim()) {
    return "A1";
  }
  const rawStart = rangeAddress.split(":")[0].trim();
  const noSheet = rawStart.includes("!") ? rawStart.split("!").pop() : rawStart;
  const normalized = (noSheet || "").replace(/\$/g, "").trim();
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
    const activeSheet = workbook.worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();
    const usedRange = activeSheet.getUsedRangeOrNullObject(true);

    activeSheet.load("name");
    selectedRange.load("address");
    usedRange.load("isNullObject,address,rowCount,columnCount,values,formulas");
    await context.sync();

    const selectionAddress = selectedRange.address || "A1";
    const selectionStartCell = getRangeStartCell(selectionAddress);

    if (usedRange.isNullObject) {
      return {
        sheet_name: activeSheet.name,
        selection_address: selectionAddress,
        selection_start_cell: selectionStartCell,
        used_range_address: null,
        row_count: 0,
        column_count: 0,
        values: [],
        formulas: [],
      };
    }

    return {
      sheet_name: activeSheet.name,
      selection_address: selectionAddress,
      selection_start_cell: selectionStartCell,
      used_range_address: usedRange.address,
      row_count: usedRange.rowCount,
      column_count: usedRange.columnCount,
      values: usedRange.values,
      formulas: usedRange.formulas,
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
  if (!prompt) {
    return;
  }

  appendMessage("user", prompt);
  state.messages.push({ role: "user", content: prompt });
  ui.chatInput.value = "";
  updateStatus(ui.actionStatus, "Thinking...");

  const config = getWorkbookConfig();
  let sheetSnapshot = null;

  if (state.isExcelHost) {
    try {
      updateStatus(ui.actionStatus, "Reading full sheet context...");
      sheetSnapshot = await collectExcelSheetSnapshot();
    } catch (error) {
      appendMessage("system", `Warning: Could not read full sheet context: ${error.message}`);
    }
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
    auto_execute: !state.isExcelHost,
  };

  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      let errorText = "";
      try {
        const payload = await response.json();
        errorText = payload && payload.detail ? String(payload.detail) : JSON.stringify(payload);
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

  const text = ui.pasteInput.value;
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

async function requestServerOcr(imageBase64) {
  const response = await fetch("/api/ocr-text", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      image_base64: imageBase64,
      use_ai_layout: true,
    }),
  });

  if (!response.ok) {
    const payload = await response.text();
    throw new Error(payload || `OCR request failed (${response.status})`);
  }

  return response.json();
}

async function pasteOcrImageToServerWorkbook(imageBase64) {
  const config = getWorkbookConfig();
  if (!config.filepath || !config.sheet_name) {
    throw new Error("Workbook file and sheet name are required for server OCR paste");
  }

  const response = await fetch("/api/paste-ocr", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      filepath: config.filepath,
      sheet_name: config.sheet_name,
      start_cell: config.start_cell,
      image_base64: imageBase64,
      use_ai_layout: true,
    }),
  });

  if (!response.ok) {
    const payload = await response.text();
    throw new Error(payload || `OCR paste failed (${response.status})`);
  }

  return response.json();
}

async function runOcrFlow() {
  const file = await pickImageFile();
  updateStatus(ui.ocrStatus, "Preparing image for OCR...");
  const imageBase64 = await fileToDataUrl(file);

  updateStatus(ui.ocrStatus, "Analyzing image layout with AI OCR...");
  const payload = await requestServerOcr(imageBase64);

  const serverRows = Array.isArray(payload.rows) ? payload.rows : [];
  const rows = serverRows.length
    ? normalizeRows(
      serverRows.map((row) =>
        (Array.isArray(row) ? row.map((cell) => String(cell ?? "")) : [String(row ?? "")])
      )
    )
    : parseTabularText(String(payload.text || ""));

  state.parsedRows = rows;
  ui.pasteInput.value = rows.length ? rowsToTsv(rows) : String(payload.text || "");
  renderPreview(rows);

  const source = payload.layout_source ? String(payload.layout_source) : "text";
  const cols = rows.length ? rows[0].length : 0;
  const lines = Number.isFinite(payload.line_count) ? Number(payload.line_count) : 0;
  updateStatus(ui.ocrStatus, `OCR complete (${source}). ${rows.length} rows x ${cols} cols from ${lines} lines.`);

  return {
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

  ui.chatInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void sendChat();
    }
  });

  ui.clearChatBtn.addEventListener("click", () => {
    state.messages = [];
    ui.messages.innerHTML = "";
    appendMessage("system", "Conversation cleared.");
  });

  ui.previewBtn.addEventListener("click", () => {
    const rows = parseTabularText(ui.pasteInput.value);
    state.parsedRows = rows;
    renderPreview(rows);
  });

  ui.pasteExcelBtn.addEventListener("click", async () => {
    try {
      const rows = state.parsedRows.length ? state.parsedRows : parseTabularText(ui.pasteInput.value);
      state.parsedRows = rows;
      await pasteToActiveSelection(rows);
      updateStatus(ui.actionStatus, "Pasted to active Excel selection.");
    } catch (error) {
      updateStatus(ui.actionStatus, error.message, true);
    }
  });

  ui.pasteServerBtn.addEventListener("click", async () => {
    try {
      const rows = state.parsedRows.length ? state.parsedRows : parseTabularText(ui.pasteInput.value);
      state.parsedRows = rows;
      const payload = await pasteToServerWorkbook(rows);
      updateStatus(ui.actionStatus, payload.message || "Data pasted to server workbook.");
    } catch (error) {
      updateStatus(ui.actionStatus, error.message, true);
    }
  });

  ui.ocrBtn.addEventListener("click", async () => {
    try {
      await runOcrFlow();
    } catch (error) {
      updateStatus(ui.ocrStatus, error.message, true);
    }
  });

  ui.ocrToExcelBtn.addEventListener("click", async () => {
    try {
      const ocrResult = await runOcrFlow();
      await pasteToActiveSelection(ocrResult.rows);
      updateStatus(ui.actionStatus, "OCR table pasted to active Excel selection.");
    } catch (error) {
      updateStatus(ui.actionStatus, error.message, true);
    }
  });

  ui.ocrToServerBtn.addEventListener("click", async () => {
    try {
      const ocrResult = await runOcrFlow();
      const payload = await pasteOcrImageToServerWorkbook(ocrResult.imageBase64);
      updateStatus(
        ui.actionStatus,
        payload.message || `OCR pasted ${payload.rows_written || 0} rows x ${payload.columns_written || 0} cols to server workbook.`
      );
    } catch (error) {
      updateStatus(ui.actionStatus, error.message, true);
    }
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
