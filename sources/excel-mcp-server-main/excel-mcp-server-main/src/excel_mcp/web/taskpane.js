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
  return rangeAddress.split(":")[0].trim() || "A1";
}

async function resolveWorksheet(context, sheetName) {
  if (!sheetName) {
    return context.workbook.getActiveWorksheet();
  }

  const candidate = context.workbook.worksheets.getItemOrNullObject(String(sheetName));
  candidate.load("isNullObject");
  await context.sync();

  if (candidate.isNullObject) {
    return context.workbook.getActiveWorksheet();
  }

  return candidate;
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
        if (typeof args.font_color === "string" && args.font_color.trim()) {
          range.format.font.color = args.font_color.trim().replace(/^#/, "");
        }
        if (typeof args.bg_color === "string" && args.bg_color.trim()) {
          range.format.fill.color = args.bg_color.trim().replace(/^#/, "");
        }
        if (typeof args.wrap_text === "boolean") {
          range.format.wrapText = args.wrap_text;
        }
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
  const payload = {
    messages: state.messages,
    filepath: config.filepath || null,
    sheet_name: config.sheet_name || null,
    start_cell: config.start_cell,
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

async function runOcrFlow() {
  if (!window.Tesseract) {
    throw new Error("Tesseract.js failed to load");
  }

  const file = await pickImageFile();
  updateStatus(ui.ocrStatus, "OCR running...");

  const result = await window.Tesseract.recognize(file, "eng", {
    logger: (info) => {
      if (info && typeof info.progress === "number") {
        const percent = Math.round(info.progress * 100);
        updateStatus(ui.ocrStatus, `OCR ${percent}% (${info.status || "working"})`);
      }
    },
  });

  const text = result.data && result.data.text ? result.data.text.trim() : "";
  ui.pasteInput.value = text;
  const rows = parseTabularText(text);
  state.parsedRows = rows;
  renderPreview(rows);

  updateStatus(ui.ocrStatus, `OCR complete. Extracted ${text.length} chars.`);
  return rows;
}

function bindEvents() {
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
      const rows = await runOcrFlow();
      await pasteToActiveSelection(rows);
      updateStatus(ui.actionStatus, "OCR text pasted to active Excel selection.");
    } catch (error) {
      updateStatus(ui.actionStatus, error.message, true);
    }
  });

  ui.ocrToServerBtn.addEventListener("click", async () => {
    try {
      const rows = await runOcrFlow();
      const payload = await pasteToServerWorkbook(rows);
      updateStatus(ui.actionStatus, payload.message || "OCR text pasted to server workbook.");
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
