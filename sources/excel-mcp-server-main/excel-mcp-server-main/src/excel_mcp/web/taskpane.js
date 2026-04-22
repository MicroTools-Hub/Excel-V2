const state = {
  messages: [],
  parsedRows: [],
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
    auto_execute: true,
  };

  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(errorText || "Chat request failed");
    }

    const data = await response.json();
    const suffix = data.operation_results && data.operation_results.length
      ? `\n\nExecuted ${data.operation_results.length} operation(s).`
      : "";
    appendMessage("assistant", `${data.reply || "No response"}${suffix}`);
    state.messages.push({ role: "assistant", content: data.reply || "" });

    updateStatus(ui.actionStatus, `Model: ${data.model}`);
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
    updateStatus(ui.officeStatus, "Office.js unavailable. Web mode only.", true);
    appendMessage("system", "Running in browser mode. Direct Excel paste needs Excel add-in host.");
    return;
  }

  Office.onReady((info) => {
    const inExcel = info && info.host && info.host.toString().toLowerCase().includes("excel");
    if (inExcel) {
      updateStatus(ui.officeStatus, "Excel connected. Direct paste is enabled.");
      appendMessage("system", "Excel host detected. Select a cell and use Paste to Active Excel Selection.");
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
