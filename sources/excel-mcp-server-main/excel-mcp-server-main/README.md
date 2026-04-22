<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-mcp-server streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Smart Excel Copilot Add-in (DeepSeek v3.2 + Ollama)

This repository now also includes a **hostable web backend + Excel taskpane UI**:

- **ChatGPT-style interface** for spreadsheet instructions and analysis
- **DeepSeek v3.2 through Ollama-compatible API**
- **Smart operation plans** that can auto-execute workbook actions
- **Typed/pasted text to sheet** (CSV/TSV/markdown/free-form parsing)
- **OCR image to sheet** (via Tesseract)
- **Office.js direct paste** into currently selected range when opened inside Excel

### Start the Smart Web App

```bash
# Install dependencies once
pip install -r requirements.txt
pip install -e .

# Start web API + add-in UI server
python -m excel_mcp.webapp
```

Or through the Typer CLI:

```bash
python -m excel_mcp webapp
```

Default URL:

- `http://localhost:10000/ui/taskpane.html`

### Smart API Endpoints

- `GET /health`
- `POST /api/chat`
- `POST /api/tool-call` (can dispatch to existing `server.py` tool names)
- `POST /api/parse-text`
- `POST /api/paste-text`
- `POST /api/ocr-text`
- `POST /api/paste-ocr`

### Environment Variables for Smart Copilot

- `OLLAMA_BASE_URL` (for example `http://localhost:11434`)
- `OLLAMA_API_KEY` (optional, if your Ollama-compatible host requires bearer auth)
- `OLLAMA_MODEL` (default `deepseek-v3.2:cloud`)
- `OLLAMA_TIMEOUT_SECONDS` (default `120`)
- `EXCEL_FILES_PATH` (default `./excel_files`)
- `PORT` (default `10000`, used by Render)

Use `.env.example` as the starter template.

## Office Add-in Manifest

- Manifest file: `manifest.xml`
- Taskpane UI: `src/excel_mcp/web/taskpane.html`

Before sideloading the add-in, replace all `https://your-render-service.onrender.com` values in `manifest.xml` with your actual Render service URL.

## Deploy to Render (No Manual Restarts)

This project includes:

- `render.yaml` (blueprint configuration)
- `Dockerfile` (Python + Tesseract image)
- `requirements.txt`

Deployment flow:

1. Push this repo to GitHub.
2. In Render, create a new Blueprint and select your repo.
3. Set secret env vars (`OLLAMA_BASE_URL`, `OLLAMA_API_KEY` if required).
4. Deploy. Render will keep the service running and auto-restart on deploys/failures.

After deployment, open:

- `https://<your-service>.onrender.com/ui/taskpane.html`

Detailed local + cloud setup: `SMART_ADDIN_SETUP.md`.


## Environment Variables & File Path Handling

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.
- With these transports, tool `filepath` values must be **relative** to that directory (e.g. `reports/q1.xlsx`); absolute paths and directory traversal are rejected.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8017` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
