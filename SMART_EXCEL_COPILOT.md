# Smart Excel Copilot

This project packages the extracted Excel MCP code into a Render-ready Office.js Excel add-in.

## What Runs Where

Render hosts the web task pane and AI API. Excel workbook edits run inside Excel through Office.js, because Render cannot run desktop Excel or Windows COM automation.

The backend also keeps the extracted openpyxl-based MCP workbook mode for server-side `.xlsx` files under `EXCEL_FILES_PATH`.

## Local Run

```powershell
cd "D:\Excel Add-in V2\sources\excel-mcp-server-main\excel-mcp-server-main"
$env:OLLAMA_BASE_URL="https://ollama.com"
$env:OLLAMA_API_KEY="your_ollama_or_compatible_key"
$env:OLLAMA_MODEL="deepseek-v3.2:cloud"
.\.venv\Scripts\python.exe -m excel_mcp.webapp
```

Open `http://localhost:10000/ui/taskpane.html` to test browser mode.

## Render Deploy

The root `render.yaml` and `Dockerfile` are ready for a Render Blueprint.

Required secret:

```text
OLLAMA_API_KEY
```

After Render deploys, update the add-in manifest URLs to your Render service URL if it differs from `https://excel-v2.onrender.com`.

## Excel Sideload

Use the manifest at:

```text
D:\Excel Add-in V2\sources\excel-mcp-server-main\excel-mcp-server-main\manifest.xml
```

In Excel Desktop on Windows, open `Insert -> My Add-ins -> Upload My Add-in`, then select the manifest.

## Current Capabilities

- ChatGPT/Copilot-style task pane for Excel.
- Reads active sheet values and formulas before sending prompts to DeepSeek.
- Applies model-planned operations directly to the active workbook.
- Writes typed CSV, TSV, markdown tables, and plain text into the selected range.
- Extracts tables from screenshots/images with Tesseract plus optional AI layout repair.
- Supports formulas, formatting, tables, charts, summaries, sheet edits, row/column edits, range copy/delete, merge/unmerge, and server workbook mode.
