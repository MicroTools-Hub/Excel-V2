# Smart Excel Copilot Setup

This guide covers local run, Render deployment, and Excel sideloading for the DeepSeek/Ollama-powered add-in.

## 1. Local Run

```bash
pip install -r requirements.txt
pip install -e .
```

Set environment variables:

```bash
# Linux/macOS
export OLLAMA_BASE_URL="https://ollama.com"
export OLLAMA_MODEL="deepseek-v3.2:cloud"
export EXCEL_FILES_PATH="./excel_files"

# Windows PowerShell
$env:OLLAMA_BASE_URL="https://ollama.com"
$env:OLLAMA_MODEL="deepseek-v3.2:cloud"
$env:EXCEL_FILES_PATH="./excel_files"
```

Start the app:

```bash
python -m excel_mcp.webapp
```

Open:

- `http://localhost:10000/ui/taskpane.html`

## 2. Deploy to Render

This repo already includes:

- `render.yaml`
- `Dockerfile`
- `requirements.txt`

Steps:

1. Push the repo to GitHub.
2. In Render, create a Blueprint from this repository.
3. Set secret env vars:
   - `OLLAMA_BASE_URL`
   - `OLLAMA_API_KEY` (only if your provider requires it)
4. Deploy and wait for health check (`/health`) to pass.

## 3. Configure the Office Add-in Manifest

The manifest is at `manifest.xml`.

Replace every `https://your-render-service.onrender.com` with your real service URL.

## 4. Sideload in Excel Desktop (Windows)

1. Save the updated `manifest.xml`.
2. Open **Excel**.
3. Go to **Insert** -> **My Add-ins** -> **Manage My Add-ins**.
4. Choose **Upload My Add-in** and select `manifest.xml`.
5. Open the add-in from the Home ribbon group `Smart Copilot`.

## 5. Core Workflows

- Chat with DeepSeek to plan workbook changes.
- Paste CSV/TSV/markdown/free text and push it to Excel selection.
- Upload screenshot/image, OCR text, then paste to sheet.
- Use server workbook mode to write into files under `EXCEL_FILES_PATH`.

## 6. API Summary

- `GET /health`
- `POST /api/chat`
- `POST /api/tool-call` (supports existing `server.py` tool names)
- `POST /api/parse-text`
- `POST /api/paste-text`
- `POST /api/ocr-text`
- `POST /api/paste-ocr`
