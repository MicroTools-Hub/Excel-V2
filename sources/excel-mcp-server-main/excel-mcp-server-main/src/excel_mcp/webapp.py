import base64
import io
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional
import inspect

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from pydantic import BaseModel, Field

from excel_mcp.calculations import apply_formula as apply_formula_impl
from excel_mcp.data import read_excel_range, write_data
from excel_mcp.formatting import format_range as format_range_impl
from excel_mcp.ollama_client import OllamaClient
from excel_mcp.smart_paste import parse_tabular_text
from excel_mcp.workbook import create_sheet, create_workbook, get_workbook_info
from excel_mcp import server as mcp_server_module

logger = logging.getLogger("excel-mcp.webapp")


class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    messages: List[ChatMessage]
    filepath: Optional[str] = None
    sheet_name: Optional[str] = None
    start_cell: str = "A1"
    auto_execute: bool = True


class ToolCallRequest(BaseModel):
    tool_name: str
    args: Dict[str, Any] = Field(default_factory=dict)


class PasteTextRequest(BaseModel):
    filepath: str
    sheet_name: str
    start_cell: str = "A1"
    text: str


class ParseTextRequest(BaseModel):
    text: str


class OcrRequest(BaseModel):
    image_base64: str


class OcrPasteRequest(BaseModel):
    filepath: str
    sheet_name: str
    start_cell: str = "A1"
    image_base64: str


EXCEL_FILES_BASE = os.path.realpath(os.environ.get("EXCEL_FILES_PATH", "./excel_files"))
os.makedirs(EXCEL_FILES_BASE, exist_ok=True)

APP_ROOT = Path(__file__).resolve().parent
WEB_ROOT = APP_ROOT / "web"

app = FastAPI(
    title="Excel Smart Add-in Backend",
    version="1.0.0",
    description="DeepSeek + Ollama powered backend for Excel add-in workflows.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

if WEB_ROOT.exists():
    app.mount("/ui", StaticFiles(directory=str(WEB_ROOT), html=True), name="ui")


def _resolved_path_is_within(base: str, candidate: str) -> bool:
    base = os.path.realpath(base)
    candidate = os.path.realpath(candidate)
    if candidate == base:
        return True
    try:
        return os.path.commonpath([base, candidate]) == base
    except ValueError:
        return False


def resolve_excel_path(filename: str) -> str:
    if not filename or "\x00" in filename:
        raise ValueError("Invalid filename")
    if os.path.isabs(filename):
        raise ValueError("filepath must be relative to EXCEL_FILES_PATH")

    candidate = os.path.realpath(os.path.join(EXCEL_FILES_BASE, filename))
    if not _resolved_path_is_within(EXCEL_FILES_BASE, candidate):
        raise ValueError("filepath escapes EXCEL_FILES_PATH")
    return candidate


def _extract_json_payload(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None

    tag_match = re.search(r"<excel_plan>(.*?)</excel_plan>", text, flags=re.DOTALL | re.IGNORECASE)
    if tag_match:
        chunk = tag_match.group(1).strip()
        try:
            return json.loads(chunk)
        except json.JSONDecodeError:
            return None

    fenced = re.search(r"```json\s*(.*?)```", text, flags=re.DOTALL | re.IGNORECASE)
    if fenced:
        chunk = fenced.group(1).strip()
        try:
            return json.loads(chunk)
        except json.JSONDecodeError:
            return None

    stripped = text.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        try:
            return json.loads(stripped)
        except json.JSONDecodeError:
            return None

    return None


def _summarize_sheet(filepath: str, sheet_name: str) -> Dict[str, Any]:
    full_path = resolve_excel_path(filepath)
    workbook_info = get_workbook_info(full_path, include_ranges=False)
    preview = read_excel_range(full_path, sheet_name, start_cell="A1", end_cell="H20")
    return {
        "workbook": workbook_info,
        "preview": preview,
    }


def _execute_operation(name: str, args: Dict[str, Any]) -> Dict[str, Any]:
    operation = name.strip().lower()

    if operation == "create_workbook":
        filepath = resolve_excel_path(args["filepath"])
        create_workbook(filepath)
        return {"ok": True, "message": f"Workbook created: {args['filepath']}"}

    if operation == "create_worksheet":
        filepath = resolve_excel_path(args["filepath"])
        result = create_sheet(filepath, args["sheet_name"])
        return {"ok": True, "message": result.get("message", "Worksheet created")}

    if operation == "write_data":
        filepath = resolve_excel_path(args["filepath"])
        result = write_data(
            filepath=filepath,
            sheet_name=args.get("sheet_name"),
            data=args["data"],
            start_cell=args.get("start_cell", "A1"),
        )
        return {"ok": True, "message": result.get("message", "Data written")}

    if operation == "read_data":
        filepath = resolve_excel_path(args["filepath"])
        values = read_excel_range(
            filepath=filepath,
            sheet_name=args["sheet_name"],
            start_cell=args.get("start_cell", "A1"),
            end_cell=args.get("end_cell"),
        )
        return {"ok": True, "rows": values}

    if operation == "apply_formula":
        filepath = resolve_excel_path(args["filepath"])
        result = apply_formula_impl(filepath, args["sheet_name"], args["cell"], args["formula"])
        return {"ok": True, "message": result.get("message", "Formula applied")}

    if operation == "format_range":
        filepath = resolve_excel_path(args["filepath"])
        format_range_impl(
            filepath=filepath,
            sheet_name=args["sheet_name"],
            start_cell=args["start_cell"],
            end_cell=args.get("end_cell"),
            bold=args.get("bold", False),
            italic=args.get("italic", False),
            underline=args.get("underline", False),
            font_size=args.get("font_size"),
            font_color=args.get("font_color"),
            bg_color=args.get("bg_color"),
            border_style=args.get("border_style"),
            border_color=args.get("border_color"),
            number_format=args.get("number_format"),
            alignment=args.get("alignment"),
            wrap_text=args.get("wrap_text", False),
            merge_cells=args.get("merge_cells", False),
            protection=args.get("protection"),
            conditional_format=args.get("conditional_format"),
        )
        return {"ok": True, "message": "Range formatted"}

    if operation == "clear_range":
        filepath = resolve_excel_path(args["filepath"])
        if not args.get("sheet_name"):
            raise ValueError("sheet_name is required for clear_range")

        start_cell = str(args.get("start_cell", "A1"))
        end_cell = str(args.get("end_cell", "")).strip()
        range_address = str(args.get("range", "")).strip()
        if not range_address:
            range_address = f"{start_cell}:{end_cell}" if end_cell else start_cell

        wb = load_workbook(filepath)
        if args["sheet_name"] not in wb.sheetnames:
            wb.close()
            raise ValueError(f"Sheet '{args['sheet_name']}' not found")

        ws = wb[args["sheet_name"]]
        for row in ws[range_address]:
            for cell in row:
                cell.value = None

        wb.save(filepath)
        wb.close()
        return {"ok": True, "message": f"Cleared range {range_address} in {args['sheet_name']}"}

    if operation == "get_workbook_metadata":
        filepath = resolve_excel_path(args["filepath"])
        info = get_workbook_info(filepath, include_ranges=bool(args.get("include_ranges", False)))
        return {"ok": True, "metadata": info}

    # Fallback: try dispatching any existing Excel MCP tool from server.py.
    blocked = {
        "run_sse",
        "run_streamable_http",
        "run_stdio",
        "get_excel_path",
        "_resolved_path_is_within",
    }
    if operation in blocked:
        raise ValueError(f"Unsupported operation: {name}")

    candidate = getattr(mcp_server_module, operation, None)
    if callable(candidate):
        signature = inspect.signature(candidate)
        accepted_args = {
            key: value
            for key, value in args.items()
            if key in signature.parameters
        }
        previous_base = mcp_server_module.EXCEL_FILES_PATH
        mcp_server_module.EXCEL_FILES_PATH = EXCEL_FILES_BASE
        try:
            output = candidate(**accepted_args)
            return {"ok": True, "message": output}
        finally:
            mcp_server_module.EXCEL_FILES_PATH = previous_base

    raise ValueError(f"Unsupported operation: {name}")


async def _run_llm_chat(request: ChatRequest) -> Dict[str, Any]:
    llm = OllamaClient()

    system_content = (
        "You are Excel Copilot running on top of DeepSeek via Ollama. "
        "Respond with practical spreadsheet guidance. "
        "When an operation should be executed, include an XML-like plan block exactly in this format: "
        "<excel_plan>{\"assistant_reply\":\"...\",\"operations\":[{\"name\":\"write_data\",\"args\":{...}}]}</excel_plan>. "
        "Supported operation names: create_workbook, create_worksheet, write_data, clear_range, read_data, apply_formula, format_range, get_workbook_metadata. "
        "For write_data, prefer args {start_cell, data, optional sheet_name}; use range only when explicitly requested. "
        "For clearing cells, prefer clear_range with args {range, optional sheet_name}. "
        "If no operation is needed, operations must be an empty list."
    )

    model_messages: List[Dict[str, str]] = [{"role": "system", "content": system_content}]

    if request.filepath and request.sheet_name:
        try:
            context = _summarize_sheet(request.filepath, request.sheet_name)
            model_messages.append(
                {
                    "role": "system",
                    "content": "Workbook context: " + json.dumps(context, default=str),
                }
            )
        except Exception as exc:
            logger.warning("Could not build workbook context: %s", exc)

    for msg in request.messages[-24:]:
        model_messages.append({"role": msg.role, "content": msg.content})

    llm_response = await llm.chat(model_messages)
    raw_text = llm_response["content"]
    plan = _extract_json_payload(raw_text)

    reply = raw_text.strip()
    operations: List[Dict[str, Any]] = []

    if isinstance(plan, dict):
        if isinstance(plan.get("assistant_reply"), str):
            reply = plan["assistant_reply"]
        ops = plan.get("operations")
        if isinstance(ops, list):
            operations = [op for op in ops if isinstance(op, dict)]

    results: List[Dict[str, Any]] = []
    if request.auto_execute and operations:
        for op in operations[:8]:
            op_name = str(op.get("name", "")).strip()
            op_args = op.get("args") if isinstance(op.get("args"), dict) else {}
            merged_args = dict(op_args)
            if request.filepath and "filepath" not in merged_args:
                merged_args["filepath"] = request.filepath
            if request.sheet_name and "sheet_name" not in merged_args:
                merged_args["sheet_name"] = request.sheet_name
            if request.start_cell and "start_cell" not in merged_args:
                merged_args["start_cell"] = request.start_cell
            try:
                result = _execute_operation(op_name, merged_args)
                results.append({"name": op_name, "ok": True, "result": result})
            except Exception as exc:
                results.append({"name": op_name, "ok": False, "error": str(exc)})

    return {
        "reply": reply,
        "model": llm_response["model"],
        "operations": operations,
        "operation_results": results,
    }


def _ocr_image_to_text(image_base64: str) -> str:
    try:
        from PIL import Image
        import pytesseract
    except ImportError as exc:
        raise HTTPException(status_code=500, detail="OCR dependencies are not installed") from exc

    payload = image_base64.split(",", 1)[1] if "," in image_base64 else image_base64
    try:
        image_bytes = base64.b64decode(payload)
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Invalid base64 image payload") from exc

    try:
        image = Image.open(io.BytesIO(image_bytes))
        text = pytesseract.image_to_string(image)
        return text.strip()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"OCR processing failed: {exc}") from exc


@app.get("/health")
def health() -> Dict[str, Any]:
    return {
        "status": "ok",
        "model": os.environ.get("OLLAMA_MODEL", "deepseek-v3.2:cloud"),
        "ollama_base_url": os.environ.get("OLLAMA_BASE_URL", "https://ollama.com"),
        "ollama_api_key_configured": bool(os.environ.get("OLLAMA_API_KEY")),
        "excel_files_path": EXCEL_FILES_BASE,
    }


@app.get("/")
def root() -> Any:
    if WEB_ROOT.exists():
        return RedirectResponse(url="/ui/taskpane.html")
    return {"status": "ok", "message": "UI files missing"}


@app.get("/ui/taskpane.html", include_in_schema=False)
def taskpane() -> Any:
    path = WEB_ROOT / "taskpane.html"
    if not path.exists():
        raise HTTPException(status_code=404, detail="taskpane.html not found")
    return FileResponse(path)


@app.post("/api/chat")
async def chat(request: ChatRequest) -> Dict[str, Any]:
    if not request.messages:
        raise HTTPException(status_code=400, detail="messages cannot be empty")
    try:
        return await _run_llm_chat(request)
    except HTTPException:
        raise
    except Exception as exc:
        logger.error("Chat request failed: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc)) from exc


@app.post("/api/tool-call")
def tool_call(request: ToolCallRequest) -> Dict[str, Any]:
    try:
        result = _execute_operation(request.tool_name, request.args)
        return {"ok": True, "result": result}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/parse-text")
def parse_text_endpoint(request: ParseTextRequest) -> Dict[str, Any]:
    rows = parse_tabular_text(request.text)
    return {
        "rows": rows,
        "row_count": len(rows),
        "column_count": max((len(r) for r in rows), default=0),
    }


@app.post("/api/paste-text")
def paste_text(request: PasteTextRequest) -> Dict[str, Any]:
    rows = parse_tabular_text(request.text)
    if not rows:
        raise HTTPException(status_code=400, detail="No tabular content detected")

    full_path = resolve_excel_path(request.filepath)
    result = write_data(
        filepath=full_path,
        sheet_name=request.sheet_name,
        data=rows,
        start_cell=request.start_cell,
    )
    return {
        "ok": True,
        "message": result.get("message", "Data written"),
        "rows_written": len(rows),
        "columns_written": max((len(r) for r in rows), default=0),
    }


@app.post("/api/ocr-text")
def ocr_text(request: OcrRequest) -> Dict[str, Any]:
    text = _ocr_image_to_text(request.image_base64)
    return {
        "text": text,
        "line_count": len([line for line in text.splitlines() if line.strip()]),
    }


@app.post("/api/paste-ocr")
def paste_ocr(request: OcrPasteRequest) -> Dict[str, Any]:
    text = _ocr_image_to_text(request.image_base64)
    rows = parse_tabular_text(text)
    if not rows:
        raise HTTPException(status_code=400, detail="OCR succeeded but no tabular data was detected")

    full_path = resolve_excel_path(request.filepath)
    result = write_data(
        filepath=full_path,
        sheet_name=request.sheet_name,
        data=rows,
        start_cell=request.start_cell,
    )
    return {
        "ok": True,
        "message": result.get("message", "Data written"),
        "ocr_text": text,
        "rows_written": len(rows),
        "columns_written": max((len(r) for r in rows), default=0),
    }


def run_web_app() -> None:
    import uvicorn

    host = os.environ.get("WEBAPP_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "10000"))
    uvicorn.run("excel_mcp.webapp:app", host=host, port=port)


if __name__ == "__main__":
    run_web_app()
