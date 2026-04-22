FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONPATH=/app/src

RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY sources/excel-mcp-server-main/excel-mcp-server-main/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY sources/excel-mcp-server-main/excel-mcp-server-main/src ./src
COPY sources/excel-mcp-server-main/excel-mcp-server-main/README.md ./README.md

ENV EXCEL_FILES_PATH=/data/excel_files
RUN mkdir -p /data/excel_files

CMD ["python", "-m", "excel_mcp.webapp"]
