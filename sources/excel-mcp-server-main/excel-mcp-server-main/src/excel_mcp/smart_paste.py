import csv
import io
import re
from typing import Any, List


_MARKDOWN_RULE_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$")
_INT_RE = re.compile(r"^-?\d+$")
_FLOAT_RE = re.compile(r"^-?\d+\.\d+$")


def _coerce_value(value: str) -> Any:
    token = value.strip()
    if token == "":
        return ""

    lower = token.lower()
    if lower == "true":
        return True
    if lower == "false":
        return False

    if _INT_RE.match(token):
        try:
            return int(token)
        except ValueError:
            return token

    if _FLOAT_RE.match(token):
        try:
            return float(token)
        except ValueError:
            return token

    return token


def _normalize_rows(rows: List[List[str]]) -> List[List[Any]]:
    if not rows:
        return []

    width = max(len(row) for row in rows)
    normalized: List[List[Any]] = []
    for row in rows:
        padded = row + [""] * (width - len(row))
        normalized.append([_coerce_value(cell) for cell in padded])
    return normalized


def parse_tabular_text(raw_text: str) -> List[List[Any]]:
    """Parse plain text into a rectangular 2D array suitable for Excel writes."""
    text = (raw_text or "").replace("\r\n", "\n").strip()
    if not text:
        return []

    lines = [line for line in text.split("\n") if line.strip()]

    # Markdown table support.
    if lines and sum("|" in line for line in lines) >= max(1, len(lines) - 1):
        md_rows: List[List[str]] = []
        for line in lines:
            if _MARKDOWN_RULE_RE.match(line):
                continue
            cleaned = line.strip().strip("|")
            md_rows.append([cell.strip() for cell in cleaned.split("|")])
        parsed = _normalize_rows(md_rows)
        if parsed:
            return parsed

    if "\t" in text:
        rows = [[cell.strip() for cell in line.split("\t")] for line in lines]
        return _normalize_rows(rows)

    sample = "\n".join(lines[:8])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;|")
        reader = csv.reader(io.StringIO(text), dialect)
        rows = [[cell.strip() for cell in row] for row in reader if row]
        parsed = _normalize_rows(rows)
        if parsed and max(len(row) for row in parsed) > 1:
            return parsed
    except csv.Error:
        pass

    rows = [re.split(r"\s{2,}", line.strip()) for line in lines]
    parsed = _normalize_rows(rows)
    if parsed and max(len(row) for row in parsed) > 1:
        return parsed

    return [[_coerce_value(line.strip())] for line in lines if line.strip()]
