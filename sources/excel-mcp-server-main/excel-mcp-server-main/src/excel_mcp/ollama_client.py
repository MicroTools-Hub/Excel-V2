import os
from typing import Any, Dict, List, Optional

import httpx


def _sanitize_env_value(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    cleaned = str(value).strip()
    if (cleaned.startswith('"') and cleaned.endswith('"')) or (
        cleaned.startswith("'") and cleaned.endswith("'")
    ):
        cleaned = cleaned[1:-1].strip()
    return cleaned or None


class OllamaClient:
    """Minimal async client for Ollama-compatible chat APIs."""

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        api_key: Optional[str] = None,
        timeout_seconds: Optional[float] = None,
    ) -> None:
        resolved_base = _sanitize_env_value(base_url) or _sanitize_env_value(os.environ.get("OLLAMA_BASE_URL")) or "https://ollama.com"
        self.base_url = resolved_base.rstrip("/")
        self.model = _sanitize_env_value(model) or _sanitize_env_value(os.environ.get("OLLAMA_MODEL")) or "deepseek-v3.2:cloud"
        self.api_key = _sanitize_env_value(api_key) or _sanitize_env_value(os.environ.get("OLLAMA_API_KEY"))
        timeout = timeout_seconds or float(os.environ.get("OLLAMA_TIMEOUT_SECONDS", "120"))
        self.timeout = httpx.Timeout(timeout)

    def _candidate_chat_endpoints(self) -> List[tuple[str, bool]]:
        """Return candidate chat endpoints and format flags.

        Tuple item format: (endpoint_url, openai_compatible_payload)
        """
        base = self.base_url.rstrip("/")
        candidates: List[tuple[str, bool]] = []

        if base.endswith("/v1"):
            candidates.append((f"{base}/chat/completions", True))
        elif base.endswith("/api"):
            candidates.append((f"{base}/chat", False))
            candidates.append((f"{base[:-4]}/v1/chat/completions", True))
        else:
            candidates.append((f"{base}/api/chat", False))
            candidates.append((f"{base}/v1/chat/completions", True))

        # Keep first occurrence only while preserving order.
        deduped: List[tuple[str, bool]] = []
        seen: set[str] = set()
        for endpoint, openai_compatible in candidates:
            if endpoint in seen:
                continue
            seen.add(endpoint)
            deduped.append((endpoint, openai_compatible))

        return deduped

    @staticmethod
    def _error_text(response: httpx.Response) -> str:
        try:
            payload = response.json()
            if isinstance(payload, dict):
                for key in ("error", "message", "detail"):
                    value = payload.get(key)
                    if isinstance(value, str) and value.strip():
                        return value.strip()
                return str(payload)
            return str(payload)
        except Exception:
            text = response.text.strip()
            return text[:600] if text else "Unknown upstream error"

    def _requires_cloud_auth(self) -> bool:
        return "ollama.com" in self.base_url.lower()

    async def chat(
        self,
        messages: List[Dict[str, str]],
        *,
        temperature: float = 0.2,
        model: Optional[str] = None,
    ) -> Dict[str, Any]:
        model_name = model or self.model

        if self._requires_cloud_auth() and not self.api_key:
            raise RuntimeError("OLLAMA_API_KEY is required when using Ollama cloud endpoints")

        headers: Dict[str, str] = {}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        attempts: List[str] = []
        async with httpx.AsyncClient(timeout=self.timeout, follow_redirects=True) as client:
            for endpoint, openai_compatible in self._candidate_chat_endpoints():
                if openai_compatible:
                    payload: Dict[str, Any] = {
                        "model": model_name,
                        "messages": messages,
                        "temperature": temperature,
                        "stream": False,
                    }
                else:
                    payload = {
                        "model": model_name,
                        "messages": messages,
                        "stream": False,
                        "options": {
                            "temperature": temperature,
                        },
                    }

                try:
                    response = await client.post(endpoint, json=payload, headers=headers)
                except httpx.RequestError as exc:
                    attempts.append(f"{endpoint} -> request error: {repr(exc)}")
                    continue

                if response.status_code >= 400:
                    details = self._error_text(response)
                    # For endpoint-shape mismatches, try the next candidate.
                    if response.status_code in {404, 405, 422}:
                        attempts.append(f"{endpoint} -> HTTP {response.status_code}: {details}")
                        continue
                    raise RuntimeError(f"Ollama API error {response.status_code} at {endpoint}: {details}")

                try:
                    data = response.json()
                except ValueError as exc:
                    attempts.append(f"{endpoint} -> invalid JSON: {repr(exc)}")
                    continue

                if openai_compatible:
                    choices = data.get("choices") or []
                    first_choice = choices[0] if choices else {}
                    message = first_choice.get("message") if isinstance(first_choice, dict) else {}
                    content = message.get("content", "") if isinstance(message, dict) else ""
                else:
                    message = data.get("message") or {}
                    content = message.get("content", "")

                return {
                    "model": data.get("model") or model_name,
                    "content": content,
                    "raw": data,
                }

        detail_lines = " | ".join(attempts) if attempts else "No attempts were made"
        raise RuntimeError(f"Could not reach a working Ollama chat endpoint from base URL {self.base_url}. Attempts: {detail_lines}")
