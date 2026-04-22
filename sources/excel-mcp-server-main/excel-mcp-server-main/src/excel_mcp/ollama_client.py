import os
from typing import Any, Dict, List, Optional

import httpx


class OllamaClient:
    """Minimal async client for Ollama-compatible chat APIs."""

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        api_key: Optional[str] = None,
        timeout_seconds: Optional[float] = None,
    ) -> None:
        self.base_url = (base_url or os.environ.get("OLLAMA_BASE_URL", "https://ollama.com")).rstrip("/")
        self.model = model or os.environ.get("OLLAMA_MODEL", "deepseek-v3.2:cloud")
        self.api_key = api_key or os.environ.get("OLLAMA_API_KEY")
        timeout = timeout_seconds or float(os.environ.get("OLLAMA_TIMEOUT_SECONDS", "120"))
        self.timeout = httpx.Timeout(timeout)

    def _chat_endpoint(self) -> tuple[str, bool]:
        """Return chat endpoint URL and whether it is OpenAI-compatible format."""
        base = self.base_url.rstrip("/")
        if base.endswith("/v1"):
            return f"{base}/chat/completions", True
        if base.endswith("/api"):
            return f"{base}/chat", False
        return f"{base}/api/chat", False

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
        endpoint, openai_compatible = self._chat_endpoint()
        model_name = model or self.model

        if self._requires_cloud_auth() and not self.api_key:
            raise RuntimeError("OLLAMA_API_KEY is required when using Ollama cloud endpoints")

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

        headers: Dict[str, str] = {}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        try:
            async with httpx.AsyncClient(timeout=self.timeout) as client:
                response = await client.post(endpoint, json=payload, headers=headers)
        except httpx.RequestError as exc:
            raise RuntimeError(f"Could not reach Ollama endpoint {endpoint}: {exc}") from exc

        if response.status_code >= 400:
            details = self._error_text(response)
            raise RuntimeError(f"Ollama API error {response.status_code} at {endpoint}: {details}")

        try:
            data = response.json()
        except ValueError as exc:
            raise RuntimeError("Ollama API returned a non-JSON response") from exc

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
