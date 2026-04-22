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

    async def chat(
        self,
        messages: List[Dict[str, str]],
        *,
        temperature: float = 0.2,
        model: Optional[str] = None,
    ) -> Dict[str, Any]:
        payload: Dict[str, Any] = {
            "model": model or self.model,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": temperature,
            },
        }
        headers: Dict[str, str] = {}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.post(f"{self.base_url}/api/chat", json=payload, headers=headers)
            response.raise_for_status()
            data = response.json()

        message = data.get("message") or {}
        content = message.get("content", "")
        return {
            "model": data.get("model") or payload["model"],
            "content": content,
            "raw": data,
        }
