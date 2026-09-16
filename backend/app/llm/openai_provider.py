"""OpenAI 兼容 Provider（§35 第一版必选实现）。

适用于 OpenAI / DeepSeek / Qwen / 智谱 / 企业自建等
兼容 /chat/completions 协议的接口，差异全部收敛到 .env 配置。
"""
from __future__ import annotations

import json
import logging
from collections.abc import AsyncIterator, Callable

import httpx

from .base import CompletionResult, LLMError, LLMProvider, LLMTimeoutError, Usage

logger = logging.getLogger(__name__)

_CONNECT_TIMEOUT = 10.0


def _parse_usage(data: dict) -> Usage:
    usage = data.get("usage") or {}
    return Usage(
        prompt_tokens=int(usage.get("prompt_tokens") or 0),
        completion_tokens=int(usage.get("completion_tokens") or 0),
    )


class OpenAICompatibleProvider(LLMProvider):
    def __init__(
        self,
        base_url: str,
        api_key: str,
        model: str,
        *,
        timeout: float = 120.0,
        default_temperature: float | None = 0.3,
        default_max_tokens: int | None = 2048,
    ) -> None:
        self._model = model
        self._default_temperature = default_temperature
        self._default_max_tokens = default_max_tokens
        self._client = httpx.AsyncClient(
            base_url=base_url.rstrip("/"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            timeout=httpx.Timeout(timeout, connect=_CONNECT_TIMEOUT),
        )

    @property
    def model_name(self) -> str:
        return self._model

    @property
    def supports_stream(self) -> bool:
        return True

    # ------------------------------------------------------------------
    # 非流式（Edit 生成）
    # ------------------------------------------------------------------

    async def chat(
        self,
        messages: list[dict[str, str]],
        *,
        temperature: float | None = None,
        max_tokens: int | None = None,
        timeout: float | None = None,
    ) -> CompletionResult:
        if not self._model:
            raise LLMError("LLM_MODEL 未配置，请在 backend/.env 中设置")
        payload = self._build_payload(
            messages,
            stream=False,
            temperature=self._default_temperature if temperature is None else temperature,
            max_tokens=self._default_max_tokens if max_tokens is None else max_tokens,
        )
        response = await self._request_json("/chat/completions", payload, timeout)
        data = response
        try:
            content = data["choices"][0]["message"]["content"] or ""
        except (KeyError, IndexError, TypeError) as exc:
            logger.warning("模型返回结构异常：%s", str(data)[:300])
            raise LLMError("模型返回结构异常（缺少 choices/message）") from exc
        return CompletionResult(content=content, usage=_parse_usage(data))

    # ------------------------------------------------------------------
    # 流式（Chat SSE）
    # ------------------------------------------------------------------

    async def stream_chat(
        self,
        messages: list[dict[str, str]],
        *,
        temperature: float | None = None,
        max_tokens: int | None = None,
        timeout: float | None = None,
        on_usage: Callable[[Usage], None] | None = None,
    ) -> AsyncIterator[str]:
        if not self._model:
            raise LLMError("LLM_MODEL 未配置，请在 backend/.env 中设置")
        payload = self._build_payload(
            messages,
            stream=True,
            temperature=self._default_temperature if temperature is None else temperature,
            max_tokens=self._default_max_tokens if max_tokens is None else max_tokens,
        )
        try:
            async with self._client.stream(
                "POST",
                "/chat/completions",
                json=payload,
                timeout=httpx.Timeout(timeout or self._client.timeout.read, connect=_CONNECT_TIMEOUT),
            ) as response:
                if response.status_code != 200:
                    body = (await response.aread()).decode("utf-8", errors="replace")
                    raise LLMError(self._upstream_error(response.status_code, body))
                async for line in response.aiter_lines():
                    token = self._parse_stream_line(line, on_usage)
                    if token:
                        yield token
        except httpx.TimeoutException as exc:
            raise LLMTimeoutError("模型请求超时") from exc
        except httpx.RequestError as exc:
            raise LLMError(f"无法连接模型服务：{exc}") from exc

    # ------------------------------------------------------------------
    # 内部
    # ------------------------------------------------------------------

    def _build_payload(
        self,
        messages: list[dict[str, str]],
        *,
        stream: bool,
        temperature: float | None,
        max_tokens: int | None,
    ) -> dict:
        payload: dict = {"model": self._model, "messages": messages, "stream": stream}
        if temperature is not None:
            payload["temperature"] = temperature
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens
        return payload

    async def _request_json(
        self, path: str, payload: dict, timeout: float | None
    ) -> dict:
        try:
            response = await self._client.post(
                path,
                json=payload,
                timeout=httpx.Timeout(timeout or self._client.timeout.read, connect=_CONNECT_TIMEOUT),
            )
            response.raise_for_status()
            return response.json()
        except httpx.TimeoutException as exc:
            raise LLMTimeoutError("模型请求超时") from exc
        except httpx.HTTPStatusError as exc:
            body = exc.response.text[:300]
            raise LLMError(self._upstream_error(exc.response.status_code, body)) from exc
        except httpx.RequestError as exc:
            raise LLMError(f"无法连接模型服务：{exc}") from exc
        except ValueError as exc:
            raise LLMError("模型服务返回了非 JSON 响应") from exc

    def _parse_stream_line(
        self, line: str, on_usage: Callable[[Usage], None] | None
    ) -> str:
        if not line.startswith("data:"):
            return ""
        data = line[5:].strip()
        if not data or data == "[DONE]":
            return ""
        try:
            chunk = json.loads(data)
        except json.JSONDecodeError:
            return ""
        if not isinstance(chunk, dict):
            return ""
        if on_usage is not None and chunk.get("usage"):
            try:
                on_usage(_parse_usage(chunk))
            except (TypeError, ValueError):
                pass
        choices = chunk.get("choices") or []
        if not choices:
            return ""
        delta = choices[0].get("delta") or {}
        content = delta.get("content")
        return content if isinstance(content, str) else ""

    @staticmethod
    def _upstream_error(status_code: int, body: str) -> str:
        detail = body.strip().replace("\n", " ")[:200]
        return f"模型服务返回错误（HTTP {status_code}）：{detail}"
