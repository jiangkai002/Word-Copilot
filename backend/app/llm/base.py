"""LLMProvider 抽象基类与统一异常（§34 / §45）。"""
from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import AsyncIterator, Callable
from dataclasses import dataclass, field


class LLMError(Exception):
    """模型调用失败（映射错误码 LLM_ERROR）。"""

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


class LLMTimeoutError(LLMError):
    """模型请求超时（映射错误码 LLM_TIMEOUT）。"""


@dataclass
class Usage:
    prompt_tokens: int = 0
    completion_tokens: int = 0

    @property
    def total_tokens(self) -> int:
        return self.prompt_tokens + self.completion_tokens

    def as_dict(self) -> dict[str, int]:
        return {
            "prompt_tokens": self.prompt_tokens,
            "completion_tokens": self.completion_tokens,
        }


@dataclass
class CompletionResult:
    """非流式补全结果。"""

    content: str
    usage: Usage = field(default_factory=Usage)


class LLMProvider(ABC):
    """所有 Provider 的统一接口（§34）。"""

    @abstractmethod
    async def chat(
        self,
        messages: list[dict[str, str]],
        *,
        temperature: float | None = None,
        max_tokens: int | None = None,
        timeout: float | None = None,
    ) -> CompletionResult:
        """非流式补全（Edit 生成用）。"""

    @abstractmethod
    def stream_chat(
        self,
        messages: list[dict[str, str]],
        *,
        temperature: float | None = None,
        max_tokens: int | None = None,
        timeout: float | None = None,
        on_usage: Callable[[Usage], None] | None = None,
    ) -> AsyncIterator[str]:
        """流式补全（Chat SSE 用），逐段 yield 文本增量。"""

    @property
    @abstractmethod
    def model_name(self) -> str:
        """当前使用的模型名（用于 /config 与日志）。"""

    @property
    @abstractmethod
    def supports_stream(self) -> bool:
        """是否支持流式输出。"""
