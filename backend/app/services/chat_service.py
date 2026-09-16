"""Chat 编排：组装消息（系统提示 + 历史 + 上下文 + 提问）并流式产出 token。"""
from __future__ import annotations

import logging
from collections.abc import AsyncIterator, Callable

from ..llm.base import Usage
from ..llm.factory import get_provider
from ..models.chat import ChatStreamRequest
from . import context_service, prompt_service

logger = logging.getLogger(__name__)

_FALLBACK_SYSTEM_PROMPT = (
    "你是 Word 文档 AI 助手。用中文简洁、准确地回答用户关于文档内容的问题，"
    "不编造文档中不存在的内容。"
)


def build_messages(request: ChatStreamRequest) -> list[dict[str, str]]:
    system_prompt = prompt_service.load_prompt("chat.md") or _FALLBACK_SYSTEM_PROMPT
    messages: list[dict[str, str]] = [{"role": "system", "content": system_prompt}]
    for entry in request.history[-10:]:
        messages.append({"role": entry.role, "content": entry.content})

    context_block = context_service.render_context(request.context)
    user_content = (
        f"{context_block}\n\n【用户提问】\n{request.message}"
        if context_block
        else request.message
    )
    messages.append({"role": "user", "content": user_content})
    return messages


async def stream_chat(
    request: ChatStreamRequest,
    *,
    on_usage: Callable[[Usage], None] | None = None,
) -> AsyncIterator[str]:
    provider = get_provider()
    messages = build_messages(request)
    async for token in provider.stream_chat(messages, on_usage=on_usage):
        yield token
