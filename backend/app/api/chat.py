"""Chat SSE 流式接口（§16 / §43）。"""
from __future__ import annotations

import asyncio
import json
import logging
import time

from fastapi import APIRouter, Request
from fastapi.responses import StreamingResponse

from ..llm.base import LLMError, LLMTimeoutError, Usage
from ..models.chat import ChatStreamRequest
from ..services import chat_service

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/v1", tags=["chat"])


def _sse_event(event: str, payload: dict) -> str:
    data = json.dumps(payload, ensure_ascii=False)
    return f"event: {event}\ndata: {data}\n\n"


@router.post("/chat/stream")
async def chat_stream(request: ChatStreamRequest) -> StreamingResponse:
    """SSE 流：event: token → … → event: done；中途错误用 event: error 结束。"""

    async def event_generator():
        usage: Usage = Usage()
        started = time.perf_counter()
        token_count = 0

        def on_usage(u: Usage) -> None:
            usage.prompt_tokens = u.prompt_tokens
            usage.completion_tokens = u.completion_tokens

        try:
            async for token in chat_service.stream_chat(request, on_usage=on_usage):
                if token:
                    token_count += 1
                    yield _sse_event("token", {"text": token})
            yield _sse_event("done", {"usage": usage.as_dict()})
        except LLMTimeoutError:
            yield _sse_event("error", {"code": "LLM_TIMEOUT", "message": "模型请求超时"})
        except LLMError as exc:
            yield _sse_event("error", {"code": "LLM_ERROR", "message": exc.message})
        except asyncio.CancelledError:
            # 客户端停止生成（AbortController）—— 正常结束，不发送事件
            raise
        except Exception:
            logger.exception("chat 流内部错误 conversation=%s", request.conversation_id)
            yield _sse_event("error", {"code": "BACKEND_ERROR", "message": "服务内部错误"})
        finally:
            logger.info(
                "chat conversation=%s tokens_events=%d usage=%s latency=%.2fs",
                request.conversation_id,
                token_count,
                usage.total_tokens,
                time.perf_counter() - started,
            )

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream; charset=utf-8",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
