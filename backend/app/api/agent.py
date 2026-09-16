"""Agent SSE 流式接口（§51：批量编辑 + 文档问答 + 格式 / 表格 / 公式 / 段落，双事件通道）。"""
from __future__ import annotations

import asyncio
import json
import logging
import time

from fastapi import APIRouter
from fastapi.responses import StreamingResponse

from ..models.agent import AgentStreamRequest
from ..services.agent_service import (
    AgentServiceError,
    DoneEvent,
    FormulaProposalEvent,
    FormatProposalEvent,
    ParagraphProposalEvent,
    ProposalEvent,
    TableProposalEvent,
    TokenEvent,
    run_agent_stream,
)

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/v1", tags=["agent"])


def _sse_event(event: str, payload: dict) -> str:
    data = json.dumps(payload, ensure_ascii=False)
    return f"event: {event}\ndata: {data}\n\n"


@router.post("/agent/stream")
async def agent_stream(request: AgentStreamRequest) -> StreamingResponse:
    """SSE 流：token（回答增量）/ proposal（修改提案）→ done；错误以 error 帧结束。

    提案帧发出后即视为已送达客户端 —— 即使流中途出错（超时等），
    已收到的提案仍会由前端逐条走独立校验管线（§15 乐观锁兜底）。
    """

    async def event_generator():
        started = time.perf_counter()
        token_count = 0
        proposal_count = 0
        try:
            async for event in run_agent_stream(request):
                if isinstance(event, TokenEvent):
                    token_count += 1
                    yield _sse_event("token", {"text": event.text})
                elif isinstance(event, ProposalEvent):
                    proposal_count += 1
                    yield _sse_event("proposal", event.model_dump())
                elif isinstance(event, FormatProposalEvent):
                    proposal_count += 1
                    yield _sse_event("proposal_format", event.model_dump())
                elif isinstance(event, TableProposalEvent):
                    proposal_count += 1
                    yield _sse_event("proposal_table", event.model_dump())
                elif isinstance(event, FormulaProposalEvent):
                    proposal_count += 1
                    yield _sse_event("proposal_formula", event.model_dump())
                elif isinstance(event, ParagraphProposalEvent):
                    proposal_count += 1
                    yield _sse_event("proposal_paragraph", event.model_dump())
                elif isinstance(event, DoneEvent):
                    yield _sse_event(
                        "done",
                        {"usage": event.usage, "proposal_count": proposal_count},
                    )
        except AgentServiceError as exc:
            yield _sse_event("error", {"code": exc.code, "message": exc.message})
        except asyncio.CancelledError:
            # 客户端停止生成（AbortController）—— 正常结束，不发送事件
            raise
        except Exception:
            logger.exception("agent 流内部错误 conversation=%s", request.conversation_id)
            yield _sse_event("error", {"code": "AGENT_ERROR", "message": "Agent 内部错误，请查看后端日志"})
        finally:
            logger.info(
                "agent conversation=%s tokens=%d proposals=%d latency=%.2fs",
                request.conversation_id,
                token_count,
                proposal_count,
                time.perf_counter() - started,
            )

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream; charset=utf-8",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
