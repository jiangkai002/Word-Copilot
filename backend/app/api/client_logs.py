"""接收 Word 任务窗格错误并写入本地结构化日志。"""
from __future__ import annotations

import uuid

from fastapi import APIRouter, Response, status

from ..logging_config import write_conversation_log
from ..models.client_log import ClientLogEntry

router = APIRouter(prefix="/api/v1", tags=["logs"])


@router.post("/logs/client", status_code=status.HTTP_204_NO_CONTENT)
async def client_log(entry: ClientLogEntry) -> Response:
    write_conversation_log({
        "timestamp": entry.timestamp,
        "call_id": uuid.uuid4().hex,
        "kind": "client",
        "conversation_id": entry.conversation_id,
        "status": "error" if entry.level == "error" else "warning",
        "level": entry.level,
        "message": entry.message,
        "details": entry.details,
        "page": entry.page,
    })
    return Response(status_code=status.HTTP_204_NO_CONTENT)
