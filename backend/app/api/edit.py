"""Edit 接口（§18）：生成修改文本，返回 original_text / new_text。"""
from __future__ import annotations

import logging
import time
import traceback
import uuid
from datetime import datetime

from fastapi import APIRouter

from ..errors import BackendHTTPError
from ..logging_config import write_conversation_log
from ..models.edit import EditRequest, EditResponse
from ..services.edit_service import EditServiceError, generate_edit

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/v1", tags=["edit"])


@router.post("/edit", response_model=EditResponse)
async def edit(request: EditRequest) -> EditResponse:
    call_id = uuid.uuid4().hex
    started = time.perf_counter()
    try:
        response = await generate_edit(request)
        elapsed_ms = (time.perf_counter() - started) * 1000
        logger.info(
            "edit conversation=%s instruction_len=%d latency=%.2fs",
            request.conversation_id,
            len(request.instruction),
            elapsed_ms / 1000,
        )
        write_conversation_log({
            "timestamp": datetime.now().astimezone().isoformat(),
            "call_id": call_id,
            "kind": "edit",
            "conversation_id": request.conversation_id,
            "status": "success",
            "duration_ms": round(elapsed_ms, 2),
            "request": request,
            "response": response,
            "error": None,
        })
        return response
    except EditServiceError as exc:
        elapsed_ms = (time.perf_counter() - started) * 1000
        logger.warning(
            "edit 失败 conversation=%s code=%s：%s",
            request.conversation_id,
            exc.code,
            exc.message,
        )
        write_conversation_log({
            "timestamp": datetime.now().astimezone().isoformat(),
            "call_id": call_id,
            "kind": "edit",
            "conversation_id": request.conversation_id,
            "status": "error",
            "duration_ms": round(elapsed_ms, 2),
            "request": request,
            "response": None,
            "error": {"code": exc.code, "message": exc.message},
        })
        raise BackendHTTPError(502, exc.code, exc.message) from exc
    except Exception as exc:
        elapsed_ms = (time.perf_counter() - started) * 1000
        write_conversation_log({
            "timestamp": datetime.now().astimezone().isoformat(),
            "call_id": call_id,
            "kind": "edit",
            "conversation_id": request.conversation_id,
            "status": "error",
            "duration_ms": round(elapsed_ms, 2),
            "request": request,
            "response": None,
            "error": {
                "code": "BACKEND_ERROR",
                "message": str(exc),
                "exception_type": type(exc).__name__,
                "traceback": traceback.format_exc(),
            },
        })
        raise
