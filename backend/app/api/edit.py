"""Edit 接口（§18）：生成修改文本，返回 original_text / new_text。"""
from __future__ import annotations

import logging
import time

from fastapi import APIRouter

from ..errors import BackendHTTPError
from ..models.edit import EditRequest, EditResponse
from ..services.edit_service import EditServiceError, generate_edit

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/v1", tags=["edit"])


@router.post("/edit", response_model=EditResponse)
async def edit(request: EditRequest) -> EditResponse:
    started = time.perf_counter()
    try:
        response = await generate_edit(request)
        logger.info(
            "edit conversation=%s instruction_len=%d latency=%.2fs",
            request.conversation_id,
            len(request.instruction),
            time.perf_counter() - started,
        )
        return response
    except EditServiceError as exc:
        logger.warning(
            "edit 失败 conversation=%s code=%s：%s",
            request.conversation_id,
            exc.code,
            exc.message,
        )
        raise BackendHTTPError(502, exc.code, exc.message) from exc
