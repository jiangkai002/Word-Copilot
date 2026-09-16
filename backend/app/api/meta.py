"""元信息接口（§44）：GET /api/v1/health、GET /api/v1/config（可选）。"""
from __future__ import annotations

from fastapi import APIRouter

from .. import config
from ..llm.factory import get_provider

router = APIRouter(prefix="/api/v1", tags=["meta"])


@router.get("/health")
async def health() -> dict:
    return {"status": "ok", "version": config.APP_VERSION}


@router.get("/config")
async def get_config() -> dict:
    provider = get_provider()
    return {
        "model": provider.model_name,
        "provider": config.LLM_PROVIDER,
        "supports_stream": provider.supports_stream,
    }
