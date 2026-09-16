"""FastAPI 应用入口（§34 / §44 / §45 / §61）。

启动：cd backend && uvicorn app.main:app --reload --port 8000
"""
from __future__ import annotations

import logging
import time
import uuid

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from . import config
from .api import agent, chat, edit, meta, skills
from .errors import BackendHTTPError
from .llm.base import LLMError, LLMTimeoutError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
)
logger = logging.getLogger("word-ai-backend")

app = FastAPI(title="Word AI Copilot Backend", version=config.APP_VERSION)

# CORS（§61）：默认仅放行本地开发 Origin；生产环境通过 BACKEND_CORS_ORIGINS 收紧
if config.BACKEND_CORS_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=config.BACKEND_CORS_ORIGINS,
        allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        allow_headers=["Content-Type", "Accept"],
    )


@app.middleware("http")
async def request_logging(request: Request, call_next):
    """记录 request_id / 路径 / 状态 / 时延（§59：不记录文档正文）。"""
    request_id = uuid.uuid4().hex[:12]
    started = time.perf_counter()
    response = await call_next(request)
    latency_ms = (time.perf_counter() - started) * 1000
    logger.info(
        "request_id=%s %s %s -> %d %.1fms",
        request_id,
        request.method,
        request.url.path,
        response.status_code,
        latency_ms,
    )
    response.headers["X-Request-ID"] = request_id
    return response


# ---- 统一错误格式（§45）----


def _error_response(status_code: int, code: str, message: str) -> JSONResponse:
    return JSONResponse(
        status_code=status_code,
        content={"error": {"code": code, "message": message}},
    )


@app.exception_handler(BackendHTTPError)
async def backend_http_error_handler(_: Request, exc: BackendHTTPError) -> JSONResponse:
    return _error_response(exc.status_code, exc.code, exc.message)


@app.exception_handler(RequestValidationError)
async def validation_error_handler(_: Request, exc: RequestValidationError) -> JSONResponse:
    first = exc.errors()[0] if exc.errors() else {}
    loc = ".".join(str(part) for part in first.get("loc", []))
    return _error_response(422, "INVALID_REQUEST", f"请求参数不合法：{loc or first.get('msg', '')}")


@app.exception_handler(LLMTimeoutError)
async def llm_timeout_handler(_: Request, exc: LLMTimeoutError) -> JSONResponse:
    return _error_response(504, "LLM_TIMEOUT", "模型请求超时")


@app.exception_handler(LLMError)
async def llm_error_handler(_: Request, exc: LLMError) -> JSONResponse:
    return _error_response(502, "LLM_ERROR", exc.message)


app.include_router(meta.router)
app.include_router(chat.router)
app.include_router(edit.router)
app.include_router(agent.router)
app.include_router(skills.router)
