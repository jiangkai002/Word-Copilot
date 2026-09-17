"""前端 Word/Office.js 运行日志上报模型。"""
from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class ClientLogEntry(BaseModel):
    level: Literal["warn", "error"]
    timestamp: str
    conversation_id: str | None = None
    message: str = Field(max_length=4000)
    details: list[Any] = Field(default_factory=list, max_length=20)
    page: str | None = Field(default=None, max_length=500)
