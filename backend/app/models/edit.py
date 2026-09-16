"""Edit 请求 / 响应模型（§18 / §37）。

Edit 接口 wire 格式为 snake_case（与需求文档示例一致）。
后端只返回 original_text / new_text —— 字符级 Diff 由 Word 客户端完成（§19）。
"""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class EditTargetPayload(BaseModel):
    type: Literal["selection", "paragraph"] = "paragraph"
    text: str = Field(min_length=1, max_length=50000)
    text_hash: str = ""


class EditContextPayload(BaseModel):
    previous: str | None = None
    next: str | None = None


class EditRequest(BaseModel):
    conversation_id: str
    instruction: str = Field(min_length=1, max_length=2000)
    target: EditTargetPayload
    context: EditContextPayload | None = None
    regenerate_count: int = 0
    avoid_texts: list[str] = Field(default_factory=list)


class EditResponse(BaseModel):
    """§37：模型输出必须经 Pydantic 校验后才返回客户端。"""

    edit_id: str
    original_text: str
    new_text: str
    summary: str
    warnings: list[str] = Field(default_factory=list)


class EditCandidate(BaseModel):
    """期望模型返回的 JSON 结构（original_text 必须逐字符回显原文）。"""

    summary: str = ""
    original_text: str
    new_text: str
    warnings: list[str] = Field(default_factory=list)
