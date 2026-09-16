"""Chat 请求模型（§16 / §43）。

前端以 camelCase 发送上下文字段（contextType / currentParagraph / ...），
这里用 alias_generator 统一映射为 Python snake_case。
"""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


def to_camel(name: str) -> str:
    head, *rest = name.split("_")
    return head + "".join(part.title() for part in rest)


class CamelModel(BaseModel):
    """接受 camelCase（前端 wire 格式）与 snake_case（内部调用）两种键名。"""

    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True, extra="ignore")


class SelectionPayload(CamelModel):
    text: str = ""


class ParagraphPayload(CamelModel):
    text: str = ""
    style: str | None = None
    unique_local_id: str | None = None
    index: int | None = None
    text_hash: str = ""


class SectionPayload(CamelModel):
    title: str = ""
    level: int = 1
    content: str = ""


class OutlineItemPayload(CamelModel):
    level: int = 1
    title: str = ""


class DocumentContextPayload(CamelModel):
    context_type: Literal["selection", "paragraph", "section", "document"] = "paragraph"
    selection: SelectionPayload | None = None
    current_paragraph: ParagraphPayload | None = None
    previous_paragraph: ParagraphPayload | None = None
    next_paragraph: ParagraphPayload | None = None
    section: SectionPayload | None = None
    outline: list[OutlineItemPayload] = Field(default_factory=list)
    document_text: str | None = None
    truncated: bool = False


class HistoryEntryPayload(BaseModel):
    role: Literal["user", "assistant"]
    content: str


class ChatStreamRequest(CamelModel):
    conversation_id: str
    message: str = Field(min_length=1, max_length=20000)
    context: DocumentContextPayload | None = None
    history: list[HistoryEntryPayload] = Field(default_factory=list)
