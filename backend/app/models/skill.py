"""后台可配置 Skill 的请求与持久化模型。"""
from __future__ import annotations

from datetime import datetime

from pydantic import BaseModel, Field, field_validator


class SkillPayload(BaseModel):
    """创建、更新 Skill 时由管理页提交的完整配置。"""

    name: str = Field(min_length=1, max_length=80)
    description: str = Field(default="", max_length=500)
    trigger_phrases: list[str] = Field(default_factory=list, max_length=20)
    instructions: str = Field(min_length=1, max_length=12000)
    enabled: bool = True

    @field_validator("name", "instructions")
    @classmethod
    def strip_required_text(cls, value: str) -> str:
        stripped = value.strip()
        if not stripped:
            raise ValueError("内容不能为空")
        return stripped

    @field_validator("description")
    @classmethod
    def strip_description(cls, value: str) -> str:
        return value.strip()

    @field_validator("trigger_phrases")
    @classmethod
    def normalize_triggers(cls, values: list[str]) -> list[str]:
        result: list[str] = []
        seen: set[str] = set()
        for raw in values:
            value = raw.strip()
            if not value:
                continue
            if len(value) > 50:
                raise ValueError("单个触发词不能超过 50 个字符")
            key = value.casefold()
            if key not in seen:
                seen.add(key)
                result.append(value)
        return result


class SkillRecord(SkillPayload):
    """返回给管理页的 Skill 记录。"""

    id: str
    created_at: datetime
    updated_at: datetime
