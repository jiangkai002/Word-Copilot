"""Skill JSON 持久化、触发匹配与 Prompt 渲染。"""
from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import UTC, datetime
from pathlib import Path

from ..models.skill import SkillPayload, SkillRecord

_BACKEND_ROOT = Path(__file__).resolve().parents[2]
_SKILLS_PATH = Path(
    os.getenv("SKILL_STORE_PATH", str(_BACKEND_ROOT / "data" / "skills.json"))
)
_LOCK = threading.RLock()


def _read_unlocked() -> list[SkillRecord]:
    if not _SKILLS_PATH.exists():
        return []
    try:
        raw = json.loads(_SKILLS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"无法读取 Skill 配置：{exc}") from exc
    if not isinstance(raw, list):
        raise RuntimeError("Skill 配置文件格式错误：根节点必须是数组")
    return [SkillRecord.model_validate(item) for item in raw]


def _write_unlocked(skills: list[SkillRecord]) -> None:
    _SKILLS_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = _SKILLS_PATH.with_suffix(f"{_SKILLS_PATH.suffix}.tmp")
    payload = [skill.model_dump(mode="json") for skill in skills]
    temp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temp_path.replace(_SKILLS_PATH)


def list_skills() -> list[SkillRecord]:
    with _LOCK:
        return sorted(_read_unlocked(), key=lambda item: item.updated_at, reverse=True)


def create_skill(payload: SkillPayload) -> SkillRecord:
    now = datetime.now(UTC)
    record = SkillRecord(
        id=uuid.uuid4().hex,
        created_at=now,
        updated_at=now,
        **payload.model_dump(),
    )
    with _LOCK:
        skills = _read_unlocked()
        skills.append(record)
        _write_unlocked(skills)
    return record


def update_skill(skill_id: str, payload: SkillPayload) -> SkillRecord | None:
    with _LOCK:
        skills = _read_unlocked()
        for index, current in enumerate(skills):
            if current.id != skill_id:
                continue
            updated = SkillRecord(
                id=current.id,
                created_at=current.created_at,
                updated_at=datetime.now(UTC),
                **payload.model_dump(),
            )
            skills[index] = updated
            _write_unlocked(skills)
            return updated
    return None


def delete_skill(skill_id: str) -> bool:
    with _LOCK:
        skills = _read_unlocked()
        remaining = [skill for skill in skills if skill.id != skill_id]
        if len(remaining) == len(skills):
            return False
        _write_unlocked(remaining)
        return True


def applicable_skills(user_text: str) -> list[SkillRecord]:
    """返回已启用且命中当前请求的 Skill；无触发词表示始终启用。"""
    normalized = user_text.casefold()
    result: list[SkillRecord] = []
    for skill in list_skills():
        if not skill.enabled:
            continue
        if not skill.trigger_phrases or any(
            phrase.casefold() in normalized for phrase in skill.trigger_phrases
        ):
            result.append(skill)
    return result


def render_skill_prompt(user_text: str) -> str:
    """把当前请求适用的 Skill 渲染成系统提示词追加块。"""
    skills = applicable_skills(user_text)
    if not skills:
        return ""
    sections = ["## 后台配置的 Skills", "以下是管理员为当前请求启用的额外工作指令："]
    for skill in skills:
        sections.append(f"\n### {skill.name}")
        if skill.description:
            sections.append(skill.description)
        sections.append(skill.instructions)
    return "\n".join(sections)

