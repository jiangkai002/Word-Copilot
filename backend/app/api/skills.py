"""Skill 管理 API 与单页后台入口。"""
from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import FileResponse, Response

from ..errors import BackendHTTPError
from ..models.skill import SkillPayload, SkillRecord
from ..services import skill_service

router = APIRouter(tags=["skills"])
_ADMIN_PAGE = Path(__file__).resolve().parent.parent / "admin" / "skills.html"


@router.get("/admin/skills", include_in_schema=False)
async def skills_admin_page() -> FileResponse:
    return FileResponse(_ADMIN_PAGE, media_type="text/html; charset=utf-8")


@router.get("/api/v1/skills", response_model=list[SkillRecord])
async def get_skills() -> list[SkillRecord]:
    return skill_service.list_skills()


@router.post("/api/v1/skills", response_model=SkillRecord, status_code=201)
async def post_skill(payload: SkillPayload) -> SkillRecord:
    return skill_service.create_skill(payload)


@router.put("/api/v1/skills/{skill_id}", response_model=SkillRecord)
async def put_skill(skill_id: str, payload: SkillPayload) -> SkillRecord:
    updated = skill_service.update_skill(skill_id, payload)
    if updated is None:
        raise BackendHTTPError(404, "SKILL_NOT_FOUND", "Skill 不存在")
    return updated


@router.delete("/api/v1/skills/{skill_id}", status_code=204)
async def remove_skill(skill_id: str) -> Response:
    if not skill_service.delete_skill(skill_id):
        raise BackendHTTPError(404, "SKILL_NOT_FOUND", "Skill 不存在")
    return Response(status_code=204)

