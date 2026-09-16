from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.models.skill import SkillPayload
from app.services import skill_service


@pytest.fixture(autouse=True)
def isolated_store(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    path = tmp_path / "skills.json"
    monkeypatch.setattr(skill_service, "_SKILLS_PATH", path)
    return path


def payload(**overrides: object) -> SkillPayload:
    values: dict[str, object] = {
        "name": "合同审查",
        "description": "检查合同中的风险条款",
        "trigger_phrases": ["合同", "条款"],
        "instructions": "逐条列出风险，并给出修改建议。",
        "enabled": True,
    }
    values.update(overrides)
    return SkillPayload.model_validate(values)


def test_skill_crud_and_json_persistence(isolated_store: Path) -> None:
    created = skill_service.create_skill(payload())
    assert isolated_store.exists()
    assert skill_service.list_skills()[0].id == created.id

    updated = skill_service.update_skill(created.id, payload(name="合同风险复核", enabled=False))
    assert updated is not None
    assert updated.name == "合同风险复核"
    assert updated.enabled is False

    assert skill_service.delete_skill(created.id) is True
    assert skill_service.delete_skill(created.id) is False
    assert skill_service.list_skills() == []


def test_skill_trigger_and_prompt_rendering() -> None:
    skill_service.create_skill(payload())
    skill_service.create_skill(
        payload(
            name="通用写作规范",
            trigger_phrases=[],
            instructions="使用简洁、正式的中文。",
        )
    )
    skill_service.create_skill(payload(name="停用项", trigger_phrases=[], enabled=False))

    matched = skill_service.applicable_skills("请审查这份合同")
    assert {item.name for item in matched} == {"合同审查", "通用写作规范"}
    unmatched = skill_service.applicable_skills("总结会议纪要")
    assert [item.name for item in unmatched] == ["通用写作规范"]

    prompt = skill_service.render_skill_prompt("检查合同条款")
    assert "### 合同审查" in prompt
    assert "逐条列出风险" in prompt
    assert "停用项" not in prompt


def test_skill_rejects_blank_required_fields() -> None:
    with pytest.raises(ValueError):
        payload(name="   ")
    with pytest.raises(ValueError):
        payload(instructions="\n\t")


def test_skill_api_crud() -> None:
    client = TestClient(app)
    response = client.post("/api/v1/skills", json=payload().model_dump(mode="json"))
    assert response.status_code == 201
    skill_id = response.json()["id"]

    assert client.get("/api/v1/skills").json()[0]["name"] == "合同审查"
    updated_payload = payload(enabled=False).model_dump(mode="json")
    assert client.put(f"/api/v1/skills/{skill_id}", json=updated_payload).json()["enabled"] is False
    assert client.delete(f"/api/v1/skills/{skill_id}").status_code == 204
    assert client.put("/api/v1/skills/missing", json=updated_payload).status_code == 404


def test_admin_page_is_served() -> None:
    client = TestClient(app)
    response = client.get("/admin/skills")
    assert response.status_code == 200
    assert "Skill 配置" in response.text
