from __future__ import annotations

from fastapi.testclient import TestClient

from app.api import agent as agent_api
from app.api import chat as chat_api
from app.api import edit as edit_api
from app.llm.base import Usage
from app.main import app
from app.services.agent_service import DoneEvent, TokenEvent
from app.services.edit_service import EditServiceError

client = TestClient(app)


def test_chat_success_is_written_to_conversation_log(monkeypatch):
    records: list[dict] = []

    async def fake_stream(_request, *, on_usage=None):
        if on_usage:
            on_usage(Usage(prompt_tokens=2, completion_tokens=1))
        yield "回答"

    monkeypatch.setattr(chat_api.chat_service, "stream_chat", fake_stream)
    monkeypatch.setattr(chat_api, "write_conversation_log", records.append)

    response = client.post(
        "/api/v1/chat/stream",
        json={"conversation_id": "c-chat", "message": "问题"},
    )

    assert response.status_code == 200
    assert "event: done" in response.text
    assert records[0]["kind"] == "chat"
    assert records[0]["status"] == "success"
    assert records[0]["response"]["text"] == "回答"
    assert records[0]["response"]["usage"] == {"prompt_tokens": 2, "completion_tokens": 1}


def test_agent_success_is_written_to_conversation_log(monkeypatch):
    records: list[dict] = []

    async def fake_agent(_request):
        yield TokenEvent(text="完成")
        yield DoneEvent(usage={"prompt_tokens": 3, "completion_tokens": 2})

    monkeypatch.setattr(agent_api, "run_agent_stream", fake_agent)
    monkeypatch.setattr(agent_api, "write_conversation_log", records.append)

    response = client.post(
        "/api/v1/agent/stream",
        json={
            "conversation_id": "c-agent",
            "instruction": "检查",
            "snapshot": {"outline": [], "paragraphs": [], "truncated": False},
        },
    )

    assert response.status_code == 200
    assert "event: done" in response.text
    assert records[0]["kind"] == "agent"
    assert records[0]["status"] == "success"
    assert records[0]["response"]["text"] == "完成"


def test_edit_failure_is_written_to_conversation_log(monkeypatch):
    records: list[dict] = []

    async def fake_edit(_request):
        raise EditServiceError("INVALID_EDIT_RESPONSE", "返回格式错误")

    monkeypatch.setattr(edit_api, "generate_edit", fake_edit)
    monkeypatch.setattr(edit_api, "write_conversation_log", records.append)

    response = client.post(
        "/api/v1/edit",
        json={
            "conversation_id": "c-edit",
            "instruction": "润色",
            "target": {"type": "selection", "text": "原文", "text_hash": "hash"},
        },
    )

    assert response.status_code == 502
    assert records[0]["kind"] == "edit"
    assert records[0]["status"] == "error"
    assert records[0]["error"]["code"] == "INVALID_EDIT_RESPONSE"
