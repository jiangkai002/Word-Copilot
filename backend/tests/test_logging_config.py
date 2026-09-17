from __future__ import annotations

import json

from app.logging_config import ConversationAuditWriter


def test_conversation_writer_writes_valid_jsonl_and_truncates(tmp_path):
    path = tmp_path / "conversations.jsonl"
    writer = ConversationAuditWriter(
        path,
        max_bytes=1024 * 1024,
        backup_count=1,
        include_content=True,
        max_text_chars=5,
    )

    writer.write({"kind": "chat", "request": {"message": "123456789"}})

    record = json.loads(path.read_text(encoding="utf-8").strip())
    assert record["kind"] == "chat"
    assert record["request"]["message"].startswith("12345…[截断")


def test_conversation_writer_can_omit_content(tmp_path):
    path = tmp_path / "conversations.jsonl"
    writer = ConversationAuditWriter(
        path,
        max_bytes=1024 * 1024,
        backup_count=1,
        include_content=False,
        max_text_chars=100,
    )

    writer.write({"kind": "chat", "request": {"message": "敏感正文"}})

    record = json.loads(path.read_text(encoding="utf-8").strip())
    assert record["kind"] == "chat"
    assert record["request"]["message"] == {"omitted": True, "length": 4}
