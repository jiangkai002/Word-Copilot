"""本地运行日志与对话审计日志。

普通日志写入 backend.log；每次 LLM 调用的请求、结果和失败写入
conversations.jsonl。两类文件均使用大小轮转，且 logs/ 已被 Git 忽略。
"""
from __future__ import annotations

import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from threading import Lock
from typing import Any

from . import config

_FILE_HANDLER_MARKER = "_word_ai_local_file_handler"
_audit_writer: "ConversationAuditWriter | None" = None
_audit_lock = Lock()


def setup_logging() -> None:
    """配置控制台和本地轮转文件日志；重复调用不会添加重复 handler。"""
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    level = getattr(logging, config.LOG_LEVEL, logging.INFO)
    root = logging.getLogger()
    root.setLevel(level)

    if not root.handlers:
        console = logging.StreamHandler()
        console.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(name)s: %(message)s"))
        root.addHandler(console)

    if not any(getattr(handler, _FILE_HANDLER_MARKER, False) for handler in root.handlers):
        file_handler = RotatingFileHandler(
            config.LOG_DIR / "backend.log",
            maxBytes=max(1, config.LOG_MAX_BYTES),
            backupCount=max(0, config.LOG_BACKUP_COUNT),
            encoding="utf-8",
        )
        setattr(file_handler, _FILE_HANDLER_MARKER, True)
        file_handler.setLevel(level)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)-7s %(name)s: %(message)s")
        )
        root.addHandler(file_handler)


def _bounded(value: Any, *, max_text_chars: int) -> Any:
    """把 Pydantic/异常等值转换为 JSON 可写结构，并限制超长文本。"""
    if hasattr(value, "model_dump"):
        value = value.model_dump(mode="json")
    if isinstance(value, dict):
        return {
            str(key): _bounded(item, max_text_chars=max_text_chars)
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple)):
        return [
            _bounded(item, max_text_chars=max_text_chars)
            for item in value
        ]
    if isinstance(value, str):
        if len(value) > max_text_chars:
            return value[:max_text_chars] + f"…[截断 {len(value) - max_text_chars} 字符]"
        return value
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return str(value)


def _content_lengths(value: Any) -> Any:
    """隐私模式下保留结构和长度，但移除可能包含正文的字符串。"""
    if hasattr(value, "model_dump"):
        value = value.model_dump(mode="json")
    if isinstance(value, dict):
        return {str(key): _content_lengths(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_content_lengths(item) for item in value]
    if isinstance(value, str):
        return {"omitted": True, "length": len(value)}
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return {"omitted": True, "type": type(value).__name__}


class ConversationAuditWriter:
    """线程安全的 JSONL 对话日志写入器。"""

    def __init__(
        self,
        path: Path,
        *,
        max_bytes: int,
        backup_count: int,
        include_content: bool,
        max_text_chars: int,
    ) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.include_content = include_content
        self.max_text_chars = max(1, max_text_chars)
        self.logger = logging.getLogger(f"word-ai-conversation-audit:{path}")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        if not self.logger.handlers:
            handler = RotatingFileHandler(
                path,
                maxBytes=max(1, max_bytes),
                backupCount=max(0, backup_count),
                encoding="utf-8",
            )
            handler.setFormatter(logging.Formatter("%(message)s"))
            self.logger.addHandler(handler)

    def write(self, record: dict[str, Any]) -> None:
        source = dict(record)
        if not self.include_content:
            for key in ("request", "response", "error", "message", "details"):
                if key in source and source[key] is not None:
                    source[key] = _content_lengths(source[key])
        payload = _bounded(
            source,
            max_text_chars=self.max_text_chars,
        )
        self.logger.info(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


def get_conversation_audit_writer() -> ConversationAuditWriter | None:
    """延迟创建默认审计写入器；关闭开关时返回 None。"""
    global _audit_writer
    if not config.CONVERSATION_LOG_ENABLED:
        return None
    if _audit_writer is None:
        with _audit_lock:
            if _audit_writer is None:
                _audit_writer = ConversationAuditWriter(
                    config.LOG_DIR / "conversations.jsonl",
                    max_bytes=config.LOG_MAX_BYTES,
                    backup_count=config.LOG_BACKUP_COUNT,
                    include_content=config.CONVERSATION_LOG_INCLUDE_CONTENT,
                    max_text_chars=config.CONVERSATION_LOG_MAX_TEXT_CHARS,
                )
    return _audit_writer


def write_conversation_log(record: dict[str, Any]) -> None:
    try:
        if record.get("kind") in {"chat", "agent", "edit"}:
            record.setdefault("llm", {"provider": config.LLM_PROVIDER, "model": config.LLM_MODEL})
        writer = get_conversation_audit_writer()
        if writer is not None:
            writer.write(record)
    except Exception:
        # 日志系统绝不能反向导致对话/编辑请求失败。
        logging.getLogger(__name__).exception("写入对话审计日志失败")
