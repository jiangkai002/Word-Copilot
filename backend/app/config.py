"""环境配置：读取 backend/.env（python-dotenv），不引入额外框架。"""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# backend/.env（config.py 位于 app/ 下，父目录即 backend 根）
_BACKEND_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(_BACKEND_ROOT / ".env")
load_dotenv()  # 兼容从其他 cwd 启动时读取当前目录 .env

APP_VERSION = "0.1.0"


def _get_str(key: str, default: str = "") -> str:
    value = os.getenv(key)
    return value.strip() if value and value.strip() else default


def _get_float(key: str, default: float) -> float:
    try:
        return float(_get_str(key, str(default)))
    except ValueError:
        return default


def _get_int(key: str, default: int) -> int:
    try:
        return int(_get_str(key, str(default)))
    except ValueError:
        return default


def _get_bool(key: str, default: bool) -> bool:
    raw = _get_str(key)
    if not raw:
        return default
    return raw.lower() in {"1", "true", "yes", "on"}


def _get_origins(key: str, default: list[str]) -> list[str]:
    raw = _get_str(key, ",".join(default))
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


# ---- LLM ----
LLM_PROVIDER: str = _get_str("LLM_PROVIDER", "openai")
LLM_BASE_URL: str = _get_str("LLM_BASE_URL", "https://api.openai.com/v1")
LLM_API_KEY: str = _get_str("LLM_API_KEY")
LLM_MODEL: str = _get_str("LLM_MODEL")

LLM_TIMEOUT_SECONDS: float = _get_float("LLM_TIMEOUT_SECONDS", 120.0)
LLM_EDIT_TIMEOUT_SECONDS: float = _get_float("LLM_EDIT_TIMEOUT_SECONDS", 60.0)
LLM_MAX_TOKENS: int | None = _get_int("LLM_MAX_TOKENS", 2048)
LLM_TEMPERATURE: float | None = _get_float("LLM_TEMPERATURE", 0.3)

# ---- Agent（§51 Microsoft Agent Framework：批量编辑 + 文档问答）----
AGENT_MAX_PROPOSALS: int = _get_int("AGENT_MAX_PROPOSALS", 20)
AGENT_TIMEOUT_SECONDS: float = _get_float("AGENT_TIMEOUT_SECONDS", 300.0)
AGENT_MAX_TOOL_CALLS: int = _get_int("AGENT_MAX_TOOL_CALLS", 40)

# ---- CORS（§61：生产环境禁止 *）----
BACKEND_CORS_ORIGINS: list[str] = _get_origins(
    "BACKEND_CORS_ORIGINS",
    ["http://localhost:3000", "https://localhost:3000"],
)

# ---- 本地日志 ----
_log_dir_value = Path(_get_str("LOG_DIR", "logs"))
LOG_DIR: Path = (
    _log_dir_value if _log_dir_value.is_absolute() else _BACKEND_ROOT / _log_dir_value
).resolve()
LOG_LEVEL: str = _get_str("LOG_LEVEL", "INFO").upper()
LOG_MAX_BYTES: int = _get_int("LOG_MAX_BYTES", 10 * 1024 * 1024)
LOG_BACKUP_COUNT: int = _get_int("LOG_BACKUP_COUNT", 5)
CONVERSATION_LOG_ENABLED: bool = _get_bool("CONVERSATION_LOG_ENABLED", True)
CONVERSATION_LOG_INCLUDE_CONTENT: bool = _get_bool("CONVERSATION_LOG_INCLUDE_CONTENT", True)
CONVERSATION_LOG_MAX_TEXT_CHARS: int = _get_int("CONVERSATION_LOG_MAX_TEXT_CHARS", 20_000)


def llm_configured() -> bool:
    """LLM 是否已配置（缺 API Key 或模型名时调用方应给出明确错误）。"""
    return bool(LLM_API_KEY and LLM_MODEL)
