"""Prompt 模板加载（app/prompts/*.md，进程内缓存）。"""
from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"
_cache: dict[str, str] = {}


def load_prompt(name: str) -> str:
    """读取并缓存模板；缺失时记录错误并返回空串（服务层有兜底提示）。"""
    if name in _cache:
        return _cache[name]
    path = _PROMPTS_DIR / name
    try:
        content = path.read_text(encoding="utf-8").strip()
    except OSError:
        logger.error("Prompt 模板缺失：%s", path)
        content = ""
    _cache[name] = content
    return content
