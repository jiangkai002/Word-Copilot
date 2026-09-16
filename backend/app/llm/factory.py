"""Provider 工厂（§35）：按 .env 配置创建 Provider 实例（进程内单例）。"""
from __future__ import annotations

import logging

from .. import config
from .base import LLMProvider
from .openai_provider import OpenAICompatibleProvider

logger = logging.getLogger(__name__)

_provider: LLMProvider | None = None


def get_provider() -> LLMProvider:
    """获取当前配置的 Provider（首次调用时创建并缓存）。

    第一版所有 provider 名义类型统一落到 OpenAI Compatible 协议
    （OpenAI / DeepSeek / Qwen / 智谱 / 企业自建网关均兼容）。
    """
    global _provider
    if _provider is None:
        logger.info(
            "初始化 LLM Provider：%s @ %s model=%s",
            config.LLM_PROVIDER,
            config.LLM_BASE_URL,
            config.LLM_MODEL or "（未配置）",
        )
        if not config.LLM_API_KEY:
            logger.warning("LLM_API_KEY 未配置 —— 模型调用将返回 LLM_ERROR")
        _provider = OpenAICompatibleProvider(
            base_url=config.LLM_BASE_URL,
            api_key=config.LLM_API_KEY,
            model=config.LLM_MODEL,
            timeout=config.LLM_TIMEOUT_SECONDS,
            default_temperature=config.LLM_TEMPERATURE,
            default_max_tokens=config.LLM_MAX_TOKENS,
        )
    return _provider
