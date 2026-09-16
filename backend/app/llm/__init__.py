"""LLM Provider 抽象层（§34 / §35）。

模型 SDK 不与业务耦合：Provider 只负责「消息进 → 文本出」，
Prompt 构造、JSON 校验、错误码映射都在 services 层完成（§75）。
"""
