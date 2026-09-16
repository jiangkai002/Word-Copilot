"""Edit 编排（§36 / §37 / §17）：

1. 组装 Edit Prompt（§36 规则 + 目标文本 + 前后文 + 重新生成约束）
2. 调用 LLM（非流式，较短超时）
3. 解析 JSON → Pydantic 校验（EditCandidate）
4. 强制核对 original_text 与捕获原文一致（逐字符，忽略行尾差异）
5. 失败时给模型一次纠错机会；仍失败 → INVALID_EDIT_RESPONSE

返回的 EditResponse.original_text 恒为前端捕获的 target.text（权威原文），
保证客户端 Diff 坐标系与文档实际内容一致（§19）。
"""
from __future__ import annotations

import json
import logging
import uuid

from pydantic import ValidationError

from .. import config
from ..llm.base import CompletionResult, LLMProvider
from ..llm.factory import get_provider
from ..models.edit import EditCandidate, EditRequest, EditResponse
from . import prompt_service

logger = logging.getLogger(__name__)

_FALLBACK_SYSTEM_PROMPT = (
    "你是 Microsoft Word 文档编辑助手。根据用户修改要求生成修改后的目标文本。"
    "只输出 JSON。"
)

_CORRECTION_PROMPT = (
    "你的上一次输出不符合要求：JSON 结构不完整，或 original_text 与【目标文本】不一致。"
    "请严格按要求的 JSON 格式重新输出，"
    'original_text 必须逐字符复制【目标文本】，new_text 为修改后的完整文本。'
)

_MAX_ATTEMPTS = 2  # 首次 + 一次纠错


class EditServiceError(Exception):
    """带错误码的 Edit 业务错误（由 API 层映射为 §45 统一格式）。"""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


def _normalize_line_endings(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def _extract_json(content: str) -> dict | None:
    """从模型输出中提取 JSON 对象；容忍 ```json 代码块与前后说明文字。"""
    text = content.strip()
    if not text:
        return None
    # 去掉 Markdown 代码块围栏
    if text.startswith("```"):
        text = text.strip("`")
        first_newline = text.find("\n")
        if first_newline != -1:
            text = text[first_newline + 1 :]
    # 直接解析
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass
    # 兜底：截取首个 { 到最后一个 } 的片段
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end > start:
        try:
            parsed = json.loads(text[start : end + 1])
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            return None
    return None


def _build_edit_messages(request: EditRequest) -> list[dict[str, str]]:
    system_prompt = prompt_service.load_prompt("edit.md") or _FALLBACK_SYSTEM_PROMPT

    lines: list[str] = [f"【用户修改要求】\n{request.instruction}", ""]
    lines.append("【目标文本】（original_text 必须与此逐字符一致）")
    lines.append(request.target.text)

    if request.context:
        context_parts: list[str] = []
        if request.context.previous:
            context_parts.append(f"【目标之前的文字】\n{request.context.previous}")
        if request.context.next:
            context_parts.append(f"【目标之后的文字】\n{request.context.next}")
        if context_parts:
            lines.append("\n".join(context_parts))

    if request.regenerate_count > 0 or request.avoid_texts:
        lines.append(
            "\n【重新生成】用户对之前的版本不满意，请给出明显不同的新版本，不要与以下历史版本雷同："
        )
        for index, text in enumerate(request.avoid_texts, start=1):
            preview = text if len(text) <= 120 else f"{text[:120]}…"
            lines.append(f"历史版本 {index}：{preview}")

    lines.append(
        "\n【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块、不要多余文字）：\n"
        '{"summary": "一句话概括本次修改", '
        '"original_text": "逐字符复制【目标文本】", '
        '"new_text": "修改后的完整文本", '
        '"warnings": []}'
    )

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "\n".join(lines)},
    ]


def _validate_candidate(
    candidate: EditCandidate, request: EditRequest
) -> list[str]:
    """返回额外 warnings；original_text 与捕获原文不一致时抛 EditServiceError。"""
    if _normalize_line_endings(candidate.original_text) != _normalize_line_endings(
        request.target.text
    ):
        raise EditServiceError(
            "INVALID_EDIT_RESPONSE",
            "模型回显的 original_text 与捕获的原文不一致",
        )
    warnings = list(candidate.warnings)
    if _normalize_line_endings(candidate.new_text) == _normalize_line_endings(
        request.target.text
    ):
        warnings.append("模型认为该内容无需修改")
    if not candidate.new_text.strip():
        warnings.append("修改结果为空（应用后将删除目标内容），请确认是否符合预期")
    return warnings


async def generate_edit(request: EditRequest) -> EditResponse:
    provider: LLMProvider = get_provider()
    messages = _build_edit_messages(request)

    candidate: EditCandidate | None = None
    extra_warnings: list[str] = []
    last_result: CompletionResult | None = None

    for attempt in range(1, _MAX_ATTEMPTS + 1):
        result = await provider.chat(
            messages,
            timeout=config.LLM_EDIT_TIMEOUT_SECONDS,
            temperature=0.2 if request.regenerate_count == 0 else 0.7,
        )
        last_result = result
        parsed = _extract_json(result.content)
        if parsed is not None:
            try:
                candidate = EditCandidate.model_validate(parsed)
            except ValidationError as exc:
                logger.info(
                    "Edit 尝试 %d：JSON 字段校验失败（%s）",
                    attempt,
                    str(exc)[:200],
                )
                candidate = None
        if candidate is not None:
            try:
                extra_warnings = _validate_candidate(candidate, request)
                break
            except EditServiceError:
                logger.info("Edit 尝试 %d：original_text 回显不一致", attempt)

        # 未通过 → 追加纠错轮
        if attempt < _MAX_ATTEMPTS:
            messages = messages + [
                {"role": "assistant", "content": result.content},
                {"role": "user", "content": _CORRECTION_PROMPT},
            ]

    if candidate is None:
        raise EditServiceError(
            "INVALID_EDIT_RESPONSE",
            "模型未返回要求的 JSON 结构，已重试仍失败",
        )

    assert last_result is not None
    if last_result.usage.total_tokens > 0:
        logger.info(
            "edit conversation=%s model=%s tokens=%s attempt=%d",
            request.conversation_id,
            provider.model_name,
            last_result.usage.total_tokens,
            attempt,
        )

    return EditResponse(
        edit_id=uuid.uuid4().hex,
        original_text=request.target.text,  # 权威原文：来自客户端捕获
        new_text=candidate.new_text,
        summary=(candidate.summary or "").strip() or "已完成修改",
        warnings=[w for w in extra_warnings if w],
    )
