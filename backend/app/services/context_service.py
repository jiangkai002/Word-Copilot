"""把前端传来的 DocumentContext 渲染为 Prompt 中的上下文文本块（§13 / §60）。

隐私原则（§60）在前端已保证：Selection / Paragraph 模式不携带全文，
后端在这里只做「有什么渲染什么」，绝不请求额外内容。
"""
from __future__ import annotations

from ..models.chat import DocumentContextPayload

_TRUNCATED_NOTE = "（文档过长，以下为与当前问题最相关的部分）"


def _paragraph_block(label: str, paragraph) -> str:
    if paragraph is None or not (paragraph.text or "").strip():
        return ""
    style_note = f"（段落样式：{paragraph.style}）" if paragraph.style else ""
    return f"【{label}】{style_note}\n{paragraph.text.rstrip()}"


def _outline_block(outline) -> str:
    if not outline:
        return ""
    lines = ["【文档结构】"]
    for item in outline:
        indent = "  " * max(0, item.level - 1)
        lines.append(f"{indent}{'#' * min(item.level, 6)} {item.title}")
    return "\n".join(lines)


def render_context(context: DocumentContextPayload | None) -> str:
    """渲染上下文块；无上下文返回空串（纯对话）。"""
    if context is None:
        return ""

    blocks: list[str] = []

    if context.context_type == "selection":
        if context.selection and context.selection.text.strip():
            blocks.append(f"【当前选区】\n{context.selection.text.rstrip()}")
        if context.current_paragraph and context.current_paragraph.text.strip():
            blocks.append(_paragraph_block("选区所在段落", context.current_paragraph))

    elif context.context_type == "paragraph":
        blocks.append(_paragraph_block("上一段", context.previous_paragraph))
        blocks.append(_paragraph_block("光标所在段落", context.current_paragraph))
        blocks.append(_paragraph_block("下一段", context.next_paragraph))
        blocks = [b for b in blocks if b]

    elif context.context_type == "section":
        blocks.append(_outline_block(context.outline))
        if context.section:
            blocks.append(
                f"【当前章节：{context.section.title}】\n{context.section.content.rstrip()}"
            )

    elif context.context_type == "document":
        if context.truncated:
            blocks.append(_TRUNCATED_NOTE)
        if context.outline:
            blocks.append(_outline_block(context.outline))
        if context.section and context.section.content.strip():
            blocks.append(
                f"【当前章节：{context.section.title}】\n{context.section.content.rstrip()}"
            )
        if context.document_text and context.document_text.strip():
            label = "【文档全文】" if not context.truncated else "【文档内容（节选）】"
            blocks.append(f"{label}\n{context.document_text.rstrip()}")

    return "\n\n".join(b for b in blocks if b.strip())
