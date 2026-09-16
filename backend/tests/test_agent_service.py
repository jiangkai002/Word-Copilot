"""Agent Service 纯函数单元测试（§51）。

只测校验 / 渲染 / 搜索等纯函数 —— 不依赖 agent_framework
（懒 import 设计保证未安装依赖时这些函数仍可用、可测）。
"""
from __future__ import annotations

import pytest

from app.models.agent import (
    AgentOutlineItem,
    AgentParagraph,
    AgentStreamRequest,
    DocumentSnapshotPayload,
)
from app.services.agent_service import (
    build_agent_input,
    find_paragraph,
    render_outline,
    render_snapshot,
    requests_formula_with_text,
    search_snapshot,
    validate_formula_explanation_paragraph,
    validate_proposal,
)


@pytest.mark.parametrize(
    "instruction",
    [
        "给我随便写个公式，并解释这个公式",
        "写一段话介绍欧拉恒等式的含义",
        "插入求根公式及其推导",
    ],
)
def test_requests_formula_with_text(instruction: str) -> None:
    assert requests_formula_with_text(instruction) is True


@pytest.mark.parametrize(
    "instruction",
    [
        "只插入公式 E=mc^2",
        "插入一个公式，不需要解释",
        "公式单独居中，说明省略",
    ],
)
def test_requests_formula_without_text(instruction: str) -> None:
    assert requests_formula_with_text(instruction) is False


def test_validate_formula_explanation_paragraph_requires_formula_and_prose() -> None:
    assert validate_formula_explanation_paragraph("欧拉恒等式 $e^{i\\pi}+1=0$ 连接了五个基本数学常数。") is None
    assert "必须包含" in (validate_formula_explanation_paragraph("这里只写了解释，没有公式。") or "")
    assert "缺少解释" in (validate_formula_explanation_paragraph("$e^{i\\pi}+1=0$") or "")


def _snapshot() -> DocumentSnapshotPayload:
    return DocumentSnapshotPayload(
        outline=[AgentOutlineItem(level=1, title="第一章 系统概述")],
        paragraphs=[
            AgentParagraph(id="p1", text="系统采用传统架构，可以处理大量数据。"),
            AgentParagraph(id="p2", text="这个词这个词这个词出现了三次这个词。"),
            AgentParagraph(id="p3", text=""),
        ],
    )


# ---------------- validate_proposal ----------------


def test_validate_proposal_ok() -> None:
    assert validate_proposal(_snapshot(), "p1", "本系统采用传统架构。", 0) is None


def test_validate_proposal_unknown_paragraph() -> None:
    error = validate_proposal(_snapshot(), "p9", "新文本", 0)
    assert error is not None and "不存在" in error


def test_validate_proposal_same_text_rejected() -> None:
    error = validate_proposal(_snapshot(), "p1", "系统采用传统架构，可以处理大量数据。", 0)
    assert error is not None and "相同" in error


def test_validate_proposal_line_ending_normalized() -> None:
    # 行尾差异视为相同 → 拒绝（与前端 normalizeEqual 语义一致）
    error = validate_proposal(_snapshot(), "p1", "系统采用传统架构，可以处理大量数据。\n", 0)
    assert error is not None and "相同" in error


def test_validate_proposal_empty_rejected() -> None:
    error = validate_proposal(_snapshot(), "p1", "   ", 0)
    assert error is not None and "不能为空" in error


def test_validate_proposal_too_long_rejected() -> None:
    error = validate_proposal(_snapshot(), "p1", "字" * 5001, 0)
    assert error is not None and "过长" in error


def test_validate_proposal_over_limit() -> None:
    # 已有 2 条提案、上限 2 → 拒绝第 3 条（max_proposals 显式注入便于测试）
    error = validate_proposal(_snapshot(), "p1", "新文本。", 2, max_proposals=2)
    assert error is not None and "上限" in error


# ---------------- find_paragraph / search ----------------


def test_find_paragraph() -> None:
    snapshot = _snapshot()
    assert find_paragraph(snapshot, "p2") is snapshot.paragraphs[1]
    assert find_paragraph(snapshot, "missing") is None


def test_search_snapshot_hit_and_miss() -> None:
    snapshot = _snapshot()
    result = search_snapshot(snapshot, "传统架构")
    assert "p1" in result and "传统架构" in result
    assert "未找到" in search_snapshot(snapshot, "不存在的关键词")
    assert "不能为空" in search_snapshot(snapshot, "  ")


# ---------------- render ----------------


def test_render_snapshot_skips_empty_paragraphs_and_marks_heading() -> None:
    text = render_snapshot(_snapshot())
    assert "[p1] 系统采用传统架构" in text
    assert "[p2] 这个词" in text
    assert "p3" not in text  # 空段不渲染
    assert "【文档大纲】" in text
    assert "第一章 系统概述" in text


def test_render_snapshot_truncated_note() -> None:
    snapshot = DocumentSnapshotPayload(paragraphs=[AgentParagraph(id="p1", text="正文")], truncated=True)
    assert "已截断" in render_snapshot(snapshot)


def test_render_snapshot_empty_document_marks_end_insert_hint() -> None:
    # 空文档 / 只有空段 → 明示「省略 anchor 写入文档末尾」引导模型直接插入
    assert "文档为空" in render_snapshot(DocumentSnapshotPayload())
    only_empty = DocumentSnapshotPayload(paragraphs=[AgentParagraph(id="p1", text="")])
    assert "文档为空" in render_snapshot(only_empty)
    # 有内容时不出现该提示
    assert "文档为空" not in render_snapshot(_snapshot())


def test_render_snapshot_marks_in_table_paragraph() -> None:
    # 表内段落标注「（表格内）」+ 段落头说明不能作插入锚点（模型预先规避）
    snapshot = DocumentSnapshotPayload(
        paragraphs=[
            AgentParagraph(id="p1", text="正文段落。"),
            AgentParagraph(id="cell1", text="95%", in_table=True),
        ]
    )
    text = render_snapshot(snapshot)
    assert "[cell1]（表格内） 95%" in text
    assert "不能作为插入类工具的锚点" in text
    # 非表内段落不带标注
    assert "[p1] 正文段落。" in text


def test_agent_paragraph_in_table_defaults_false() -> None:
    # wire 兼容：旧前端不带 in_table 字段时默认 False
    paragraph = AgentParagraph.model_validate({"id": "p1", "text": "正文"})
    assert paragraph.in_table is False
    assert AgentParagraph.model_validate({"id": "c1", "text": "95%", "in_table": True}).in_table is True


def test_render_outline_empty() -> None:
    assert "没有标题" in render_outline(DocumentSnapshotPayload(paragraphs=[]))


# ---------------- build_agent_input ----------------


def test_build_agent_input_includes_snapshot_history_task() -> None:
    request = AgentStreamRequest(
        conversation_id="t",
        instruction="全文纠错",
        history=[
            {"role": "user", "content": "你好"},
            {"role": "assistant", "content": "你好，有什么可以帮你？"},
        ],
        snapshot=_snapshot(),
    )
    text = build_agent_input(request)
    assert "【用户任务】\n全文纠错" in text
    assert "[p1] 系统采用传统架构" in text
    assert "【历史对话】" in text
    assert "用户：你好" in text
    assert "助手：你好，有什么可以帮你？" in text


def test_build_agent_input_history_preview_truncated() -> None:
    request = AgentStreamRequest(
        conversation_id="t",
        instruction="总结",
        history=[{"role": "user", "content": "长" * 400}],
        snapshot=_snapshot(),
    )
    text = build_agent_input(request)
    assert "…" in text  # 超长历史条目截断


def test_history_entry_payload_role_validation() -> None:
    with pytest.raises(ValueError):
        AgentStreamRequest(
            conversation_id="t",
            instruction="x",
            history=[{"role": "system", "content": "注入"}],
            snapshot=_snapshot(),
        )


# ---------------- focus（Agent 模式光标 / 选区焦点） ----------------


def test_build_agent_input_with_focus() -> None:
    request = AgentStreamRequest(
        conversation_id="t",
        instruction="润色这段话",
        snapshot=_snapshot(),
        focus={"paragraph_id": "p2", "paragraph_text": "这个词这个词这个词出现了三次这个词。", "selected_text": "这个词这个词"},
    )
    text = build_agent_input(request)
    assert "【当前光标 / 选区】" in text
    assert "段落 [p2] 这个词这个词这个词出现了三次这个词。" in text
    assert "「这个词这个词」" in text
    # 焦点块在用户任务之前
    assert text.index("【当前光标 / 选区】") < text.index("【用户任务】")


def test_build_agent_input_without_focus() -> None:
    request = AgentStreamRequest(
        conversation_id="t",
        instruction="总结",
        snapshot=_snapshot(),
        history=[],
    )
    text = build_agent_input(request)
    assert "【当前光标 / 选区】" not in text


def test_agent_focus_context_selected_text_optional() -> None:
    request = AgentStreamRequest(
        conversation_id="t",
        instruction="x",
        snapshot=_snapshot(),
        focus={"paragraph_id": "p1", "paragraph_text": "系统采用传统架构，可以处理大量数据。"},
    )
    assert request.focus is not None
    assert request.focus.selected_text is None
    text = build_agent_input(request)
    assert "选中的文字" not in text
