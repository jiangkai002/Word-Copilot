"""提案校验器与事件模型单元测试（格式 / 表格 / 公式 / 段落插入）。

只测纯函数与 Pydantic 模型 —— 不依赖 agent_framework
（懒 import 设计保证未安装依赖时这些函数仍可用、可测）。
"""
from __future__ import annotations

import pytest
from pydantic import ValidationError

from app.models.agent import (
    AgentParagraph,
    FormulaProposalEvent,
    FormatProposalEvent,
    HeadingProposalEvent,
    ParagraphProposalEvent,
    ProposalEvent,
    TableProposalEvent,
)
from app.models.agent import DocumentSnapshotPayload
from app.services.agent_service import (
    validate_formula_proposal,
    validate_format_proposal,
    validate_heading_proposal,
    validate_paragraph_proposal,
    validate_proposal,
    validate_table_proposal,
)
from tests.test_agent_service import _snapshot

# ---------------- validate_format_proposal ----------------


def test_format_ok_bold_only() -> None:
    assert validate_format_proposal(_snapshot(), "p1", 0, bold=True) is None


@pytest.mark.parametrize(
    "style",
    ["normal", "title", "subtitle", "heading1", "heading2", "heading9"],
)
def test_format_ok_paragraph_style(style: str) -> None:
    assert validate_format_proposal(_snapshot(), "p1", 0, paragraph_style=style) is None


@pytest.mark.parametrize("style", ["正文", "Heading1", "heading10", "", 1])
def test_format_rejects_invalid_paragraph_style(style: object) -> None:
    error = validate_format_proposal(_snapshot(), "p1", 0, paragraph_style=style)
    assert error is not None and "paragraph_style" in error


def test_format_ok_all_fields() -> None:
    assert (
        validate_format_proposal(
            _snapshot(),
            "p1",
            0,
            bold=True,
            italic=False,
            underline=True,
            strikethrough=False,
            font_name="微软雅黑",
            font_size=14,
            color="#1A2b3C",
            alignment="justify",
        )
        is None
    )


def test_format_unknown_paragraph() -> None:
    error = validate_format_proposal(_snapshot(), "p9", 0, bold=True)
    assert error is not None and "不存在" in error


def test_format_empty_paragraph_rejected() -> None:
    # p3 为空段落 —— 无法定位格式目标
    error = validate_format_proposal(_snapshot(), "p3", 0, bold=True)
    assert error is not None and "为空段落" in error


def test_format_no_fields_rejected() -> None:
    error = validate_format_proposal(_snapshot(), "p1", 0)
    assert error is not None and "至少提供" in error


def test_format_bool_type_enforced() -> None:
    # 框架 schema 漂移时可能传字符串 —— 必须按类型拒绝
    error = validate_format_proposal(_snapshot(), "p1", 0, bold="true")
    assert error is not None and "bold" in error and "布尔" in error


def test_format_font_size_bounds() -> None:
    snapshot = _snapshot()
    assert validate_format_proposal(snapshot, "p1", 0, font_size=0.5) is not None
    assert validate_format_proposal(snapshot, "p1", 0, font_size=101) is not None
    assert validate_format_proposal(snapshot, "p1", 0, font_size=1) is None
    assert validate_format_proposal(snapshot, "p1", 0, font_size=100) is None


def test_format_font_size_string_rejected() -> None:
    error = validate_format_proposal(_snapshot(), "p1", 0, font_size="14")
    assert error is not None and "font_size" in error and "数字" in error


def test_format_color_pattern() -> None:
    snapshot = _snapshot()
    assert validate_format_proposal(snapshot, "p1", 0, color="red") is not None
    assert validate_format_proposal(snapshot, "p1", 0, color="#12345") is not None
    assert validate_format_proposal(snapshot, "p1", 0, color="#GGGGGG") is not None
    assert validate_format_proposal(snapshot, "p1", 0, color="#abcdef") is None


def test_format_font_name_length() -> None:
    snapshot = _snapshot()
    assert validate_format_proposal(snapshot, "p1", 0, font_name="") is not None
    assert validate_format_proposal(snapshot, "p1", 0, font_name="字" * 65) is not None
    assert validate_format_proposal(snapshot, "p1", 0, font_name="字" * 64) is None


def test_format_alignment_enum() -> None:
    snapshot = _snapshot()
    assert validate_format_proposal(snapshot, "p1", 0, alignment="middle") is not None
    for value in ("left", "center", "right", "justify"):
        assert validate_format_proposal(snapshot, "p1", 0, alignment=value) is None


def test_format_budget() -> None:
    error = validate_format_proposal(_snapshot(), "p1", 20, bold=True, max_proposals=20)
    assert error is not None and "上限" in error


def test_format_underline_false_alone_valid() -> None:
    # false = 显式关闭（如取消加粗场景的 underline=false）
    assert validate_format_proposal(_snapshot(), "p1", 0, underline=False) is None


# ---------------- validate_table_proposal ----------------


def test_table_ok() -> None:
    assert (
        validate_table_proposal(
            _snapshot(),
            "p1",
            [["项目", "数值"], ["预算", "100万"], ["实际", "120万"]],
            0,
        )
        is None
    )


def test_table_unknown_anchor() -> None:
    error = validate_table_proposal(_snapshot(), "p9", [["a"]], 0)
    assert error is not None and "不存在" in error


def test_table_empty_anchor_paragraph_rejected() -> None:
    error = validate_table_proposal(_snapshot(), "p3", [["a"]], 0)
    assert error is not None and "为空段落" in error


def test_table_empty_values_rejected() -> None:
    snapshot = _snapshot()
    assert validate_table_proposal(snapshot, "p1", [], 0) is not None
    assert validate_table_proposal(snapshot, "p1", "not-a-list", 0) is not None


def test_table_row_limits() -> None:
    snapshot = _snapshot()
    values = [["x"]] * 21
    assert validate_table_proposal(snapshot, "p1", values, 0) is not None
    assert validate_table_proposal(snapshot, "p1", [["x"]] * 20, 0) is None


def test_table_column_limits() -> None:
    snapshot = _snapshot()
    wide = [["x"] * 9]
    assert validate_table_proposal(snapshot, "p1", wide, 0) is not None
    assert validate_table_proposal(snapshot, "p1", [["x"] * 8], 0) is None


def test_table_ragged_rejected() -> None:
    error = validate_table_proposal(_snapshot(), "p1", [["a", "b"], ["c"]], 0)
    assert error is not None and "一致" in error


def test_table_cell_length() -> None:
    snapshot = _snapshot()
    long_cell = "x" * 201
    assert validate_table_proposal(snapshot, "p1", [[long_cell]], 0) is not None
    assert validate_table_proposal(snapshot, "p1", [["x" * 200]], 0) is None


def test_table_non_string_cell_rejected() -> None:
    error = validate_table_proposal(_snapshot(), "p1", [["a", 1]], 0)
    assert error is not None and "字符串" in error


def test_table_budget() -> None:
    error = validate_table_proposal(_snapshot(), "p1", [["a"]], 20, max_proposals=20)
    assert error is not None and "上限" in error


# ---------------- anchor 省略（文档末尾 / 空文档） ----------------


def test_table_anchor_none_ok() -> None:
    # anchor 省略 = 插入到文档末尾（空文档场景的解）
    snapshot = _snapshot()
    assert validate_table_proposal(snapshot, None, [["a"]], 0) is None
    # 空文档（无段落）同样通过
    empty = DocumentSnapshotPayload()
    assert validate_table_proposal(empty, None, [["a"]], 0) is None


def test_table_anchor_none_budget_enforced() -> None:
    error = validate_table_proposal(_snapshot(), None, [["a"]], 20, max_proposals=20)
    assert error is not None and "上限" in error


def test_table_anchor_non_string_rejected() -> None:
    error = validate_table_proposal(_snapshot(), 123, [["a"]], 0)
    assert error is not None and "anchor_paragraph_id" in error and "字符串" in error


def test_formula_anchor_none_ok() -> None:
    snapshot = _snapshot()
    assert validate_formula_proposal(snapshot, None, "E=mc^2", 0) is None
    empty = DocumentSnapshotPayload()
    assert validate_formula_proposal(empty, None, "E=mc^2", 0) is None


def test_formula_anchor_none_budget_enforced() -> None:
    error = validate_formula_proposal(_snapshot(), None, "E=mc^2", 20, max_proposals=20)
    assert error is not None and "上限" in error


def test_insert_error_messages_mention_document_end_fallback() -> None:
    # 用户给错锚点时，错误信息引导「省略 anchor 插入文档末尾」
    error = validate_table_proposal(_snapshot(), "p9", [["a"]], 0)
    assert error is not None and "省略" in error and "文档末尾" in error


# ---------------- 表内段落不能作插入锚点（in_table） ----------------


def _table_snapshot() -> DocumentSnapshotPayload:
    """含表内段落的快照（Body.paragraphs 自 WordApi 1.3 起含表内段）。"""
    return DocumentSnapshotPayload(
        paragraphs=[
            AgentParagraph(id="p1", text="正文段落。"),
            AgentParagraph(id="cell1", text="95%", in_table=True),
        ]
    )


def test_table_in_table_anchor_rejected() -> None:
    # 锚点指向表格单元格内段落 —— 插进去会写进单元格，必须拒绝并引导文档末尾
    error = validate_table_proposal(_table_snapshot(), "cell1", [["a"]], 0)
    assert error is not None and "表格单元格内" in error and "文档末尾" in error


def test_formula_in_table_anchor_rejected() -> None:
    error = validate_formula_proposal(_table_snapshot(), "cell1", "E=mc^2", 0)
    assert error is not None and "表格单元格内" in error and "文档末尾" in error


def test_paragraph_in_table_anchor_rejected() -> None:
    error = validate_paragraph_proposal(_table_snapshot(), "cell1", "新段落", 0)
    assert error is not None and "表格单元格内" in error and "文档末尾" in error


def test_in_table_anchor_null_still_ok() -> None:
    # in_table 只拦截显式锚点；省略 anchor（文档末尾）不受影响
    snapshot = _table_snapshot()
    assert validate_table_proposal(snapshot, None, [["a"]], 0) is None
    assert validate_formula_proposal(snapshot, None, "E=mc^2", 0) is None
    assert validate_paragraph_proposal(snapshot, None, "新段落", 0) is None


def test_in_table_paragraph_still_editable() -> None:
    # propose_edit / propose_format 作用于表内段落是合法的（改的是该段自身，
    # 不是「插到段后」）—— 只拦插入锚点
    snapshot = _table_snapshot()
    assert validate_format_proposal(snapshot, "cell1", 0, bold=True) is None
    assert validate_proposal(snapshot, "cell1", "95% 置信度", 0) is None


# ---------------- validate_formula_proposal ----------------


def test_formula_ok() -> None:
    assert validate_formula_proposal(_snapshot(), "p1", "E=mc^2", 0) is None


def test_formula_unknown_anchor() -> None:
    error = validate_formula_proposal(_snapshot(), "p9", "E=mc^2", 0)
    assert error is not None and "不存在" in error


def test_formula_empty_anchor_paragraph_rejected() -> None:
    error = validate_formula_proposal(_snapshot(), "p3", "E=mc^2", 0)
    assert error is not None and "为空段落" in error


def test_formula_empty_latex_rejected() -> None:
    snapshot = _snapshot()
    assert validate_formula_proposal(snapshot, "p1", "", 0) is not None
    assert validate_formula_proposal(snapshot, "p1", "   ", 0) is not None


def test_formula_length_limit() -> None:
    snapshot = _snapshot()
    assert validate_formula_proposal(snapshot, "p1", "x" * 501, 0) is not None
    assert validate_formula_proposal(snapshot, "p1", "x" * 500, 0) is None


def test_formula_budget() -> None:
    error = validate_formula_proposal(_snapshot(), "p1", "E=mc^2", 20, max_proposals=20)
    assert error is not None and "上限" in error


# ---------------- validate_paragraph_proposal ----------------


def test_paragraph_ok() -> None:
    assert validate_paragraph_proposal(_snapshot(), "p1", "这是一段新文字。", 0) is None


def test_paragraph_anchor_none_ok_empty_document() -> None:
    # 用户反馈的核心场景：空文档 + 随便写点什么 → 必须可插入
    empty = DocumentSnapshotPayload()
    assert validate_paragraph_proposal(empty, None, "随便写的内容", 0) is None
    assert validate_paragraph_proposal(_snapshot(), None, "第一段\n第二段", 0) is None


def test_paragraph_multiline_within_limit() -> None:
    snapshot = _snapshot()
    text = "\n".join(f"第{i}行" for i in range(1, 21))  # 恰好 20 行
    assert validate_paragraph_proposal(snapshot, None, text, 0) is None
    over = "\n".join(f"第{i}行" for i in range(1, 22))  # 21 行
    error = validate_paragraph_proposal(snapshot, None, over, 0)
    assert error is not None and "段落" in error and "21" in error


def test_paragraph_blank_lines_ignored() -> None:
    # 空行只作分隔不计段落数；整串空白拒绝
    snapshot = _snapshot()
    assert validate_paragraph_proposal(snapshot, None, "第一段\n\n第二段", 0) is None
    assert validate_paragraph_proposal(snapshot, None, " \n \n ", 0) is not None


def test_paragraph_empty_text_rejected() -> None:
    snapshot = _snapshot()
    assert validate_paragraph_proposal(snapshot, None, "", 0) is not None
    assert validate_paragraph_proposal(snapshot, None, "   ", 0) is not None
    assert validate_paragraph_proposal(snapshot, None, 123, 0) is not None  # 非字符串


def test_paragraph_length_limit() -> None:
    snapshot = _snapshot()
    assert validate_paragraph_proposal(snapshot, None, "字" * 2001, 0) is not None
    assert validate_paragraph_proposal(snapshot, None, "字" * 2000, 0) is None


def test_paragraph_unknown_anchor_rejected() -> None:
    error = validate_paragraph_proposal(_snapshot(), "p9", "文字", 0)
    assert error is not None and "不存在" in error


def test_paragraph_empty_anchor_paragraph_rejected() -> None:
    error = validate_paragraph_proposal(_snapshot(), "p3", "文字", 0)
    assert error is not None and "为空段落" in error


def test_paragraph_budget() -> None:
    error = validate_paragraph_proposal(_snapshot(), None, "文字", 20, max_proposals=20)
    assert error is not None and "上限" in error


# ---------------- 事件模型（wire 格式钉死） ----------------


def test_format_event_serialization() -> None:
    event = FormatProposalEvent(paragraph_id="p1", original_text="原文", summary="加粗", bold=True)
    data = event.model_dump()
    assert data["kind"] == "format"
    assert data["paragraph_id"] == "p1"
    assert data["original_text"] == "原文"
    assert data["bold"] is True
    # 未提供的字段序列化为 None（前端按「不修改」处理）
    assert data["italic"] is None
    assert data["font_name"] is None
    assert data["alignment"] is None


def test_table_event_serialization() -> None:
    event = TableProposalEvent(
        anchor_paragraph_id="p2",
        anchor_text="锚点原文",
        summary="插入表格",
        values=[["a", "b"], ["c", "d"]],
    )
    data = event.model_dump()
    assert data["kind"] == "insert-table"
    assert data["header"] is True  # 默认表头
    assert data["values"] == [["a", "b"], ["c", "d"]]


def test_formula_event_serialization() -> None:
    event = FormulaProposalEvent(
        anchor_paragraph_id="p2",
        anchor_text="锚点原文",
        summary="插入公式",
        latex="\\frac{a}{b}",
        display=False,
    )
    data = event.model_dump()
    assert data["kind"] == "insert-formula"
    assert data["latex"] == "\\frac{a}{b}"
    assert data["display"] is False


def test_paragraph_event_serialization() -> None:
    event = ParagraphProposalEvent(
        anchor_paragraph_id=None,
        anchor_text="",
        paragraph_text="第一段\n第二段",
        summary="插入文字",
    )
    data = event.model_dump()
    assert data["kind"] == "insert-paragraph"
    # anchor 为 None（文档末尾）时序列化为 null —— 前端走 Body.insert*("End")
    assert data["anchor_paragraph_id"] is None
    assert data["paragraph_text"] == "第一段\n第二段"


def test_paragraph_event_length_rejected() -> None:
    with pytest.raises(ValidationError):
        ParagraphProposalEvent(paragraph_text="", summary="x")
    with pytest.raises(ValidationError):
        ParagraphProposalEvent(paragraph_text="字" * 2001, summary="x")


@pytest.mark.parametrize("level", [1, 2, 3, 9])
def test_heading_proposal_valid(level: int) -> None:
    assert validate_heading_proposal(_snapshot(), None, "系统概述", level, 0) is None
    assert validate_heading_proposal(_snapshot(), "p1", "系统概述", level, 0) is None


@pytest.mark.parametrize("level", [0, 10, 1.5, True, "1"])
def test_heading_proposal_rejects_invalid_level(level: object) -> None:
    error = validate_heading_proposal(_snapshot(), None, "系统概述", level, 0)
    assert error is not None and "1-9" in error


def test_heading_proposal_rejects_blank_multiline_and_table_anchor() -> None:
    assert "非空" in (validate_heading_proposal(_snapshot(), None, "   ", 1, 0) or "")
    assert "不能换行" in (validate_heading_proposal(_snapshot(), None, "第一章\n说明", 1, 0) or "")
    table_snapshot = DocumentSnapshotPayload(
        paragraphs=[AgentParagraph(id="cell", text="单元格", in_table=True)]
    )
    assert "表格单元格" in (
        validate_heading_proposal(table_snapshot, "cell", "表内标题", 2, 0) or ""
    )


def test_heading_event_serialization() -> None:
    event = HeadingProposalEvent(
        anchor_paragraph_id="p1",
        anchor_text="锚点原文",
        heading_text="第一章 系统概述",
        level=1,
        summary="插入章标题",
    )
    data = event.model_dump()
    assert data["kind"] == "insert-heading"
    assert data["heading_text"] == "第一章 系统概述"
    assert data["level"] == 1


def test_text_event_has_no_kind_field() -> None:
    # 既有文本提案 wire 格式保持不变（proposal 帧不加 kind 字段；
    # 前端 AgentApi 解析时按帧名补 kind:"text"）
    event = ProposalEvent(
        paragraph_id="p1",
        original_text="原文",
        new_text="新文",
        summary="修改",
    )
    assert "kind" not in event.model_dump()


def test_format_event_invalid_values_rejected() -> None:
    # Pydantic 层兜底：非法颜色 / 字号直接构造失败（工具层 try/except 回喂模型）
    with pytest.raises(ValidationError):
        FormatProposalEvent(paragraph_id="p1", summary="x", color="red")
    with pytest.raises(ValidationError):
        FormatProposalEvent(paragraph_id="p1", summary="x", font_size=200)
    with pytest.raises(ValidationError):
        FormulaProposalEvent(anchor_paragraph_id="p1", summary="x", latex="x" * 501)
