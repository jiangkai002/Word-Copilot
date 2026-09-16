"""Agent 编排（§51：Agent → Tool Call → Result → Agent 循环）。

Microsoft Agent Framework 单 Agent，快照以闭包注入工具：
- 只读工具（get_outline / read_paragraph / search_paragraphs）：只检索请求内
  的文档快照 —— 后端碰不到 Word（§75），LLM 永不直接操作文档（§14）。
- propose_edit：结构化提案通道 —— 批量修改以工具参数提交（Pydantic 强校验），
  不依赖模型的结构化输出能力（DeepSeek 的 response_format 行为未验证）。

安全设计：
- 提案的 original_text 由后端从快照回显（模型只提供 paragraph_id + new_text），
  前端快照文本为权威原文（§19 Diff 坐标系）。
- 每条提案独立校验（段落存在 / 有变化 / 长度 / 总数上限），违规信息回喂模型。
- agent_framework 懒 import：未安装依赖时后端其余功能不受影响
  （AGENT_NOT_INSTALLED 错误帧）。

运行结构：agent 运行在独立 worker task 中（框架迭代不跨任务恢复，
保证内部取消作用域安全），事件经 asyncio.Queue 传给 SSE 生成器；
超时用 asyncio.timeout 包裹整个运行（AGENT_TIMEOUT_SECONDS）。
"""
from __future__ import annotations

import asyncio
import contextlib
import logging
import re
import time
from collections.abc import AsyncIterator
from dataclasses import dataclass

from pydantic import ValidationError

from .. import config
from ..models.agent import (
    AgentParagraph,
    AgentStreamRequest,
    DocumentSnapshotPayload,
    FormulaProposalEvent,
    FormatProposalEvent,
    ParagraphProposalEvent,
    ProposalEvent,
    TableProposalEvent,
)
from . import prompt_service, skill_service

logger = logging.getLogger(__name__)

_FALLBACK_SYSTEM_PROMPT = (
    "你是 Microsoft Word 文档 AI 助手。基于文档快照回答问题，或通过 propose_edit "
    "工具提交批量修改提案；其余时间用中文简洁回答。"
)

_HISTORY_LIMIT = 10
_HISTORY_ENTRY_PREVIEW = 300
_MAX_NEW_TEXT_LENGTH = 5000
_SEARCH_PREVIEW_LIMIT = 30

# 格式 / 插入提案校验常量（与前端 EditPlan / 引擎约定一致）
_MAX_FONT_NAME_LENGTH = 64
_FONT_SIZE_MIN = 1.0
_FONT_SIZE_MAX = 100.0
_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")
_ALIGNMENTS = ("left", "center", "right", "justify")
_TABLE_MAX_ROWS = 20
_TABLE_MAX_COLUMNS = 8
_TABLE_CELL_MAX_LENGTH = 200
_MAX_LATEX_LENGTH = 500
_MAX_PARAGRAPH_INSERT_LENGTH = 2000
_MAX_PARAGRAPH_INSERT_LINES = 20

_FORMULA_TEXT_REQUEST_RE = re.compile(r"解释|说明|含义|意义|推导|介绍|一段话|一段文字|正文")
_FORMULA_TEXT_NEGATION_RE = re.compile(
    r"(?:不需要|不用|不要|无需|省略).{0,6}(?:解释|说明|文字)|"
    r"(?:解释|说明|文字).{0,6}(?:不需要|不用|不要|无需|省略)"
)
_INLINE_FORMULA_MARKER_RE = re.compile(r"\$[^$\r\n]+\$|\\\(.+?\\\)")


def requests_formula_with_text(instruction: str) -> bool:
    """用户是否要求把公式和解释/正文一起写入（此时必须走富段落）。"""
    if _FORMULA_TEXT_NEGATION_RE.search(instruction):
        return False
    return _FORMULA_TEXT_REQUEST_RE.search(instruction) is not None


def validate_formula_explanation_paragraph(paragraph_text: str) -> str | None:
    """确保“公式 + 解释”提案确实同时包含可转换公式和实质说明。"""
    if _INLINE_FORMULA_MARKER_RE.search(paragraph_text) is None:
        return "paragraph_text 必须包含用 $LaTeX$ 标注的公式"
    prose = _INLINE_FORMULA_MARKER_RE.sub("", paragraph_text)
    meaningful = re.sub(r"[^\w\u4e00-\u9fff]+", "", prose, flags=re.UNICODE)
    if len(meaningful) < 8:
        return "paragraph_text 只有公式、缺少解释正文；请把公式含义和说明一并写入"
    return None


class AgentServiceError(Exception):
    """带错误码的 Agent 业务错误（由 API 层映射为 SSE error 帧，§45）。"""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass
class TokenEvent:
    """自然语言增量（SSE token 帧载荷）。"""

    text: str


@dataclass
class DoneEvent:
    """运行正常结束（SSE done 帧载荷；usage 为 best-effort，可为空）。"""

    usage: dict[str, int]


@dataclass
class _ErrorSentinel:
    """worker 内部错误传递（转换为 AgentServiceError 重新抛出）。"""

    code: str
    message: str


AgentEvent = (
    TokenEvent
    | ProposalEvent
    | FormatProposalEvent
    | TableProposalEvent
    | FormulaProposalEvent
    | ParagraphProposalEvent
    | DoneEvent
)


def _normalize_line_endings(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def _clean_proposal_text(text: str) -> str:
    """提案文本规范化：行尾统一 + 剥离首尾换行。

    段落标记不属于段落文本（快照已 stripTrailingMarks）；提案 new_text
    尾部换行会经 toWordText 变成插入段落标记，必须在服务端剥离。
    """
    return _normalize_line_endings(text).strip("\n")


# ------------------------------------------------------------------
# 纯函数（不依赖 agent_framework，供单元测试）
# ------------------------------------------------------------------


def find_paragraph(snapshot: DocumentSnapshotPayload, paragraph_id: str) -> AgentParagraph | None:
    """按 id 查找快照段落（段落量为数百级，线性查找可接受）。"""
    for paragraph in snapshot.paragraphs:
        if paragraph.id == paragraph_id:
            return paragraph
    return None


def validate_proposal(
    snapshot: DocumentSnapshotPayload,
    paragraph_id: str,
    new_text: str,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> str | None:
    """校验一条提案；返回错误说明（None = 通过）。纯函数。

    - paragraph_id 必须存在于快照
    - new_text 非空、长度受限、（行尾归一后）与原文不同
    - 提案总数不超过上限
    """
    limit = config.AGENT_MAX_PROPOSALS if max_proposals is None else max_proposals
    paragraph = find_paragraph(snapshot, paragraph_id)
    if paragraph is None:
        return f"段落 {paragraph_id} 不存在，请使用文档快照中 [id] 标注的段落 id"
    normalized_new = _clean_proposal_text(new_text)
    if not normalized_new.strip():
        return "new_text 不能为空"
    if len(new_text) > _MAX_NEW_TEXT_LENGTH:
        return f"new_text 过长（超过 {_MAX_NEW_TEXT_LENGTH} 字符），请缩短"
    if normalized_new == _clean_proposal_text(paragraph.text):
        return "new_text 与原文相同；无需修改的段落不要提交提案"
    if proposal_count >= limit:
        return f"提案数已达上限（{limit} 条），请停止提交并在总结中说明"
    return None


def _validate_anchor(
    snapshot: DocumentSnapshotPayload,
    anchor_paragraph_id: str,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> AgentParagraph | str:
    """格式 / 插入提案共用的锚点校验：存在、非空段、预算内。

    通过返回段落对象（可回显原文），失败返回错误说明。
    """
    limit = config.AGENT_MAX_PROPOSALS if max_proposals is None else max_proposals
    paragraph = find_paragraph(snapshot, anchor_paragraph_id)
    if paragraph is None:
        return f"段落 {anchor_paragraph_id} 不存在，请使用文档快照中 [id] 标注的段落 id"
    if not paragraph.text.strip():
        return f"段落 {anchor_paragraph_id} 为空段落，无法作为目标 / 锚点；请选择有文字的段落"
    if proposal_count >= limit:
        return f"提案数已达上限（{limit} 条），请停止提交并在总结中说明"
    return paragraph


def _validate_insert_anchor(
    snapshot: DocumentSnapshotPayload,
    anchor_paragraph_id: object,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> tuple[AgentParagraph | None, str | None]:
    """插入类（表格 / 公式 / 段落）共用的锚点校验。

    anchor 为 None / 缺省 = 插入到文档末尾（空文档也可插入，无需锚点）；
    显式锚点须为字符串且对应快照中存在、非空、且不在表格单元格内的段落
    （Body.paragraphs 自 WordApi 1.3 起含表内段落，插到表内段后 = 插进单元格）。

    返回 (锚点段落或 None（文档末尾）, 错误说明或 None)。
    """
    limit = config.AGENT_MAX_PROPOSALS if max_proposals is None else max_proposals
    if anchor_paragraph_id is None:
        if proposal_count >= limit:
            return None, f"提案数已达上限（{limit} 条），请停止提交并在总结中说明"
        return None, None  # 文档末尾
    if not isinstance(anchor_paragraph_id, str):
        return None, f"anchor_paragraph_id 必须是字符串段落 id 或省略（当前类型：{type(anchor_paragraph_id).__name__}）"
    paragraph = find_paragraph(snapshot, anchor_paragraph_id)
    if paragraph is None:
        return None, f"段落 {anchor_paragraph_id} 不存在，请使用文档快照中 [id] 标注的段落 id，或省略 anchor 插入到文档末尾"
    if not paragraph.text.strip():
        return None, f"段落 {anchor_paragraph_id} 为空段落，无法作为锚点；请选择有文字的段落，或省略 anchor 插入到文档末尾"
    if paragraph.in_table:
        return (
            None,
            f"段落 {anchor_paragraph_id} 位于表格单元格内（快照中标注「（表格内）」），不能作为插入锚点——"
            "插入内容会写进单元格。请省略 anchor_paragraph_id 插入到文档末尾，或选择正文段落作锚点",
        )
    if proposal_count >= limit:
        return None, f"提案数已达上限（{limit} 条），请停止提交并在总结中说明"
    return paragraph, None


def validate_format_proposal(
    snapshot: DocumentSnapshotPayload,
    paragraph_id: str,
    proposal_count: int,
    *,
    bold: object = None,
    italic: object = None,
    underline: object = None,
    strikethrough: object = None,
    font_name: object = None,
    font_size: object = None,
    color: object = None,
    alignment: object = None,
    max_proposals: int | None = None,
) -> str | None:
    """校验一条格式提案；返回错误说明（None = 通过）。纯函数。

    入参为 object：框架可能把布尔 / 数字以字符串传入（schema 漂移），
    必须按类型显式拒绝而非依赖 Pydantic 自动转换。
    """
    anchor = _validate_anchor(snapshot, paragraph_id, proposal_count, max_proposals=max_proposals)
    if isinstance(anchor, str):
        return anchor

    fields = (bold, italic, underline, strikethrough, font_name, font_size, color, alignment)
    if all(value is None for value in fields):
        return "至少提供一个格式字段（bold / italic / underline / strikethrough / font_name / font_size / color / alignment）"

    for name, value in (("bold", bold), ("italic", italic), ("underline", underline), ("strikethrough", strikethrough)):
        if value is not None and not isinstance(value, bool):
            return f"{name} 必须是布尔值 true / false（当前值：{value!r}）"

    if font_name is not None:
        if not isinstance(font_name, str) or not font_name.strip() or len(font_name) > _MAX_FONT_NAME_LENGTH:
            return f"font_name 必须是 1-{_MAX_FONT_NAME_LENGTH} 字符的字体名称"

    if font_size is not None:
        if isinstance(font_size, bool) or not isinstance(font_size, (int, float)):
            return f"font_size 必须是数字（1-{_FONT_SIZE_MAX:.0f} 磅，当前值：{font_size!r}）"
        if not _FONT_SIZE_MIN <= float(font_size) <= _FONT_SIZE_MAX:
            return f"font_size 必须在 {_FONT_SIZE_MIN:.0f}-{_FONT_SIZE_MAX:.0f} 磅之间"

    if color is not None:
        if not isinstance(color, str) or not _COLOR_RE.match(color):
            return "color 必须是 #RRGGBB 格式的十六进制颜色，如 #FF0000"

    if alignment is not None:
        if not isinstance(alignment, str) or alignment not in _ALIGNMENTS:
            return f"alignment 必须是 {' / '.join(_ALIGNMENTS)} 之一"
    return None


def validate_table_proposal(
    snapshot: DocumentSnapshotPayload,
    anchor_paragraph_id: object,
    values: object,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> str | None:
    """校验一条表格插入提案；返回错误说明（None = 通过）。纯函数。

    anchor_paragraph_id 为 None / 缺省 = 插入到文档末尾（空文档可用）。
    """
    _, anchor_error = _validate_insert_anchor(
        snapshot, anchor_paragraph_id, proposal_count, max_proposals=max_proposals
    )
    if anchor_error:
        return anchor_error

    if not isinstance(values, list) or not values:
        return "values 必须是非空的字符串二维数组（list[list[str]]）"
    if len(values) > _TABLE_MAX_ROWS:
        return f"表格最多 {_TABLE_MAX_ROWS} 行（当前 {len(values)} 行）"

    column_count: int | None = None
    for i, row in enumerate(values):
        if not isinstance(row, list) or not row:
            return f"values 第 {i} 行必须是非空数组"
        if len(row) > _TABLE_MAX_COLUMNS:
            return f"表格最多 {_TABLE_MAX_COLUMNS} 列（第 {i} 行有 {len(row)} 列）"
        if column_count is None:
            column_count = len(row)
        elif len(row) != column_count:
            return f"表格每行的列数必须一致（矩形矩阵；第 {i} 行 {len(row)} 列，首行 {column_count} 列）"
        for j, cell in enumerate(row):
            if not isinstance(cell, str):
                return f"values[{i}][{j}] 必须是字符串（当前类型：{type(cell).__name__}）"
            if len(cell) > _TABLE_CELL_MAX_LENGTH:
                return f"单元格内容不能超过 {_TABLE_CELL_MAX_LENGTH} 字符（第 {i} 行第 {j} 列）"
    return None


def validate_formula_proposal(
    snapshot: DocumentSnapshotPayload,
    anchor_paragraph_id: object,
    latex: object,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> str | None:
    """校验一条公式插入提案；返回错误说明（None = 通过）。纯函数。

    anchor_paragraph_id 为 None / 缺省 = 插入到文档末尾（空文档可用）。
    """
    _, anchor_error = _validate_insert_anchor(
        snapshot, anchor_paragraph_id, proposal_count, max_proposals=max_proposals
    )
    if anchor_error:
        return anchor_error

    if not isinstance(latex, str) or not latex.strip():
        return "latex 必须是非空的 LaTeX 源码字符串"
    if len(latex) > _MAX_LATEX_LENGTH:
        return f"latex 过长（超过 {_MAX_LATEX_LENGTH} 字符），请拆分或简化公式"
    return None


def validate_paragraph_proposal(
    snapshot: DocumentSnapshotPayload,
    anchor_paragraph_id: object,
    paragraph_text: object,
    proposal_count: int,
    *,
    max_proposals: int | None = None,
) -> str | None:
    """校验一条纯文字段落插入提案；返回错误说明（None = 通过）。纯函数。

    anchor_paragraph_id 为 None / 缺省 = 插入到文档末尾（空文档可用）；
    paragraph_text 非空、≤ _MAX_PARAGRAPH_INSERT_LENGTH 字符、
    ≤ _MAX_PARAGRAPH_INSERT_LINES 行（换行即分段，空行忽略）。
    """
    _, anchor_error = _validate_insert_anchor(
        snapshot, anchor_paragraph_id, proposal_count, max_proposals=max_proposals
    )
    if anchor_error:
        return anchor_error

    if not isinstance(paragraph_text, str) or not paragraph_text.strip():
        return "paragraph_text 必须是非空的文字内容"
    if len(paragraph_text) > _MAX_PARAGRAPH_INSERT_LENGTH:
        return f"paragraph_text 过长（超过 {_MAX_PARAGRAPH_INSERT_LENGTH} 字符），请拆分为多次插入"
    lines = [line for line in paragraph_text.split("\n") if line.strip()]
    if not lines:
        return "paragraph_text 不能只含空白字符"
    if len(lines) > _MAX_PARAGRAPH_INSERT_LINES:
        return f"paragraph_text 最多 {_MAX_PARAGRAPH_INSERT_LINES} 个段落（换行分段，当前 {len(lines)} 行）"
    return None


def render_outline(snapshot: DocumentSnapshotPayload) -> str:
    """大纲渲染（纯函数）。"""
    if not snapshot.outline:
        return "（文档没有标题结构）"
    lines = [
        f"{'  ' * max(0, item.level - 1)}{'#' * max(1, min(item.level, 6))} {item.title}".rstrip()
        for item in snapshot.outline
    ]
    return "\n".join(lines)


def search_snapshot(snapshot: DocumentSnapshotPayload, query: str) -> str:
    """在快照中按文字片段搜索段落（纯函数，供 search_paragraphs 工具复用）。"""
    if not query.strip():
        return "query 不能为空"
    hits = [p for p in snapshot.paragraphs if query in p.text]
    if not hits:
        return f"未找到包含「{query}」的段落"
    lines = [f"[{p.id}] {p.text[:80]}" for p in hits[:_SEARCH_PREVIEW_LIMIT]]
    if len(hits) > _SEARCH_PREVIEW_LIMIT:
        lines.append(f"（共 {len(hits)} 段命中，仅显示前 {_SEARCH_PREVIEW_LIMIT} 段）")
    return "\n".join(lines)


def render_snapshot(snapshot: DocumentSnapshotPayload) -> str:
    """把快照渲染为模型可读文本（纯函数）。空段不渲染（无需提案）。"""
    lines: list[str] = []
    if snapshot.outline:
        lines.append("【文档大纲】")
        lines.append(render_outline(snapshot))
        lines.append("")
    lines.append("【文档段落】（[id] 为段落 id，修改时作为 propose_edit 的 paragraph_id；标注「（表格内）」的段落位于表格单元格中，不能作为插入类工具的锚点）")
    has_content = False
    for paragraph in snapshot.paragraphs:
        if not paragraph.text.strip():
            continue
        has_content = True
        if paragraph.in_table:
            style_note = "（表格内）"
        elif paragraph.level is not None:
            style_note = f"（标题 {paragraph.level}）"
        elif paragraph.style:
            style_note = f"（样式：{paragraph.style}）"
        else:
            style_note = ""
        lines.append(f"[{paragraph.id}]{style_note} {paragraph.text}")
    if not has_content:
        lines.append("（文档为空 —— 插入类工具请省略 anchor_paragraph_id，内容将写入文档末尾）")
    if snapshot.truncated:
        lines.append("（注意：快照已截断，仅包含文档前一部分）")
    return "\n".join(lines)


def build_agent_input(request: AgentStreamRequest) -> str:
    """run 输入：快照 + 历史对话（近 10 轮）+ 当前焦点 + 用户任务（纯函数）。"""
    parts: list[str] = [render_snapshot(request.snapshot)]
    history = request.history[-_HISTORY_LIMIT:]
    if history:
        lines = ["【历史对话】"]
        for entry in history:
            role = "用户" if entry.role == "user" else "助手"
            content = entry.content
            if len(content) > _HISTORY_ENTRY_PREVIEW:
                content = content[:_HISTORY_ENTRY_PREVIEW] + "…"
            lines.append(f"{role}：{content}")
        parts.append("\n".join(lines))
    if request.focus:
        lines = ["【当前光标 / 选区】用户说「这段」「这段话」「选中的内容」时指："]
        lines.append(f"段落 [{request.focus.paragraph_id}] {request.focus.paragraph_text}")
        if request.focus.selected_text:
            lines.append(f"用户当前选中的文字：「{request.focus.selected_text}」")
        parts.append("\n".join(lines))
    parts.append(f"【用户任务】\n{request.instruction}")
    return "\n\n".join(parts)


def _extract_usage(final: object) -> dict[str, int]:
    """从框架最终响应中 best-effort 提取 usage。"""
    usage = getattr(final, "usage", None)
    if usage is None:
        return {}
    try:
        prompt = int(getattr(usage, "prompt_tokens", 0) or 0)
        completion = int(getattr(usage, "completion_tokens", 0) or 0)
    except (TypeError, ValueError):
        return {}
    if prompt or completion:
        return {"prompt_tokens": prompt, "completion_tokens": completion}
    return {}


# ------------------------------------------------------------------
# 框架接入（懒 import）
# ------------------------------------------------------------------


def _load_framework():
    """懒加载 agent_framework；缺失时抛 AGENT_NOT_INSTALLED（其余功能不受影响）。"""
    try:
        from agent_framework import Agent, tool  # type: ignore[import-untyped]
    except ImportError as exc:
        raise AgentServiceError(
            "AGENT_NOT_INSTALLED",
            "后端未安装 agent-framework：py -m pip install agent-framework-core agent-framework-openai",
        ) from exc
    try:
        from agent_framework.openai import OpenAIChatCompletionClient  # type: ignore[import-untyped]
    except ImportError as exc:
        raise AgentServiceError(
            "AGENT_NOT_INSTALLED",
            "后端未安装 agent-framework-openai（OpenAI 兼容接入包）",
        ) from exc
    return Agent, tool, OpenAIChatCompletionClient


async def run_agent_stream(request: AgentStreamRequest) -> AsyncIterator[AgentEvent]:
    """运行 Agent 并按序产出事件（token / proposal / done）。

    失败抛 AgentServiceError（含超时 / 未安装 / 模型错误）；
    客户端断开时由消费方取消（CancelledError 透传）。
    """
    if not config.llm_configured():
        raise AgentServiceError("LLM_NOT_CONFIGURED", "未配置 LLM_API_KEY / LLM_MODEL")

    Agent, tool, OpenAIChatCompletionClient = _load_framework()

    system_prompt = prompt_service.load_prompt("agent.md") or _FALLBACK_SYSTEM_PROMPT
    skill_prompt = skill_service.render_skill_prompt(request.instruction)
    if skill_prompt:
        system_prompt += f"\n\n{skill_prompt}"
    formula_with_text = requests_formula_with_text(request.instruction)
    if formula_with_text:
        system_prompt += (
            "\n\n## 当前任务强制路由\n"
            "用户要求把公式和解释/正文一起写入。当前任务不提供 insert_formula；"
            "必须调用且只调用一次 insert_paragraph，把公式以 $LaTeX$ 嵌在完整说明中。"
            "不要只在聊天回答里给出说明。"
        )
    snapshot = request.snapshot
    pending: list[
        ProposalEvent
        | FormatProposalEvent
        | TableProposalEvent
        | FormulaProposalEvent
        | ParagraphProposalEvent
    ] = []
    tool_calls = {"count": 0}
    queue: asyncio.Queue[AgentEvent | _ErrorSentinel | None] = asyncio.Queue()

    def _check_budget() -> str | None:
        """工具调用预算：超限信息回喂模型（软限制；硬兜底是总超时）。"""
        tool_calls["count"] += 1
        if tool_calls["count"] > config.AGENT_MAX_TOOL_CALLS:
            return (
                f"工具调用次数已达上限（{config.AGENT_MAX_TOOL_CALLS}），"
                "请基于已有信息直接给出最终回答。"
            )
        return None

    @tool(approval_mode="never_require")
    def get_outline() -> str:
        """获取文档大纲（标题层级结构）。"""
        budget = _check_budget()
        if budget:
            return budget
        return render_outline(snapshot)

    @tool(approval_mode="never_require")
    def read_paragraph(paragraph_id: str) -> str:
        """按 id 读取段落的完整文本。

        Args:
            paragraph_id: 段落 id（文档快照中 [id] 括号内的值）。
        """
        budget = _check_budget()
        if budget:
            return budget
        paragraph = find_paragraph(snapshot, paragraph_id)
        if paragraph is None:
            return f"段落 {paragraph_id} 不存在，请使用文档快照中的段落 id"
        return f"[{paragraph.id}] {paragraph.text}"

    @tool(approval_mode="never_require")
    def search_paragraphs(query: str) -> str:
        """在全文中搜索包含指定文字片段的段落，返回段落 id 与内容预览。

        Args:
            query: 要搜索的文字片段。
        """
        budget = _check_budget()
        if budget:
            return budget
        return search_snapshot(snapshot, query)

    @tool(approval_mode="never_require")
    def propose_edit(paragraph_id: str, new_text: str, summary: str) -> str:
        """提交一条整段替换的修改提案；每处需要修改的段落各调用一次。

        Args:
            paragraph_id: 要修改的段落 id（文档快照中 [id] 括号内的值）。
            new_text: 该段修改后的完整文本（整段完整替换，不是增量补丁）。
            summary: 一句话说明本次修改，不超过 30 字。
        """
        budget = _check_budget()
        if budget:
            return budget
        error = validate_proposal(snapshot, paragraph_id, new_text, len(pending))
        if error:
            return f"提案被拒绝：{error}"
        paragraph = find_paragraph(snapshot, paragraph_id)
        assert paragraph is not None  # validate_proposal 已确认存在
        pending.append(
            ProposalEvent(
                paragraph_id=paragraph_id,
                original_text=paragraph.text,  # 后端从快照回显（前端为权威原文）
                new_text=_clean_proposal_text(new_text),
                summary=summary.strip() or "批量修改",
            )
        )
        return (
            f"已记录第 {len(pending)} 条提案（段落 {paragraph_id}）。"
            "继续处理其余段落，完成后请总结。"
        )

    @tool(approval_mode="never_require")
    def propose_format(
        paragraph_id: str,
        summary: str,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        color: str | None = None,
        alignment: str | None = None,
    ) -> str:
        """修改指定段落的格式（不改文字）。只传需要修改的字段，未传的保持不变。

        Args:
            paragraph_id: 要修改格式的段落 id（文档快照中 [id] 括号内的值）。
            summary: 一句话说明本次修改，不超过 30 字。
            bold: 加粗（true 开 / false 关）。
            italic: 斜体（true 开 / false 关）。
            underline: 下划线（true 开 / false 关）。
            strikethrough: 删除线（true 开 / false 关）。
            font_name: 字体名称，如 "微软雅黑"、"宋体"。
            font_size: 字号（磅，1-100）。
            color: 文字颜色，#RRGGBB 格式，如 "#FF0000"。
            alignment: 对齐方式：left / center / right / justify。
        """
        budget = _check_budget()
        if budget:
            return budget
        error = validate_format_proposal(
            snapshot,
            paragraph_id,
            len(pending),
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            color=color,
            alignment=alignment,
        )
        if error:
            return f"提案被拒绝：{error}"
        paragraph = find_paragraph(snapshot, paragraph_id)
        assert paragraph is not None  # validate 已确认存在
        try:
            pending.append(
                FormatProposalEvent(
                    paragraph_id=paragraph_id,
                    original_text=paragraph.text,  # 后端从快照回显（前端为权威原文）
                    summary=summary.strip() or "格式修改",
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    strikethrough=strikethrough,
                    font_name=font_name.strip() if isinstance(font_name, str) else font_name,
                    font_size=font_size,
                    color=color.upper() if isinstance(color, str) else color,
                    alignment=alignment,  # type: ignore[arg-type] — validate 已限定枚举
                )
            )
        except ValidationError as exc:
            return f"提案被拒绝：参数无效（{exc.errors()[0].get('msg', '格式错误')}）"
        return (
            f"已记录第 {len(pending)} 条格式提案（段落 {paragraph_id}）。"
            "继续处理，完成后请总结。"
        )

    @tool(approval_mode="never_require")
    def insert_table(
        anchor_paragraph_id: str | None,
        values: list[list[str]],
        summary: str,
        header: bool = True,
    ) -> str:
        """插入新表格（成为待审修订，用户可接受或拒绝）。

        Args:
            anchor_paragraph_id: 锚点段落 id，表格插入到该段之后；省略或 null = 插入到文档末尾（文档为空时也用 null）。
            values: 表格内容，字符串二维数组（每行列数一致，最多 20 行 x 8 列）。
            summary: 一句话说明，不超过 30 字。
            header: 第一行是否按表头处理（默认 true）。
        """
        budget = _check_budget()
        if budget:
            return budget
        error = validate_table_proposal(snapshot, anchor_paragraph_id, values, len(pending))
        if error:
            return f"提案被拒绝：{error}"
        anchor = find_paragraph(snapshot, anchor_paragraph_id) if anchor_paragraph_id else None
        try:
            pending.append(
                TableProposalEvent(
                    anchor_paragraph_id=anchor_paragraph_id,
                    anchor_text=anchor.text if anchor else "",  # 锚点原文回显（前端哈希校验用）
                    summary=summary.strip() or "插入表格",
                    header=header,
                    values=values,
                )
            )
        except ValidationError as exc:
            return f"提案被拒绝：参数无效（{exc.errors()[0].get('msg', '格式错误')}）"
        where = f"锚点段落 {anchor_paragraph_id} 之后" if anchor_paragraph_id else "文档末尾"
        return (
            f"已记录第 {len(pending)} 条表格提案（插入到{where}）。"
            "继续处理，完成后请总结。"
        )

    @tool(approval_mode="never_require")
    def insert_formula(
        anchor_paragraph_id: str | None,
        latex: str,
        summary: str,
        display: bool = True,
    ) -> str:
        """只插入一个独立数学公式（LaTeX 语法，自动转为 Word 公式）。

        用户同时要求解释、定义或正文时不要调用本工具；应调用 insert_paragraph，
        并在 paragraph_text 中用 $LaTeX$ 把公式和全部说明合并提交。

        Args:
            anchor_paragraph_id: 锚点段落 id，公式插入到该段之后；省略或 null = 插入到文档末尾（文档为空时也用 null）。
            latex: LaTeX 源码，如 "E=mc^2"、"\\\\frac{a}{b}"。
            summary: 一句话说明，不超过 30 字。
            display: true 独立成行居中；false 行内公式（默认 true）。
        """
        budget = _check_budget()
        if budget:
            return budget
        error = validate_formula_proposal(snapshot, anchor_paragraph_id, latex, len(pending))
        if error:
            return f"提案被拒绝：{error}"
        anchor = find_paragraph(snapshot, anchor_paragraph_id) if anchor_paragraph_id else None
        try:
            pending.append(
                FormulaProposalEvent(
                    anchor_paragraph_id=anchor_paragraph_id,
                    anchor_text=anchor.text if anchor else "",  # 锚点原文回显（前端哈希校验用）
                    summary=summary.strip() or "插入公式",
                    latex=latex,
                    display=display,
                )
            )
        except ValidationError as exc:
            return f"提案被拒绝：参数无效（{exc.errors()[0].get('msg', '格式错误')}）"
        where = f"锚点段落 {anchor_paragraph_id} 之后" if anchor_paragraph_id else "文档末尾"
        return (
            f"已记录第 {len(pending)} 条公式提案（插入到{where}）。"
            "继续处理，完成后请总结。"
        )

    @tool(approval_mode="never_require")
    def insert_paragraph(
        anchor_paragraph_id: str | None,
        paragraph_text: str,
        summary: str,
    ) -> str:
        """插入文字段落，可包含 $LaTeX$ 行内公式（用户可接受或拒绝）。

        「写个公式并解释」之类的混合内容应整体使用本工具，不要只提交公式。

        Args:
            anchor_paragraph_id: 锚点段落 id，新内容插入到该段之后；省略或 null = 插入到文档末尾（文档为空时也用 null）。
            paragraph_text: 要插入的文字（换行 \\n 分段，最多 20 段、共 2000 字；行内公式用 $LaTeX$ 标注）。
            summary: 一句话说明，不超过 30 字。
        """
        budget = _check_budget()
        if budget:
            return budget
        if formula_with_text:
            mixed_error = validate_formula_explanation_paragraph(paragraph_text)
            if mixed_error:
                return f"提案被拒绝：{mixed_error}"
        error = validate_paragraph_proposal(snapshot, anchor_paragraph_id, paragraph_text, len(pending))
        if error:
            return f"提案被拒绝：{error}"
        anchor = find_paragraph(snapshot, anchor_paragraph_id) if anchor_paragraph_id else None
        try:
            pending.append(
                ParagraphProposalEvent(
                    anchor_paragraph_id=anchor_paragraph_id,
                    anchor_text=anchor.text if anchor else "",  # 锚点原文回显（前端哈希校验用）
                    paragraph_text=paragraph_text,
                    summary=summary.strip() or "插入文字",
                )
            )
        except ValidationError as exc:
            return f"提案被拒绝：参数无效（{exc.errors()[0].get('msg', '格式错误')}）"
        where = f"锚点段落 {anchor_paragraph_id} 之后" if anchor_paragraph_id else "文档末尾"
        return (
            f"已记录第 {len(pending)} 条文字提案（插入到{where}）。"
            "继续处理，完成后请总结。"
        )

    client = OpenAIChatCompletionClient(
        base_url=config.LLM_BASE_URL,
        api_key=config.LLM_API_KEY,
        model=config.LLM_MODEL,
    )
    # 用户要求“公式 + 解释/正文”时，不把独立公式工具暴露给模型。
    # 之前在工具内部拒绝会让模型反复重试 insert_formula，最终触发
    # Function invocation limit；从工具集合中移除后，模型只能用支持
    # $LaTeX$ 富段落的 insert_paragraph，一次提交完整内容。
    mutation_tools = [propose_edit, propose_format, insert_table, insert_paragraph]
    if not formula_with_text:
        mutation_tools.insert(3, insert_formula)

    agent = Agent(
        client=client,
        name="word-ai-copilot",
        instructions=system_prompt,
        tools=[
            get_outline,
            read_paragraph,
            search_paragraphs,
            *mutation_tools,
        ],
    )
    user_input = build_agent_input(request)

    async def worker() -> None:
        """在独立 task 中驱动框架迭代（单 task 端到端，避免跨任务恢复）。"""
        try:
            stream = agent.run(user_input, stream=True)
            async with asyncio.timeout(config.AGENT_TIMEOUT_SECONDS):
                async for update in stream:
                    while pending:
                        await queue.put(pending.pop(0))
                    text = getattr(update, "text", None)
                    if text:
                        await queue.put(TokenEvent(text))
                while pending:
                    await queue.put(pending.pop(0))
            usage: dict[str, int] = {}
            with contextlib.suppress(Exception):
                final = await asyncio.wait_for(stream.get_final_response(), timeout=2.0)
                usage = _extract_usage(final)
            await queue.put(DoneEvent(usage))
        except asyncio.CancelledError:
            raise
        except AgentServiceError as exc:
            await queue.put(_ErrorSentinel(exc.code, exc.message))
        except TimeoutError:
            await queue.put(
                _ErrorSentinel(
                    "LLM_TIMEOUT",
                    f"Agent 运行超时（超过 {config.AGENT_TIMEOUT_SECONDS:.0f} 秒）",
                )
            )
        except Exception as exc:  # noqa: BLE001 — 框架异常统一转译为错误帧
            logger.exception("agent 运行异常 conversation=%s", request.conversation_id)
            await queue.put(_ErrorSentinel("LLM_ERROR", f"模型调用失败：{exc}"))
        finally:
            await queue.put(None)  # 结束哨兵

    task = asyncio.create_task(worker())
    try:
        while True:
            event = await queue.get()
            if event is None:
                break
            if isinstance(event, _ErrorSentinel):
                raise AgentServiceError(event.code, event.message)
            yield event
    finally:
        # 消费方提前退出（客户端断开 / 异常）→ 取消后台运行
        if not task.done():
            task.cancel()
            with contextlib.suppress(Exception):
                await task
