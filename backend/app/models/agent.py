"""Agent 请求 / 事件模型（§51 Agent 层：批量编辑 + 文档问答）。

Agent 接口 wire 格式为 snake_case（与 edit 一致）。
快照由前端从 Word 采集（paragraph uniqueLocalId 作为段落 id）；
提案（ProposalEvent）的 original_text 由后端从快照回显 —— 前端捕获的
快照文本为权威原文，保证客户端 Diff 坐标系有效（§19）。

提案共六类（判别字段 kind，SSE 帧名亦不同）：
- ProposalEvent（文本，kind 隐含为 "text"，帧名 proposal —— 不加 kind 字段
  以保持既有 wire 格式不变）
- FormatProposalEvent（格式，帧名 proposal_format）
- TableProposalEvent（表格，帧名 proposal_table）
- FormulaProposalEvent（公式，帧名 proposal_formula）
- ParagraphProposalEvent（纯文字段落，帧名 proposal_paragraph）
- HeadingProposalEvent（Word 内置标题样式，帧名 proposal_heading）

插入类提案（表格 / 公式 / 段落 / 标题）的 anchor_paragraph_id 可为 None —— 表示插入到
文档末尾（空文档 / 用户未指明位置时的默认行为），前端经 Body.insert*("End")
落地，无需锚点段落。
"""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from .chat import HistoryEntryPayload


class AgentParagraph(BaseModel):
    """快照段落（id = Word paragraph uniqueLocalId）。

    in_table：Body.paragraphs 自 WordApi 1.3 起包含表格单元格内的段落；
    表内段落不能作为插入类提案的锚点（内容会插进单元格）——
    服务层据此拒绝，快照渲染时也会标注（表格内）。
    """

    id: str = Field(min_length=1, max_length=64)
    text: str = Field(default="", max_length=50000)
    style: str | None = None
    level: int | None = None
    in_table: bool = False


class AgentOutlineItem(BaseModel):
    level: int = 1
    title: str = ""


class DocumentSnapshotPayload(BaseModel):
    """文档快照（前端扫描全文生成，超长时截断并标记）。"""

    outline: list[AgentOutlineItem] = Field(default_factory=list)
    paragraphs: list[AgentParagraph] = Field(default_factory=list, max_length=2000)
    truncated: bool = False


class AgentFocusContext(BaseModel):
    """当前光标 / 选区焦点（Agent 模式下前端随请求注入）。

    用户说「这段 / 选中的内容」时指此段落；selected_text 为非空选区文本。
    paragraph_text 以快照为准（前端已解析为快照段落文本）。
    """

    paragraph_id: str = Field(min_length=1, max_length=64)
    paragraph_text: str = Field(default="", max_length=50000)
    selected_text: str | None = Field(default=None, max_length=50000)


class AgentStreamRequest(BaseModel):
    conversation_id: str
    instruction: str = Field(min_length=1, max_length=2000)
    history: list[HistoryEntryPayload] = Field(default_factory=list)
    snapshot: DocumentSnapshotPayload
    focus: AgentFocusContext | None = None


class ProposalEvent(BaseModel):
    """SSE proposal 帧载荷：一条整段替换的修改提案。"""

    paragraph_id: str
    original_text: str
    new_text: str
    summary: str


FormatAlignment = Literal["left", "center", "right", "justify"]
ParagraphStyle = Literal[
    "normal",
    "title",
    "subtitle",
    "heading1",
    "heading2",
    "heading3",
    "heading4",
    "heading5",
    "heading6",
    "heading7",
    "heading8",
    "heading9",
]


class FormatProposalEvent(BaseModel):
    """SSE proposal_format 帧载荷：整段格式修改提案。

    None / 缺省 = 不修改该项；布尔 False = 显式关闭（如取消加粗）。
    original_text 语义同文本提案（前端权威原文校验用）。
    """

    kind: Literal["format"] = "format"
    paragraph_id: str
    original_text: str = ""
    summary: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    font_name: str | None = Field(default=None, max_length=64)
    font_size: float | None = Field(default=None, ge=1, le=100)
    color: str | None = Field(default=None, pattern=r"^#[0-9a-fA-F]{6}$")
    alignment: FormatAlignment | None = None
    paragraph_style: ParagraphStyle | None = None


class TableProposalEvent(BaseModel):
    """SSE proposal_table 帧载荷：在锚点段落之后插入表格。

    values 为矩形字符串矩阵（行 / 列 / 单元格长度上限见服务层校验）。
    """

    kind: Literal["insert-table"] = "insert-table"
    anchor_paragraph_id: str | None = None
    anchor_text: str = ""
    summary: str
    header: bool = True
    values: list[list[str]] = Field(min_length=1, max_length=20)


class FormulaProposalEvent(BaseModel):
    """SSE proposal_formula 帧载荷：在锚点段落之后插入公式（LaTeX）。"""

    kind: Literal["insert-formula"] = "insert-formula"
    anchor_paragraph_id: str | None = None
    anchor_text: str = ""
    summary: str
    latex: str = Field(min_length=1, max_length=500)
    display: bool = True


class ParagraphProposalEvent(BaseModel):
    """SSE proposal_paragraph 帧载荷：插入纯文字段落（锚点段后 / 文档末尾）。

    paragraph_text 可含换行（\\n）—— 每行一个段落，前端转为多个 w:p 插入。
    """

    kind: Literal["insert-paragraph"] = "insert-paragraph"
    anchor_paragraph_id: str | None = None
    anchor_text: str = ""
    paragraph_text: str = Field(min_length=1, max_length=2000)
    summary: str


class HeadingProposalEvent(BaseModel):
    """SSE proposal_heading 帧载荷：插入 Word 标题 1～9 段落。"""

    kind: Literal["insert-heading"] = "insert-heading"
    anchor_paragraph_id: str | None = None
    anchor_text: str = ""
    heading_text: str = Field(min_length=1, max_length=300)
    level: int = Field(ge=1, le=9)
    summary: str
