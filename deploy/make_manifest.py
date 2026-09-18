"""生成分发版 manifest：把 frontend/manifest.xml 中的本机开发地址
（https://localhost:3000）替换为实际对外服务的 HTTPS 源。

分发给使用者的唯一文件就是产出的 manifest.xml —— 插件本体（任务窗格前端）
与后端都运行在服务方机器上，使用者机器无需安装任何东西（共享目录法）。

用法（仓库根目录执行）：
    py deploy\\make_manifest.py --base https://wordai.intranet.local:8443
    py deploy\\make_manifest.py --base https://wordai.example.com --new-id
    py deploy\\make_manifest.py --base https://wordai.example.com -o share\\manifest.xml

产物默认写到 deploy/dist/manifest.xml（该目录已被 Git 忽略）。
前置条件与分发步骤见 deploy/README.md。

只用标准库，无第三方依赖。
"""
from __future__ import annotations

import argparse
import re
import sys
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urlsplit

# 开发版 manifest 中的本机源（frontend/manifest.xml）
DEV_ORIGIN = "https://localhost:3000"

# <Id>GUID</Id>（manifest 顶层元素，Word 以此识别插件身份）
_ID_PATTERN = re.compile(r"(<Id>)([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
                         r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})(</Id>)")


def normalize_base(base: str) -> str:
    """校验并归一化对外服务源。允许 https://host[:port][/子路径]，去尾部斜杠。"""
    parts = urlsplit(base.strip())
    if parts.scheme != "https":
        raise ValueError("base 必须是 https://…（Word 任务窗格强制要求 HTTPS，http 会被拒绝加载）")
    if not parts.hostname:
        raise ValueError("base 缺少主机名")
    if parts.query or parts.fragment:
        raise ValueError("base 不能携带 query 或 fragment")
    if parts.hostname in {"localhost", "127.0.0.1", "::1"}:
        raise ValueError(
            "base 指向 localhost：分发版 manifest 只在服务所在的那台机器上有效。"
            "本机开发请直接使用 frontend/manifest.xml（npm run sideload）。"
        )
    return f"{parts.scheme}://{parts.netloc}{parts.path}".rstrip("/")


def replace_id(xml_text: str, new_id: str) -> str:
    """替换 manifest 顶层 <Id>（换新插件身份时使用）。"""
    uuid.UUID(new_id)  # 非法 GUID 直接抛 ValueError
    if not _ID_PATTERN.search(xml_text):
        raise ValueError("源 manifest 中未找到 <Id> 元素")
    return _ID_PATTERN.sub(rf"\g<1>{new_id}\g<3>", xml_text, count=1)


def build_manifest(source_xml: str, base: str, new_id: str | None = None) -> str:
    """把 source_xml 中的开发源全部替换为 base，返回分发版 XML 文本。"""
    normalized = normalize_base(base)
    count = source_xml.count(DEV_ORIGIN)
    if count == 0:
        raise ValueError("源 manifest 中未找到开发地址（frontend/manifest.xml 内容可能已变更）")
    result = source_xml.replace(DEV_ORIGIN, normalized)

    # AppDomain 按规范只写源（scheme://host[:port]），不带子路径
    origin = f"{urlsplit(normalized).scheme}://{urlsplit(normalized).netloc}"
    if normalized != origin:
        result = result.replace(f"<AppDomain>{normalized}</AppDomain>", f"<AppDomain>{origin}</AppDomain>")

    if new_id is not None:
        result = replace_id(result, new_id)

    ET.fromstring(result)  # 产物必须是良构 XML（替换出错在这里抛出）
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="make_manifest",
        description="生成指向对外 HTTPS 源的分发版 manifest（分发给使用者仅此一个文件）",
    )
    parser.add_argument(
        "--base",
        required=True,
        help="任务窗格实际服务的 HTTPS 源，如 https://wordai.intranet.local:8443（可含子路径）",
    )
    parser.add_argument(
        "--new-id",
        action="store_true",
        help="生成新的插件 GUID（默认沿用源 manifest 的 Id：跨机器互不影响；"
             "同一台机器上与 localhost 开发版互相顶替）",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="输出路径（默认 deploy/dist/manifest.xml）",
    )
    args = parser.parse_args(argv)

    repo_root = Path(__file__).resolve().parents[1]
    source_path = repo_root / "frontend" / "manifest.xml"
    if not source_path.exists():
        print(f"找不到源 manifest：{source_path}", file=sys.stderr)
        return 2
    source_xml = source_path.read_text(encoding="utf-8")

    try:
        result = build_manifest(source_xml, args.base, str(uuid.uuid4()) if args.new_id else None)
    except ValueError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1

    output = Path(args.output) if args.output else repo_root / "deploy" / "dist" / "manifest.xml"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(result, encoding="utf-8")

    base = normalize_base(args.base)
    print(f"已生成：{output}")
    print(f"  指向源：{base}（替换 {source_xml.count(DEV_ORIGIN)} 处开发地址）")
    print("服务方自检：")
    print(f"  1. {base}/taskpane.html 可被使用者机器打开（证书需受信任）")
    print(f"  2. 同源 /api 反代到 FastAPI（8100）")
    print(f"  3. backend/.env 的 BACKEND_CORS_ORIGINS 加入 {base}")
    print("分发方式见 deploy/README.md（共享文件夹受信目录 / office-addin-dev-settings register）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
