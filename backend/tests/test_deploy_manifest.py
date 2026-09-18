"""deploy/make_manifest.py 的单元测试（纯函数 + CLI 端到端）。

manifest 是分发给使用者的唯一文件，替换出错会导致 Word 静默拒绝加载，
这里钉住：全部开发地址替换、外部 URL 不受影响、AppDomain 规范化、
校验规则（https / 非 localhost / GUID / XML 良构）、CLI 输出文件。
"""
import importlib.util
import sys
from pathlib import Path

import pytest

_REPO_ROOT = Path(__file__).resolve().parents[2]


def _load_module():
    spec = importlib.util.spec_from_file_location(
        "deploy_make_manifest", _REPO_ROOT / "deploy" / "make_manifest.py"
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def mm():
    return _load_module()


@pytest.fixture(scope="module")
def source_xml() -> str:
    text = (_REPO_ROOT / "frontend" / "manifest.xml").read_text(encoding="utf-8")
    assert text.count("https://localhost:3000") == 8  # 前置：源 manifest 形态稳定
    return text


class TestBuildManifest:
    def test_替换全部开发地址且外部_URL_不受影响(self, mm, source_xml):
        result = mm.build_manifest(source_xml, "https://wordai.intranet.local:8443")
        assert "https://localhost:3000" not in result
        assert result.count("https://wordai.intranet.local:8443") == 8
        # 学习链接是外部 URL，不能被碰
        assert "https://learn.microsoft.com/office/dev/add-ins/develop/office-versions" in result

    def test_关键元素都被替换(self, mm, source_xml):
        result = mm.build_manifest(source_xml, "https://wordai.example.com")
        assert '<SourceLocation DefaultValue="https://wordai.example.com/taskpane.html"' in result
        assert "<AppDomain>https://wordai.example.com</AppDomain>" in result
        assert 'DefaultValue="https://wordai.example.com/assets/icon-16.png"' in result

    def test_子路径时_AppDomain_只保留源(self, mm, source_xml):
        result = mm.build_manifest(source_xml, "https://host.example.com/addins/wordai")
        assert "<AppDomain>https://host.example.com</AppDomain>" in result
        assert "https://host.example.com/addins/wordai/taskpane.html" in result

    def test_尾部斜杠被归一化(self, mm, source_xml):
        result = mm.build_manifest(source_xml, "https://host.example.com/")
        assert "https://host.example.com/taskpane.html" in result

    def test_产物是良构_XML(self, mm, source_xml):
        import xml.etree.ElementTree as ET

        result = mm.build_manifest(source_xml, "https://host.example.com")
        ET.fromstring(result)  # 不抛即通过


class TestValidation:
    def test_http_被拒绝(self, mm, source_xml):
        with pytest.raises(ValueError, match="https"):
            mm.build_manifest(source_xml, "http://host.example.com")

    def test_localhost_被拒绝并提示用开发版(self, mm, source_xml):
        with pytest.raises(ValueError, match="localhost"):
            mm.build_manifest(source_xml, "https://localhost:3001")

    def test_带_query_被拒绝(self, mm, source_xml):
        with pytest.raises(ValueError, match="query"):
            mm.build_manifest(source_xml, "https://host.example.com?x=1")

    def test_非法_GUID_被拒绝(self, mm, source_xml):
        with pytest.raises(ValueError):
            mm.build_manifest(source_xml, "https://host.example.com", "not-a-guid")

    def test_new_id_替换顶层_Id(self, mm, source_xml):
        new_id = "12345678-1234-1234-1234-123456789abc"
        result = mm.build_manifest(source_xml, "https://host.example.com", new_id)
        assert f"<Id>{new_id}</Id>" in result
        assert "<Id>6f2e5d3a-8b7c-4a9e-b1d2-3c4f5a6b7c8d</Id>" not in result

    def test_缺主机名_被拒绝(self, mm, source_xml):
        with pytest.raises(ValueError, match="主机名"):
            mm.normalize_base("https://")


class TestCli:
    def test_端到端生成文件(self, mm, tmp_path, capsys):
        out = tmp_path / "share" / "manifest.xml"
        code = mm.main(["--base", "https://wordai.intranet.local:8443", "-o", str(out)])
        assert code == 0
        content = out.read_text(encoding="utf-8")
        assert "https://localhost:3000" not in content
        assert content.count("https://wordai.intranet.local:8443") == 8
        # 摘要输出包含下一步指引
        captured = capsys.readouterr()
        assert "BACKEND_CORS_ORIGINS" in captured.out

    def test_非法参数返回非零退出码(self, mm, capsys):
        code = mm.main(["--base", "http://bad.example.com"])
        assert code == 1
        assert "错误" in capsys.readouterr().err

    def test_默认输出到_deploy_dist(self, mm):
        out = _REPO_ROOT / "deploy" / "dist" / "manifest.xml"
        try:
            code = mm.main(["--base", "https://host.example.com"])
            assert code == 0
            content = out.read_text(encoding="utf-8")
            assert "https://localhost:3000" not in content
        finally:
            if out.exists():
                out.unlink()


def test_cli_module_importable_as_script():
    """脚本可独立执行（python deploy/make_manifest.py --help 不炸）。"""
    import subprocess

    proc = subprocess.run(
        [sys.executable, str(_REPO_ROOT / "deploy" / "make_manifest.py"), "--help"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert proc.returncode == 0
    assert "--base" in proc.stdout
