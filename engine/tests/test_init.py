"""init 的模板选择与样例入库测试：.doc 降级、样例扫描、样例文本抽取。"""

import shutil
import sys
import unittest
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from lrkit import cmd_init  # noqa: E402

_TMP_ROOT = Path(__file__).resolve().parent / "_tmp"


def scratch_dir() -> Path:
    """不用 tempfile：本机沙箱把 mkdtemp 建出来的目录判为不可访问。"""
    _TMP_ROOT.mkdir(parents=True, exist_ok=True)
    path = _TMP_ROOT / uuid.uuid4().hex[:12]
    path.mkdir()
    return path


class PickTemplateTest(unittest.TestCase):
    def setUp(self):
        self.dir = scratch_dir()

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_prefers_docx(self):
        (self.dir / "实验报告报告头-word.docx").write_bytes(b"x")
        (self.dir / "实验报告报告头-word.doc").write_bytes(b"x")
        tpl, legacy = cmd_init._pick_template(self.dir, {})
        self.assertEqual(tpl.name, "实验报告报告头-word.docx")
        self.assertIsNone(legacy)

    def test_legacy_doc_is_reported(self):
        (self.dir / "实验报告报告头-word.doc").write_bytes(b"x")
        tpl, legacy = cmd_init._pick_template(self.dir, {})
        self.assertIsNone(tpl)
        self.assertEqual(legacy.name, "实验报告报告头-word.doc")

    def test_explicit_doc_param_is_reported(self):
        doc = self.dir / "模板.doc"
        doc.write_bytes(b"x")
        tpl, legacy = cmd_init._pick_template(self.dir, {"template": str(doc)})
        self.assertIsNone(tpl)
        self.assertEqual(legacy, doc)

    def test_nothing_found(self):
        self.assertEqual(cmd_init._pick_template(self.dir, {}), (None, None))


class SampleTest(unittest.TestCase):
    def setUp(self):
        self.dir = scratch_dir()

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_scans_sample_folder(self):
        folder = self.dir / "样例"
        folder.mkdir()
        (folder / "a.docx").write_bytes(b"x")
        (folder / "b.pdf").write_bytes(b"x")
        (folder / "c.zip").write_bytes(b"x")
        names = [p.name for p in cmd_init._sample_sources(self.dir, {})]
        self.assertEqual(sorted(names), ["a.docx", "b.pdf"])

    def test_explicit_sample_wins(self):
        other = self.dir / "x" / "范文.docx"
        other.parent.mkdir()
        other.write_bytes(b"x")
        got = cmd_init._sample_sources(self.dir, {"sample": str(other)})
        self.assertEqual([p.name for p in got], ["范文.docx"])

    def test_docx_to_text_keeps_headings_tables(self):
        from docx import Document

        doc = Document()
        doc.add_heading("一、目的及要求", level=1)
        doc.add_paragraph("熟悉温度测量方法。")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "t/℃"
        table.cell(0, 1).text = "E/mV"
        table.cell(1, 0).text = "30"
        table.cell(1, 1).text = "0.5"
        src = self.dir / "sample.docx"
        doc.save(str(src))

        text = cmd_init.docx_to_text(src)
        self.assertIn("# 一、目的及要求", text)
        self.assertIn("熟悉温度测量方法。", text)
        self.assertIn("| t/℃ | E/mV |", text)
        self.assertIn("| 30 | 0.5 |", text)

    def test_ingest_copies_and_extracts(self):
        from lrkit.config import Config

        cfg = Config.from_env({"project_dir": str(self.dir / "_lr"), "data_dir": str(self.dir / "_lr" / "data")})
        (self.dir / "样例").mkdir()
        src = self.dir / "样例" / "范文.md"
        src.write_text("样例正文", encoding="utf-8")
        made = cmd_init._ingest_samples(cfg, [src])
        self.assertEqual(made[0]["kind"], "md")
        self.assertTrue((cfg.book / "sample" / "范文.md").exists())
        self.assertGreater(made[0]["chars"], 0)


if __name__ == "__main__":
    unittest.main()
