"""md-lite 解析器的回归测试：md 方言一变，报告排版整篇就会跟着变，所以这块必须有测试。"""

import json
import shutil
import sys
import unittest
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from lrkit.mdlite import expand, parse_blocks  # noqa: E402

_TMP_ROOT = Path(__file__).resolve().parent / "_tmp"


def scratch_dir() -> Path:
    """不用 tempfile：本机沙箱把 mkdtemp 建出来的目录判为不可访问。"""
    _TMP_ROOT.mkdir(parents=True, exist_ok=True)
    path = _TMP_ROOT / uuid.uuid4().hex[:12]
    path.mkdir()
    return path


class ParseBlocksTest(unittest.TestCase):
    def test_headings_and_paragraphs(self):
        blocks = parse_blocks("# 一、目的\n## 1.1 掌握方法\n正文第一段。\n\n正文第二段。")
        self.assertEqual(
            blocks,
            [
                ("heading", 1, "一、目的"),
                ("heading", 2, "1.1 掌握方法"),
                ("para", "正文第一段。"),
                ("para", "正文第二段。"),
            ],
        )

    def test_table_skips_separator_and_pads_rows(self):
        text = "@table 表1 电阻\n| t/s | R/Ω |\n| --- | --- |\n| 0 | 100 |\n| 40 |\n"
        blocks = parse_blocks(text)
        self.assertEqual(len(blocks), 1)
        kind, data = blocks[0]
        self.assertEqual(kind, "table")
        self.assertEqual(data["caption"], "表1 电阻")
        self.assertEqual(data["columns"], ["t/s", "R/Ω"])
        self.assertEqual(data["rows"], [["0", "100"], ["40", ""]])

    def test_table_without_body_rows_is_error(self):
        with self.assertRaises(ValueError):
            parse_blocks("@table 表1 空表\n没有数据行") 

    def test_figure_with_caption_and_width(self):
        blocks = parse_blocks("![图1 冷却曲线](figures/fig1.png){width=10.5}")
        kind, data = blocks[0]
        self.assertEqual(kind, "figure")
        self.assertEqual(data["caption"], "图1 冷却曲线")
        self.assertEqual(data["path"], "figures/fig1.png")
        self.assertAlmostEqual(data["width"], 10.5)

    def test_figure_without_width_has_no_key(self):
        blocks = parse_blocks("![图2 示意图](media/a.png)")
        self.assertNotIn("width", blocks[0][1])

    def test_numbered_formula(self):
        blocks = parse_blocks(r"@formula \bar{x}=\frac{1}{n}\sum x_i #(3)")
        kind, data = blocks[0]
        self.assertEqual(kind, "formula")
        self.assertEqual(data["number"], "3")
        self.assertEqual(data["tex"], r"\bar{x}=\frac{1}{n}\sum x_i")

    def test_dollar_display_formula(self):
        blocks = parse_blocks(r"$$\lambda=\frac{2L}{n}$$")
        self.assertEqual(blocks[0], ("formula", {"tex": r"\lambda=\frac{2L}{n}", "number": None}))

    def test_pagebreak(self):
        self.assertEqual(parse_blocks("@pagebreak"), [("pagebreak",)])

    def test_from_reference_is_kept_raw(self):
        self.assertEqual(parse_blocks("@table from=tab1")[0], ("table", {"from": "tab1"}))
        self.assertEqual(parse_blocks("@figure from=fig1")[0], ("figure", {"from": "fig1"}))


class ExpandTest(unittest.TestCase):
    def test_expand_reads_calc_json(self):
        calc = scratch_dir()
        try:
            (calc / "tables").mkdir()
            (calc / "figures").mkdir()
            (calc / "tables" / "tab1.json").write_text(
                json.dumps({"caption": "表1 温度", "columns": ["t"], "rows": [["1"]]}, ensure_ascii=False),
                encoding="utf-8",
            )
            (calc / "figures" / "fig1.json").write_text(
                json.dumps({"path": "figures/a.png", "caption": "图1"}, ensure_ascii=False), encoding="utf-8"
            )
            blocks = parse_blocks("@table from=tab1\n@figure from=fig1")
            out = expand(blocks, calc)
            self.assertEqual(out[0][0], "table")
            self.assertEqual(out[0][1]["id"], "tab1")
            self.assertEqual(out[0][1]["columns"], ["t"])
            self.assertEqual(out[1][0], "figure")
            self.assertEqual(out[1][1]["id"], "fig1")
        finally:
            shutil.rmtree(calc, ignore_errors=True)

    def test_expand_missing_file_raises(self):
        calc = scratch_dir()
        try:
            with self.assertRaises(FileNotFoundError):
                expand(parse_blocks("@table from=nope"), calc)
        finally:
            shutil.rmtree(calc, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
