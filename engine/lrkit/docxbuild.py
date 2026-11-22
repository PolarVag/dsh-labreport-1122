"""把 md-lite 渲染成排好版的 .docx（在模板的表头页基础上往后写正文）。

排版口径来自《基础物理实验》书写要求：
  页边距 2 cm、正文宋体小四、行距固定 20 磅、首行缩进 2 字符；
  表题在表上方（五号加粗居中）、图题在图下方；三线表；公式为 Word 原生 OMML 居中；
  层次用「一、」「1.1」「（1）」，不使用项目符号；页脚「第 X 页」。
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .config import Config
from .util import read_json

BODY_CJK = "宋体"
HEAD_CJK = "黑体"
LATIN = "Times New Roman"
BODY_PT = 12.0
CAPTION_PT = 10.5
LINE_PT = 20.0

_INLINE = re.compile(r"(\*\*.+?\*\*|\$[^$]+\$)")


def _font(run, cjk=BODY_CJK, size=BODY_PT, bold=False, italic=False):
    run.font.name = LATIN
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cjk)
    rfonts.set(qn("w:ascii"), LATIN)
    rfonts.set(qn("w:hAnsi"), LATIN)


def _page_field(paragraph):
    run = paragraph.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE ")
    inner = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    inner.append(t)
    fld.append(inner)
    run._element.addnext(fld)


def _clear_body(doc):
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def _three_line(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (("top", 12), ("bottom", 12)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    for edge in ("left", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    tbl_pr.append(borders)


def _no_borders(table):
    """1×1 装图的容器表：四边和内部都不画线。"""
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    table._tbl.tblPr.append(borders)


def _header_bottom_line(table):
    for cell in table.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        borders.append(bottom)
        tc_pr.append(borders)


def _cell_text(cell, text):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    _font(run, size=CAPTION_PT)


class Builder:
    def __init__(self, cfg: Config, report: dict):
        self.cfg = cfg
        self.report = report
        self.styles = report.get("styles") or {}
        self.inline_omml: dict[str, str] = {}
        tpl = Path(report.get("template") or "") if report.get("template") else None
        if not tpl or not tpl.exists():
            cand = cfg.book / "template.docx"
            tpl = cand if cand.exists() else None
        self.doc = Document(str(tpl)) if tpl else Document()
        self.text_width = self._text_width()

    # ---------- 页面 ----------
    def _text_width(self) -> Cm:
        sec = self.doc.sections[0]
        return Cm((sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm))


    def setup_page(self):
        # 页边距、标题字体等一律以模板为准：不要在这里改 Normal 样式或页边距，
        # 否则模板里没写字号的下划线填空 run 会被顶成小四，表头就不再是原样。
        self.text_width = self._text_width()

    def footer_page_number(self):
        sec = self.doc.sections[0]
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("第 ")
        _font(run, size=9)
        _page_field(p)
        run2 = p.add_run(" 页")
        _font(run2, size=9)

    # ---------- 表头页 ----------
    def _blank_groups(self, para):
        """把段落里连续的下划线填空 run 归成一组（一组＝一个要填的空中括号）。"""
        groups, cur = [], []
        for run in para.runs:
            if run.font.underline:
                cur.append(run)
            elif cur:
                groups.append(cur)
                cur = []
        if cur:
            groups.append(cur)
        return groups

    def _set_group(self, runs, value: str, size=None, tail: str = ""):
        """把一组填空 run 的文字整体换成 value（保留首个 run 的下划线，字号显式统一）。

        模板里有的填空位（如"指导教师"后面那段）自带 18 pt 字号，直接写字会被撑大，
        这里按整段正文的标签字号统一，和样例报告的做法一致。
        """
        if not value:
            return
        run = runs[0]
        run.text = f" {value}{tail} "
        if size:
            run.font.size = Pt(size)
        for r in runs[1:]:
            r.text = ""

    def _label_size(self, para) -> float:
        """取整段里标签文字的字号。

        模板里的标签（"实验名称""指导教师"…）通常没有显式 sz，字号来自样式；
        而填空位里却可能带 18 pt 的孤立 run，不能拿它当基准。这里优先用标签 run
        的字号，其次是段落样式，最后退回正文小五 10.5 pt。
        """
        for run in para.runs:
            if not run.font.underline and run.font.size:
                return run.font.size.pt
        try:
            style_size = para.style.font.size
            if style_size:
                return style_size.pt
        except AttributeError:
            pass
        return 10.5

    def _date_parts(self, raw: str) -> tuple[str, str, str]:
        digits = re.findall(r"\d+", raw or "")
        if len(digits) >= 3:
            return digits[0], str(int(digits[1])), str(int(digits[2]))
        if len(digits) == 1 and len(digits[0]) == 8:
            d = digits[0]
            return d[:4], str(int(d[4:6])), str(int(d[6:8]))
        return "", "", ""

    def fill_header(self, meta: dict) -> bool:
        """在模板原有的表头上原地填空，不重绘、不分页。

        模板正文前四段就是表头：标题 / 实验名称+指导教师 / 姓名+学号+座号 /
        日期+地点+调课+成绩评定。每段的填空位都是带下划线的空白 run，
        这里只改这些 run 的文字，字体、字号、行距、下划线一律沿用模板。
        """
        paras = self.doc.paragraphs
        if len(paras) < 4:
            return False
        filled = False

        def fill_grouped(para, values, tails=None):
            nonlocal filled
            groups = self._blank_groups(para)
            size = self._label_size(para)
            for k, (runs, value) in enumerate(zip(groups, values)):
                if value:
                    tail = (tails or {}).get(k, "")
                    self._set_group(runs, value, size=size, tail=tail)
                    filled = True
            return len(groups)

        year, month, day = self._date_parts(meta.get("date", ""))
        makeup = bool(meta.get("makeup"))
        # 第 2 段：实验名称 / 指导教师
        fill_grouped(paras[1], [meta.get("experiment", ""), meta.get("teacher", "")])
        # 第 3 段：姓名 / 学号 / 分班分组及座号（座位号那格末尾的"号"字要保留）
        fill_grouped(
            paras[2],
            [meta.get("student", ""), meta.get("student_id", ""), meta.get("group", "")][: len(self._blank_groups(paras[2]))],
            tails={2: "号"},
        )
        # 第 4 段：日期（年/月/日三段空白分开）/ 实验地点 / 调课勾选 / 成绩评定留空
        groups = self._blank_groups(paras[3])
        size = self._label_size(paras[3])
        if len(groups) >= 5:
            for runs, value in ((groups[0], year), (groups[1], month), (groups[2], day), (groups[3], meta.get("place", ""))):
                if value:
                    self._set_group(runs, value, size=size)
                    filled = True
            box = groups[4]
            if box:
                box[0].text = " ☑是 " if makeup else " □是 "
                if size:
                    box[0].font.size = Pt(size)
                for run in box[1:]:
                    run.text = ""
        elif groups:
            fill_grouped(paras[3], [year, month, day, meta.get("place", "")])

        # 删掉表头后面那些空的占位段落，正文紧接表头
        for para in list(self.doc.paragraphs)[4:]:
            if para.text.strip():
                break
            para._p.getparent().remove(para._p)
        return filled

    def title_line(self, meta: dict):
        """实验题目：表头下面一行，居中加粗。"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(24)
        run = p.add_run(meta.get("experiment") or "")
        _font(run, cjk=HEAD_CJK, size=14, bold=True)

    def front_page(self, meta: dict):
        _clear_body(self.doc)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(40)
        p.paragraph_format.space_after = Pt(30)
        run = p.add_run(meta.get("title") or "《基础物理实验》实验报告")
        _font(run, cjk=HEAD_CJK, size=22, bold=True)

        rows = [
            [("实验名称：", meta.get("experiment", "")), ("指导教师：", meta.get("teacher", ""))],
            [("姓名：", meta.get("student", "")), ("学号：", meta.get("student_id", "")), ("分班分组及座号：", meta.get("group", ""))],
            [("实验日期：", meta.get("date", "")), ("实验地点：", meta.get("place", ""))],
            [("调课/补课：□是　□否", None)],
            [("成绩评定：", None)],
        ]
        stops = [Cm(4.6), Cm(9.4), Cm(13.2)]
        for cells in rows:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(10)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            for stop in stops:
                p.paragraph_format.tab_stops.add_tab_stop(stop)
            for idx, (label, value) in enumerate(cells):
                if idx:
                    p.add_run("\t")
                run = p.add_run(label)
                _font(run, cjk="楷体", size=10.5)
                if value:
                    run2 = p.add_run(value)
                    _font(run2, cjk="楷体", size=10.5)
                    run2.font.underline = True
                elif label.endswith("成绩评定："):
                    run3 = p.add_run("　　　　　　　　")
                    _font(run3, cjk="楷体", size=10.5)
                    run3.font.underline = True

    # ---------- 正文 ----------
    def _fmt(self, p, indent: bool = True, before: float = 0, after: float = 0):
        """正文段落格式：固定 20 磅行距（模板的 Normal 不许动，只能逐段落设）。"""
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(LINE_PT)
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.first_line_indent = Pt(BODY_PT * 2) if indent else Pt(0)

    def heading(self, level: int, text: str, page_break: bool = False):
        p = self.doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        self._fmt(p, indent=False, before=12 if level <= 1 else (6 if level == 2 else 3), after=6 if level <= 1 else 3)
        if page_break:  # 用“段前分页”而不是插一个空段落，避免内容正好排满时多出空白页
            p.paragraph_format.page_break_before = True
        if level <= 1:
            run = p.add_run(text)
            _font(run, cjk=HEAD_CJK, size=14, bold=True)
        else:
            run = p.add_run(text)
            # 三级及以下用宋体加粗，和黑体的二级标题区分开
            _font(run, cjk=HEAD_CJK if level == 2 else BODY_CJK, size=12, bold=True)

    def para(self, text: str):
        p = self.doc.add_paragraph()
        self._fmt(p, indent=True)
        self._inline(p, text)

    def _inline(self, p, text: str):
        for part in _INLINE.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                run = p.add_run(part[2:-2])
                _font(run, bold=True)
            elif part.startswith("$") and part.endswith("$") and len(part) > 2:
                omml = self.inline_omml.get(part[1:-1].strip())
                if omml is None:
                    run = p.add_run(part[1:-1])
                    _font(run, italic=True)
                else:
                    self._append_omml(p, omml)
            else:
                run = p.add_run(part)
                _font(run)

    def _append_omml(self, p, omml: str):
        try:
            p._p.append(parse_xml(omml))
        except Exception as exc:  # OMML 坏了就退回文字，别整篇失败
            run = p.add_run(f"[公式解析失败: {exc}]")
            _font(run)

    def formula(self, tex: str, omml: str, number: str | None = None):
        p = self.doc.add_paragraph()
        self._fmt(p, indent=False, before=6, after=6)
        if number:
            # 居中制表位放公式、右制表位放编号——这是 Word 里公式编号的标准做法
            p.paragraph_format.tab_stops.add_tab_stop(Cm(self.text_width.cm / 2), WD_TAB_ALIGNMENT.CENTER)
            p.paragraph_format.tab_stops.add_tab_stop(self.text_width, WD_TAB_ALIGNMENT.RIGHT)
            p.add_run("\t")
            self._append_omml(p, omml)
            p.add_run("\t")
            run = p.add_run(f"（{number}）")
            _font(run)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._append_omml(p, omml)

    def caption(self, text: str):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._fmt(p, indent=False, before=6, after=3)
        run = p.add_run(text)
        _font(run, size=CAPTION_PT, bold=True)

    def table(self, data: dict):
        columns = data.get("columns") or []
        rows = data.get("rows") or []
        if not columns:
            raise ValueError(f"表格 {data.get('id') or data.get('caption')} 缺少 columns")
        if data.get("caption"):
            self.caption(data["caption"])
        t = self.doc.add_table(rows=1, cols=len(columns))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = True
        for idx, name in enumerate(columns):
            _cell_text(t.rows[0].cells[idx], name)
        for row in rows:
            cells = t.add_row().cells
            for idx in range(len(columns)):
                _cell_text(cells[idx], row[idx] if idx < len(row) else "")
        _three_line(t)
        _header_bottom_line(t)
        if data.get("note"):
            p = self.doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(3)
            run = p.add_run(str(data["note"]))
            _font(run, size=CAPTION_PT)

    def figure(self, data: dict):
        path = Path(data.get("path") or "")
        if not path.is_absolute():
            path = self.cfg.resolve(path)
        if not path.exists():
            raise FileNotFoundError(f"图不存在：{path}")
        width = float(data.get("width") or self.cfg.fig_width_cm)
        max_h = float(data.get("maxHeightCm") or 15.5)
        try:
            from docx.image.image import Image as _Img

            im = _Img.from_file(str(path))
            ratio = im.px_height / im.px_width if im.px_width else 1.0
            if width * ratio > max_h:  # 竖长的图按页高反算宽度，避免图被页顶/页底裁掉
                width = max_h / ratio
        except Exception:
            pass
        # 图与图题放进 1×1 无框表格：整行 cantSplit，Word 只会整块挪页，不会拦腰截断
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        _no_borders(t)
        trpr = t.rows[0]._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        cell = t.cell(0, 0)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        # 图所在段落必须是单倍行距：正文的“固定 20 磅”会把高于 20 磅的图片顶出页面
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.add_run().add_picture(str(path), width=Cm(width))
        if data.get("caption"):
            cap = cell.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.first_line_indent = Pt(0)
            cap.paragraph_format.space_before = Pt(4)
            cap.paragraph_format.space_after = Pt(3)
            run = cap.add_run(data["caption"])
            _font(run, size=CAPTION_PT, bold=True)

    def pagebreak(self):
        p = self.doc.add_paragraph()
        run = p.add_run()
        run.add_break(WD_BREAK.PAGE)

    def save(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(path))
        return path


def inline_texts(blocks: list[tuple]) -> list[str]:
    """收集段落里的 $...$ 行内公式原文（去重、保序）。"""
    seen: list[str] = []
    for kind, *rest in blocks:
        if kind != "para" or not rest:
            continue
        for part in _INLINE.split(rest[0]):
            if part.startswith("$") and part.endswith("$") and len(part) > 2:
                tex = part[1:-1].strip()
                if tex not in seen:
                    seen.append(tex)
    return seen
