"""build：md-lite → 排好版的 docx（可选再导 PDF）。"""
from __future__ import annotations

import datetime as _dt
import re
from pathlib import Path

from . import docxbuild, mdlite, pdfout, project
from .config import Config
from .formula import build_formulas
from .util import char_count, read_text_auto, safe_name, stamp

CN_NUM = "一二三四五六七八九十"


def heading_text(index: int, title: str) -> str:
    num = CN_NUM[index] if index < len(CN_NUM) else str(index + 1)
    return f"{num}、{title}"


_NUM_PREFIX = re.compile(r"^\s*(?:\d+(?:\.\d+)*)[.、．)）]?\s*")


def renumber(blocks: list[tuple], chapter: int) -> list[tuple]:
    """把二级/三级小标题编号统一成"章号.序号"，避免写成"1.1"却出现在第二章里。

    层级各自计数：``## …`` → ``5.1``、``5.2``；其下的 ``### …`` → ``5.1.1``、``5.1.2``。
    """
    counters: dict[int, int] = {}
    out: list[tuple] = []
    for b in blocks:
        if b[0] == "heading" and b[1] >= 2:
            level = b[1]
            for i in range(2, level + 1):
                counters.setdefault(i, 0)
            if level >= 3 and counters.get(level - 1, 0) == 0:
                # 三级标题前面没有二级标题时，补一个 1，避免出现「8.0.1」
                counters[level - 1] = 1
            counters[level] = counters.get(level, 0) + 1
            for deeper in [k for k in counters if k > level]:
                counters[deeper] = 0
            parts = [str(chapter)] + [str(counters[i]) for i in range(2, level + 1)]
            num = ".".join(parts)
            text = _NUM_PREFIX.sub("", b[2] or "")
            out.append(("heading", level, f"{num} {text}" if text else num))
        else:
            out.append(b)
    return out


def _collect(parsed: list[tuple]) -> tuple[list[dict], dict[str, str], int, int, int]:
    display: list[dict] = []
    inline_tex: list[str] = []
    tables = figures = 0
    for _sec, blocks in parsed:
        for b in blocks:
            if b[0] == "table":
                tables += 1
            elif b[0] == "figure":
                figures += 1
            elif b[0] == "formula":
                item = b[1]
                item["_id"] = f"f{len(display) + 1:02d}"
                display.append({"id": item["_id"], "tex": item["tex"], "display": True})
        for tex in docxbuild.inline_texts(blocks):
            if tex not in inline_tex:
                inline_tex.append(tex)
    return display, {tex: f"i{k + 1:03d}" for k, tex in enumerate(inline_tex)}, tables, figures, len(display)


def run(cfg: Config, req: dict) -> dict:
    project.prepare(cfg)
    report = project.load_report(cfg) or project.default_report(cfg)
    meta = dict(report.get("meta") or {})
    meta.setdefault("date", report.get("build", {}).get("date") or _dt.date.today().strftime("%Y年%m月%d日"))

    parsed: list[tuple] = []
    warnings: list[str] = []
    empty: list[str] = []
    for sec in report.get("sections") or []:
        rel = sec.get("file") or ""
        path = cfg.resolve(rel) if rel else None
        if not path or not path.exists():
            warnings.append(f"缺正文：{rel or sec.get('id')}")
            empty.append(sec.get("title") or rel or str(sec.get("id")))
            continue
        text, _enc = read_text_auto(path)
        try:
            blocks = mdlite.expand(mdlite.parse_blocks(text), cfg.calc)
        except (FileNotFoundError, ValueError) as exc:
            raise RuntimeError(f"{rel}：{exc}") from exc
        if blocks:
            parsed.append((sec, blocks))
        else:
            empty.append(sec.get("title") or rel)

    if req.get("strict") and empty:
        raise RuntimeError("严格模式：以下小节还没有内容 —— " + "、".join(empty))

    if not parsed:
        raise RuntimeError("没有任何正文可排版：先在 sections/*.md 里写内容（或检查 report.yaml 的 sections.file）")

    display, inline_ids, n_tables, n_figures, n_formulas = _collect(parsed)

    converted = build_formulas(
        cfg,
        display + [{"id": tid, "tex": tex, "display": False} for tex, tid in inline_ids.items()],
    )
    by_id = {c["id"]: c["omml"] for c in converted}
    inline_map = {tex: by_id[tid] for tex, tid in inline_ids.items()}
    formula_map = {item["id"]: by_id[item["id"]] for item in display}

    builder = docxbuild.Builder(cfg, report)
    builder.inline_omml = inline_map
    builder.setup_page()
    # 表头严格用模板原格式：原地填空、不分页，正文从第一页接着写
    if builder.fill_header(meta):
        builder.title_line(meta)
    else:
        builder.front_page(meta)
    if report.get("build", {}).get("footerPageNumber", True):
        builder.footer_page_number()

    for index, (sec, blocks) in enumerate(parsed):
        builder.heading(
            1,
            heading_text(index, sec.get("title") or sec.get("id") or "正文"),
            page_break=(sec.get("mode") == "scan"),
        )
        for b in renumber(blocks, index + 1):
            kind = b[0]
            if kind == "heading":
                builder.heading(b[1], b[2])
            elif kind == "para":
                builder.para(b[1])
            elif kind == "table":
                builder.table(b[1])
            elif kind == "figure":
                builder.figure(b[1])
            elif kind == "formula":
                builder.formula(b[1]["tex"], formula_map[b[1]["_id"]], b[1].get("number"))
            elif kind == "pagebreak":
                builder.pagebreak()

    base = report.get("build", {}).get("outputName") or "{student}-{experiment}"
    try:
        base_name = base.format(**{k: meta.get(k, "") for k in ("student", "experiment", "student_id", "group")})
    except (KeyError, IndexError):
        base_name = base
    out_dir = cfg.out
    if req.get("docxOut"):
        docx_path = Path(str(req["docxOut"]))
        if not docx_path.is_absolute():
            docx_path = cfg.project_dir / docx_path
        if docx_path.suffix.lower() != ".docx":
            raise ValueError(f"docxOut 必须是 .docx 路径：{docx_path}")
        out_dir = docx_path.parent
        name = docx_path.stem
        docx_path = builder.save(docx_path)
    else:
        name = f"{safe_name(base_name) or 'report'}-{stamp()}"
        docx_path = builder.save(out_dir / f"{name}.docx")

    pdf_path = None
    pages = 0
    pdf_script = None
    pdf_target = None
    if req.get("pdf", True):
        pdf_script, pdf_target = pdfout.prepare_script(cfg, docx_path, out_dir / f"{name}.pdf")

    total_chars = sum(char_count(read_text_auto(cfg.resolve(sec["file"]))[0]) for sec, _ in parsed if sec.get("file"))
    docx_name = Path(docx_path).name
    head = (
        f"已生成 docx：{docx_name}（{len(parsed)} 节 / {n_tables} 表 / {n_figures} 图 / "
        f"{n_formulas + len(inline_ids)} 公式 / {total_chars} 字）"
    )
    if empty:
        head += f"；{len(empty)} 节还是空的：{'、'.join(empty)}"
    if warnings:
        head += f"\n警告 {len(warnings)} 条：" + "；".join(warnings[:4])
    return {
        "ok": True,
        "summary": head,
        "data": {
            "docx": str(docx_path),
            "pdf": str(pdf_path) if pdf_path else None,
            "pdfScript": str(pdf_script) if pdf_script else None,
            "pdfTarget": str(pdf_target) if pdf_target else None,
            "pages": pages,
            "sections": len(parsed),
            "empty": empty,
            "tables": n_tables,
            "figures": n_figures,
            "formulas": n_formulas + len(inline_ids),
            "chars": total_chars,
            "warnings": warnings,
        },
    }
