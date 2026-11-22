"""status：工程进度。"""
from __future__ import annotations

from pathlib import Path

from . import project
from .config import Config
from .util import char_count, read_text_auto


def _scan_sections(cfg: Config, report: dict) -> list[dict]:
    out = []
    for s in report.get("sections") or []:
        p = project.section_path(cfg, s)
        item = {"id": s.get("id"), "title": s.get("title"), "file": str(p), "exists": p.exists(), "chars": 0}
        if p.exists():
            text, _enc = read_text_auto(p)
            item["chars"] = char_count(text)
            item["state"] = "ok" if item["chars"] >= 20 else "empty"
        else:
            item["state"] = "missing"
        out.append(item)
    return out


def _last_build(cfg: Config) -> dict | None:
    if not cfg.out.exists():
        return None
    files = sorted(cfg.out.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return None
    docx = files[0]
    pdf = docx.with_suffix(".pdf")
    stat = docx.stat()
    return {
        "docx": str(docx),
        "pdf": str(pdf) if pdf.exists() else None,
        "mtime": int(stat.st_mtime),
        "size": stat.st_size,
    }


def run(cfg: Config, request: dict) -> dict:
    cfg.project_dir.mkdir(parents=True, exist_ok=True)
    report = project.load_report(cfg)
    if report is None:
        summary = (
            "工程还没初始化：缺少 report.yaml。\n"
            f"目录：{cfg.project_dir}\n"
            "下一步：调用 labreport_init 建工程（会自动抽取指导书/书写要求/样例）。"
        )
        return {"ok": True, "summary": summary, "data": {"initialized": False, "project": str(cfg.project_dir), "nextStep": "labreport_init"}}

    sections = _scan_sections(cfg, report)
    tables = len(report.get("tables") or [])
    figures = len(report.get("figures") or [])
    formulas = len(report.get("formulas") or [])
    last = _last_build(cfg)
    missing = [s["title"] for s in sections if s["state"] != "ok"]
    written = [s for s in sections if s["state"] == "ok"]

    exp = report.get("experiment") or {}
    if missing:
        next_step = "写 sections/ 里缺的正文（先写 1 目的/2 仪器/3 原理，确认后再写数据处理）"
    elif tables == 0 and figures == 0:
        next_step = "写 calc/compute.py 并跑 labreport_compute（表格与图来自计算，不能抄照片）"
    elif not last:
        next_step = "跑 labreport_build 组装 docx"
    else:
        next_step = "跑 labreport_check 自检，按 fail/warn 修改后重新 build"

    lines = [
        f"工程：{cfg.project_dir}",
        f"实验：{exp.get('name') or '（未命名）'}    日期：{exp.get('date') or '-'}    指导教师：{exp.get('instructor') or '-'}",
        f"章节：{len(written)}/{len(sections)} 已写    表格 {tables}    图 {figures}    公式 {formulas}",
    ]
    if last:
        lines.append(f"上次构建：{Path(last['docx']).name}（{last['size']} B）" + (" + PDF" if last["pdf"] else "（无 PDF）"))
    else:
        lines.append("上次构建：还没有")
    if missing:
        lines.append("缺：" + "、".join(missing))
    lines.append(f"下一步：{next_step}")

    return {
        "ok": True,
        "summary": "\n".join(lines),
        "data": {
            "initialized": True,
            "project": str(cfg.project_dir),
            "experiment": exp,
            "sections": [{"id": s["id"], "title": s["title"], "state": s["state"], "chars": s["chars"]} for s in sections],
            "tables": tables,
            "figures": figures,
            "formulas": formulas,
            "lastBuild": last,
            "missing": missing,
            "nextStep": next_step,
        },
    }
