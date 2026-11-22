"""工程脚手架与 report.yaml 读写（唯一事实源）。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from .config import Config
from .util import today_iso, write_json

# 九节结构：把「获取数据表格」与「数据处理及结果」合成一节，
# 好让每个分实验的表格后面紧跟它自己的处理（逐差/拟合/图/中间结果），
# 而不是"所有数据先摆一遍、再统一处理一遍"。
DEFAULT_SECTIONS: list[dict[str, Any]] = [
    {"id": "purpose", "title": "目的及要求", "file": "report-making/sections/01_purpose.md", "maxPages": None},
    {"id": "apparatus", "title": "仪器用具", "file": "report-making/sections/02_apparatus.md", "maxPages": None},
    {"id": "principle", "title": "实验原理", "file": "report-making/sections/03_principle.md", "maxPages": 2},
    {"id": "procedure", "title": "实验内容", "file": "report-making/sections/04_procedure.md", "maxPages": None},
    {"id": "data", "title": "获取数据表格与数据处理", "file": "report-making/sections/05_data.md", "maxPages": None,
     "note": "每个分实验：先数据表格，紧跟该分实验的处理与结果；最后汇总表与误差分析"},
    {"id": "discussion", "title": "问题讨论", "file": "report-making/sections/06_discussion.md", "maxPages": None},
    {"id": "summary", "title": "总结", "file": "report-making/sections/07_summary.md", "maxPages": None},
    {"id": "questions", "title": "思考题", "file": "report-making/sections/08_questions.md", "maxPages": None},
    {"id": "handout", "title": "预习报告及原始数据", "file": "report-making/sections/09_handout.md", "mode": "scan", "maxPages": None},
]

# 旧版十节工程的兼容：这两节存在时要提示合并
LEGACY_PAIR = ("rawdata", "processing")

DEFAULT_STYLES = {
    "body": {"font": "宋体", "sizePt": 12, "lineSpacing": 1.5, "firstLineIndentChars": 0, "align": "both"},
    "heading1": {"font": "黑体", "sizePt": 14, "bold": True},
    "heading2": {"font": "黑体", "sizePt": 12, "bold": True},
    "caption": {"font": "宋体", "sizePt": 10.5, "align": "center"},
    "table": {"font": "宋体", "sizePt": 10.5, "align": "center"},
}


def ensure_dirs(cfg: Config) -> None:
    for d in (cfg.book, cfg.media, cfg.data_dir, cfg.sections, cfg.calc, cfg.figures, cfg.logs, cfg.tmp, cfg.out):
        d.mkdir(parents=True, exist_ok=True)


def migrate_layout(cfg: Config) -> list[str]:
    """把旧布局（中间产物直接摊在 _lr/ 下、成品在 _lr/out/）搬进 report-making/ 与 output/。

    只搬真存在且目标还不在的东西，幂等；返回搬了哪些。不删也不覆盖已有文件。
    """
    import shutil

    moved: list[str] = []
    cfg.making.mkdir(parents=True, exist_ok=True)
    for name in ("book", "media", "data", "sections", "calc", "figures", "logs"):
        old = cfg.project_dir / name
        new = cfg.making / name
        if old.is_dir() and not new.exists():
            shutil.move(str(old), str(new))
            moved.append(name)
    old_out = cfg.project_dir / "out"
    if old_out.is_dir():
        cfg.out.mkdir(parents=True, exist_ok=True)
        for item in old_out.iterdir():
            target = cfg.out / item.name
            if target.exists():
                continue
            shutil.move(str(item), str(target))
        try:
            old_out.rmdir()
        except OSError:
            pass
        moved.append("out")
    return moved


def prepare(cfg: Config) -> list[str]:
    """每个子命令开头统一调用：搬旧布局、补齐目录。"""
    moved = migrate_layout(cfg)
    ensure_dirs(cfg)
    return moved


def _fix_rel(rel: str) -> str:
    """旧工程里写的是 `sections/xx.md`，新布局住在 report-making/ 下。"""
    if not rel:
        return rel
    p = Path(rel)
    if p.is_absolute() or not p.parts:
        return rel
    if p.parts[0] == "report-making":
        return rel
    if p.parts[0] in ("sections", "calc", "figures", "book", "media", "data", "logs"):
        return "report-making/" + p.as_posix()
    return rel


def default_report(cfg: Config) -> dict[str, Any]:
    name = cfg.dir.name
    return {
        "experiment": {
            "name": name,
            "short": name,
            "date": today_iso(),
            "location": "",
            "instructor": "",
            "makeUp": False,
        },
        "student": {"name": "", "id": "", "classGroupSeat": "", "grade": ""},
        "template": {"path": "", "convertedDocx": "", "headerParagraphs": "auto"},
        "sections": [dict(s) for s in DEFAULT_SECTIONS],
        "tables": [],
        "figures": [],
        "formulas": [],
        "data": {"files": [], "normalized": [], "notes": []},
        "compute": {"script": "report-making/calc/compute.py", "outputs": ["report-making/calc/tables.json", "report-making/calc/results.json"]},
        "build": {
            "output": "output/report.docx",
            "pdf": True,
            "page": {"marginCm": 2.0, "footerPageNumber": True},
            "styles": DEFAULT_STYLES,
        },
        "checks": {"req": "report-making/book/req.checklist.json"},
        "sources": {},
        "todo": [],
    }


def migrate_legacy_sections(data: dict[str, Any]) -> None:
    """旧版十节工程：把「获取数据表格」+「数据处理及结果」并成一节。

    只改 report.yaml 的节结构，不动磁盘上的 md 文件——被并入的那份文件记在
    ``legacyProcessing`` 里，由 build/check 提示作者把内容手工搬进合并后的正文。
    合并后每节仍然保留原来的 title 以外的字段（mode 等）。
    """
    sections = data.get("sections") or []
    ids = [s.get("id") for s in sections]
    if not all(k in ids for k in LEGACY_PAIR):
        return
    merged: list[dict[str, Any]] = []
    taken = False
    for sec in sections:
        sid = sec.get("id")
        if sid == "rawdata":
            item = dict(sec)
            item["id"] = "data"
            item["title"] = "获取数据表格与数据处理"
            item["note"] = "每个分实验：先数据表格，紧跟该分实验的处理与结果"
            merged.append(item)
            taken = True
        elif sid == "processing":
            if not taken:
                item = dict(sec)
                item["id"] = "data"
                item["title"] = "获取数据表格与数据处理"
                merged.append(item)
                taken = True
            else:
                data["legacyProcessing"] = sec.get("file") or "sections/06_processing.md"
        else:
            merged.append(sec)
    data["sections"] = merged


def load_report(cfg: Config) -> dict[str, Any] | None:
    p = cfg.report_yaml
    if not p.exists():
        return None
    data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    return normalize(data, cfg)


def normalize(data: dict[str, Any], cfg: Config) -> dict[str, Any]:
    base = default_report(cfg)
    for key, value in base.items():
        if key not in data or data[key] in (None, ""):
            data[key] = value
    # sections 允许用户增删，但缺字段要补齐
    sections = []
    for s in data.get("sections") or []:
        item = {"id": "", "title": "", "file": "", "maxPages": None}
        item.update({k: v for k, v in s.items() if v is not None})
        item["file"] = _fix_rel(item.get("file") or "")
        if item.get("title"):
            sections.append(item)
    data["sections"] = sections or [dict(s) for s in DEFAULT_SECTIONS]
    migrate_legacy_sections(data)
    if not isinstance(data.get("tables"), list):
        data["tables"] = []
    if not isinstance(data.get("figures"), list):
        data["figures"] = []
    if not isinstance(data.get("formulas"), list):
        data["formulas"] = []
    if not isinstance(data.get("data"), dict):
        data["data"] = {"files": [], "normalized": [], "notes": []}
    data["data"].setdefault("files", [])
    data["data"].setdefault("normalized", [])
    data["data"].setdefault("notes", [])
    build = data.get("build") or {}
    build.setdefault("output", "out/report.docx")
    build.setdefault("pdf", True)
    build.setdefault("page", {"marginCm": 2.0, "footerPageNumber": True})
    styles = build.get("styles") or {}
    for key, value in DEFAULT_STYLES.items():
        styles.setdefault(key, value)
    build["styles"] = styles
    data["build"] = build
    return data


def save_report(cfg: Config, data: dict[str, Any]) -> Path:
    cfg.project_dir.mkdir(parents=True, exist_ok=True)
    text = yaml.safe_dump(data, allow_unicode=True, sort_keys=False, width=120)
    cfg.report_yaml.write_text(text, encoding="utf-8", newline="\n")
    return cfg.report_yaml


def section_path(cfg: Config, section: dict[str, Any]) -> Path:
    raw = section.get("file") or f"report-making/sections/{section.get('id', 'x')}.md"
    return cfg.resolve(raw)


def checklist_path(cfg: Config) -> Path:
    return cfg.book / "req.checklist.json"


def write_checklist(cfg: Config, rules: list[dict[str, Any]], meta: dict[str, Any] | None = None) -> Path:
    return write_json(checklist_path(cfg), {"meta": meta or {}, "rules": rules})
