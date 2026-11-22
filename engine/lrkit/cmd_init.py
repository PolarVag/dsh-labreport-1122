"""init：在实验目录下建 `_lr` 工程，抽取指导书/书写要求/模板，铺好目录骨架。"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

from . import project
from .config import Config
from .util import read_text_auto, today_iso, write_json

GUIDE_KEYS = ("指导书", "实验指导", "guide")
REQ_KEYS = ("书写要求", "要求", "requirement")
TPL_KEYS = ("报告头", "模板", "template")
SAMPLE_DIRS = ("样例", "sample", "samples", "范文", "示例")
SAMPLE_SUFFIXES = (".docx", ".pdf", ".md", ".txt")
DATA_SUFFIXES = (".xlsx", ".xls", ".csv", ".tsv", ".txt", ".json", ".dat")
INFO_KEYS = ("学号", "姓名", "学生", "信息", "PROMPT", "注意", "说明", "Extra", "Information", "info")
_INFO_LINE = re.compile(r"^\s*([^：:]{2,14})[：:]\s*(.*?)\s*$")
_DATE_CN = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")

# 学生信息文件里的键 → report.yaml 路径（不假设任何具体姓名，全部从文件读）
_INFO_FIELDS = {
    "实验名称": ("experiment", "name"),
    "实验项目": ("experiment", "name"),
    "实验题目": ("experiment", "name"),
    "指导教师": ("experiment", "instructor"),
    "任课教师": ("experiment", "instructor"),
    "老师": ("experiment", "instructor"),
    "实验地点": ("experiment", "location"),
    "地点": ("experiment", "location"),
    "上课地点": ("experiment", "location"),
    "实验日期": ("experiment", "date"),
    "实验时间": ("experiment", "date"),
    "日期": ("experiment", "date"),
    "姓名": ("student", "name"),
    "学生姓名": ("student", "name"),
    "同学姓名": ("student", "name"),
    "学号": ("student", "id"),
    "分班分组及座号": ("student", "classGroupSeat"),
    "分班分组及座位号": ("student", "classGroupSeat"),
    "座位号": ("student", "classGroupSeat"),
    "分班分组": ("student", "classGroupSeat"),
}


def _norm_date(text: str) -> str:
    m = _DATE_CN.search(text or "")
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text or "")
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def _pick(folder: Path, keys: tuple[str, ...], suffix: str) -> Path | None:
    best: list[tuple[int, Path]] = []
    for p in folder.glob(f"*{suffix}"):
        if p.name.startswith("~$") or not p.is_file():
            continue
        score = sum(1 for k in keys if k in p.name)
        if score:
            best.append((score, p))
    if not best:
        cands = [p for p in folder.glob(f"*{suffix}") if p.is_file() and not p.name.startswith("~$")]
        return max(cands, key=lambda p: p.stat().st_size) if len(cands) == 1 else (cands[0] if cands else None)
    best.sort(key=lambda t: (-t[0], -t[1].stat().st_size))
    return best[0][1]


def _pick_template(src_dir: Path, req: dict) -> tuple[Path | None, Path | None]:
    """挑报告模板：优先 .docx；只有旧版 .doc 时把它交回去，让插件侧先转换。"""
    if req.get("template"):
        p = Path(req["template"])
        if p.exists():
            return (p, None) if p.suffix.lower() == ".docx" else (None, p)
    tpl = _pick(src_dir, TPL_KEYS, ".docx")
    if tpl and tpl.exists():
        return tpl, None
    legacy = _pick(src_dir, TPL_KEYS, ".doc")
    if legacy and legacy.exists():
        return None, legacy
    return None, None


def docx_to_text(docx_path: Path, limit: int = 60_000) -> str:
    """把样例 .docx 抽成可读文本（段落 + 表格），供写作时参照结构与口吻。"""
    from docx import Document

    doc = Document(str(docx_path))
    out: list[str] = []
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if text:
            style = (p.style.name or "").lower()
            if style.startswith("heading") or style.startswith("title"):
                out.append(f"# {text}")
            else:
                out.append(text)
    for t in doc.tables:
        for row in t.rows:
            cells = [(c.text or "").strip().replace("\n", " ") for c in row.cells]
            out.append("| " + " | ".join(cells) + " |")
    text = "\n\n".join(out)
    return text[:limit]


def _sample_sources(src_dir: Path, req: dict) -> list[Path]:
    """样例来源：显式 sample 参数优先，否则扫实验目录下的「样例/sample」子目录。"""
    explicit = req.get("sample")
    if isinstance(explicit, (str, Path)):
        explicit = [explicit]
    paths: list[Path] = []
    if explicit:
        paths = [Path(p) for p in explicit]
    else:
        for name in SAMPLE_DIRS:
            folder = src_dir / name
            if folder.is_dir():
                paths += sorted(p for p in folder.iterdir() if p.is_file())
    out: list[Path] = []
    for p in paths:
        if p.is_file() and p.suffix.lower() in SAMPLE_SUFFIXES and p.stat().st_size < 30_000_000:
            if not p.name.startswith("~$"):
                out.append(p)
    return out[:6]


def _ingest_samples(cfg: Config, sources: list[Path]) -> list[dict]:
    """样例复制进 book/sample/ 并顺带抽文本，写报告时直接读 .md 即可。"""
    made: list[dict] = []
    for src in sources:
        dst = _copy(src, cfg.book / "sample" / src.name)
        item = {"name": src.name, "path": str(dst), "kind": src.suffix.lower().lstrip("."), "chars": 0, "text": None}
        try:
            suffix = src.suffix.lower()
            if suffix == ".docx":
                text = docx_to_text(dst)
                md = dst.with_suffix(".md")
                md.write_text(text + "\n", encoding="utf-8")
                item.update(chars=len(text), text=str(md))
            elif suffix == ".pdf":
                _pages, text = pdf_to_text(dst, dst.with_suffix(".md"))
                item.update(chars=len(text or ""), text=str(dst.with_suffix(".md")))
            elif suffix in (".md", ".txt"):
                text, _enc = read_text_auto(dst)
                item.update(chars=len(text or ""), text=str(dst))
        except Exception as exc:  # 样例解析失败不该挡住 init
            item["error"] = str(exc)
        made.append(item)
    return made


def _unblock(path: Path) -> None:
    """去掉 Windows 的 Zone.Identifier（带 MOTW 的 docx 会被 Word 丢进受保护视图，自动化会卡死）。"""
    try:
        (path.parent / f"{path.name}:Zone.Identifier").unlink()
    except OSError:
        pass


def _copy(src: Path, dst: Path) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if src.resolve() != dst.resolve():
        shutil.copyfile(src, dst)
    _unblock(dst)
    return dst


def pdf_to_text(pdf: Path, out_md: Path, max_pages: int = 0) -> tuple[int, str]:
    """PDF → 带页码标记的 markdown 文本。"""
    import pymupdf

    chunks: list[str] = []
    with pymupdf.open(str(pdf)) as doc:
        total = doc.page_count
        limit = min(total, max_pages) if max_pages else total
        for i in range(limit):
            text = doc[i].get_text("text").strip()
            if text:
                chunks.append(f"\n\n<!-- 第 {i + 1} 页 -->\n\n{text}")
    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_md.write_text("".join(chunks).strip() + "\n", encoding="utf-8")
    return total, "".join(chunks)


def _info_files(folder: Path) -> list[Path]:
    out: list[Path] = []
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.suffix.lower() not in (".txt", ".md"):
            continue
        if p.name.startswith("~$") or p.stat().st_size > 200_000:
            continue
        if any(k in p.name for k in INFO_KEYS):
            out.append(p)
    return out


def meta_from_notes(folder: Path) -> tuple[dict, list[str], str]:
    """从实验目录里的学生信息/说明文本（学号.txt、PROMPT.txt 等）读表头信息。

    只认「键：值」行，编码自动探测（UTF-8/BOM/GB18030）；读不到就留空，
    绝不内置任何姓名学号——换人换班只要改文件。
    """
    meta: dict = {"experiment": {}, "student": {}}
    filled: list[str] = []
    used = ""
    for path in _info_files(folder):
        text, _enc = read_text_auto(path)
        hit = 0
        for line in (text or "").splitlines():
            # 一行里常塞多组「键：值」，用中文/英文分号切开再逐段解析
            for seg in re.split(r"[；;]", line):
                m = _INFO_LINE.match(seg.strip())
                if not m:
                    continue
                key, value = m.group(1).strip(), m.group(2).strip()
                if not value or len(value) > 60:
                    continue
                if key in ("调课/补课", "调课补课", "调课"):
                    meta["experiment"]["makeUp"] = value not in ("否", "无", "false", "False", "0")
                    hit += 1
                    continue
                target = _INFO_FIELDS.get(key)
                if not target:
                    continue
                group, field = target
                if field == "date":
                    iso = _norm_date(value)
                    if not iso:
                        continue
                    value = iso
                if meta[group].get(field):
                    continue
                meta[group][field] = value
                filled.append(f"{group}.{field}")
                hit += 1
        if hit and not used:
            used = path.name
    meta["experiment"].setdefault("short", meta["experiment"].get("name", ""))
    return meta, filled, used


def _merge_meta(cur: dict, new: dict) -> dict:
    out = {k: (dict(v) if isinstance(v, dict) else v) for k, v in (cur or {}).items()}
    for group, fields in (new or {}).items():
        if isinstance(fields, dict):
            slot = out.setdefault(group, {})
            if not isinstance(slot, dict):
                slot = out[group] = {}
            for k, v in fields.items():
                if v not in (None, "", False):
                    slot[k] = v
        elif fields not in (None, ""):
            out[group] = fields
    return out


def run(cfg: Config, req: dict) -> dict:
    src_dir = cfg.dir
    if not src_dir.exists():
        raise FileNotFoundError(f"实验目录不存在：{src_dir}")
    project.prepare(cfg)
    warnings: list[str] = []

    report = project.load_report(cfg) or project.default_report(cfg)
    exp0 = report.get("experiment") or {}
    stu0 = report.get("student") or {}
    existing = {
        "experiment": exp0.get("name") or "",
        "short": exp0.get("short") or "",
        "date": exp0.get("date") or "",
        "place": exp0.get("location") or "",
        "teacher": exp0.get("instructor") or "",
        "makeup": bool(exp0.get("makeUp")),
        "student": stu0.get("name") or "",
        "student_id": stu0.get("id") or "",
        "group": stu0.get("classGroupSeat") or "",
        "grade": stu0.get("grade") or "",
    }
    existing.update({k: v for k, v in (report.get("meta") or {}).items() if v not in (None, "")})
    # 默认日期是"今天"，不代表真实实验日期：检测到的日期优先把它换掉
    date_is_today = existing.get("date") == today_iso()

    detected, meta_filled, meta_source = meta_from_notes(src_dir)
    d_exp = detected.get("experiment") or {}
    d_stu = detected.get("student") or {}
    det_flat = {
        "experiment": d_exp.get("name"),
        "short": d_exp.get("short"),
        "date": d_exp.get("date"),
        "place": d_exp.get("location"),
        "teacher": d_exp.get("instructor"),
        "makeup": d_exp.get("makeUp"),
        "student": d_stu.get("name"),
        "student_id": d_stu.get("id"),
        "group": d_stu.get("classGroupSeat"),
    }
    given = dict(req.get("meta") or {})
    given.setdefault("group", given.get("classGroupSeat", ""))
    given_flat = {k: given.get(k) for k in existing if given.get(k) not in (None, "")}

    keep = dict(existing)
    if date_is_today and det_flat.get("date"):
        keep.pop("date", None)
    meta: dict = {}
    for src in (det_flat, keep, given_flat):
        meta.update({k: v for k, v in src.items() if v not in (None, "")})
    meta.setdefault("experiment", src_dir.name)
    meta.setdefault("short", meta["experiment"])
    if not meta.get("date"):
        meta["date"] = today_iso()

    report["experiment"] = {
        **exp0,
        "name": meta["experiment"],
        "short": meta.get("short") or meta["experiment"],
        "date": meta.get("date", ""),
        "location": meta.get("place", ""),
        "instructor": meta.get("teacher", ""),
        "makeUp": bool(meta.get("makeup")),
    }
    report["student"] = {
        **stu0,
        "name": meta.get("student", ""),
        "id": meta.get("student_id", ""),
        "classGroupSeat": meta.get("group", ""),
        "grade": meta.get("grade", ""),
    }
    report["meta"] = {k: v for k, v in meta.items() if v not in (None, "")}
    meta = report["meta"]

    sources: dict[str, str] = {}
    tpl_src, legacy_doc = _pick_template(src_dir, req)
    if tpl_src:
        dst = _copy(tpl_src, cfg.book / "template.docx")
        report["template"] = str(dst)
        sources["template"] = str(tpl_src)
    elif legacy_doc:
        sources["template"] = str(legacy_doc)
        warnings.append(
            f"报告模板是旧版 {legacy_doc.suffix}（{legacy_doc.name}）：需要先转成 .docx，"
            "插件会自动转换；手动跑命令行时请先用 Word 另存为 .docx"
        )
    else:
        warnings.append("没找到报告模板（实验报告报告头*.docx），将用内置版式生成表头页")

    guide_src = Path(req["guide"]) if req.get("guide") else _pick(src_dir, GUIDE_KEYS, ".pdf")
    if guide_src and not req.get("guide") and "guide" not in guide_src.name.lower():
        # 目录里只有一个 PDF 时 _pick 会把它当指导书，若它其实是书写要求就放过
        if any(k in guide_src.name for k in REQ_KEYS) and not any(k in guide_src.name for k in GUIDE_KEYS):
            guide_src = None
    if guide_src and guide_src.exists():
        dst = _copy(guide_src, cfg.book / "guide.pdf")
        pages, _ = pdf_to_text(dst, cfg.book / "guide.md")
        sources["guide"] = str(guide_src)
        report.setdefault("sources", {})["guidePages"] = pages
    else:
        warnings.append("没找到实验指导书 PDF")

    req_src = Path(req["requirements"]) if req.get("requirements") else _pick(src_dir, REQ_KEYS, ".pdf")
    rules_text = ""
    if req_src and req_src.exists():
        dst = _copy(req_src, cfg.book / "requirements.pdf")
        pages, rules_text = pdf_to_text(dst, cfg.book / "requirements.md")
        sources["requirements"] = str(req_src)
        report.setdefault("sources", {})["requirementPages"] = pages
    else:
        warnings.append("没找到书写要求 PDF")

    for extra in req.get("copy") or []:
        p = Path(extra)
        if p.is_file():
            # 数据进 data/（compute.py 默认在那里找），图片等其它素材进 media/
            dst_dir = cfg.data_dir if p.suffix.lower() in DATA_SUFFIXES else cfg.media
            _copy(p, dst_dir / p.name)

    samples = _ingest_samples(cfg, _sample_sources(src_dir, req))
    if samples:
        sources["samples"] = "、".join(str(s["name"]) for s in samples)

    report["sources"] = {**(report.get("sources") or {}), **sources}
    project.save_report(cfg, report)

    created = []
    for sec in report.get("sections") or []:
        path = cfg.resolve(sec["file"])
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("", encoding="utf-8")
            created.append(sec["file"])

    project.write_checklist(cfg, _rules_from(rules_text), meta)

    summary = f"工程已就绪：{cfg.project_dir}" + (f"；新建 {len(created)} 个空小节" if created else "")
    if meta_source:
        summary += f"\n表头信息取自 {meta_source}：" + "、".join(
            f"{k}={meta[k]}" for k in ("experiment", "teacher", "student", "student_id", "group", "date", "place") if meta.get(k)
        )
    missing = [k for k in ("experiment", "teacher", "student", "student_id", "group", "date", "place") if not meta.get(k)]
    if missing:
        summary += "\n还需补的表头项：" + "、".join(missing)
    if samples:
        summary += "\n样例已入库：" + "、".join(s["name"] for s in samples) + "（文本在 book/sample/）"
    if warnings:
        summary += "\n警告：" + "；".join(warnings)
    return {
        "ok": True,
        "summary": summary,
        "data": {
            "project": str(cfg.project_dir),
            "template": report.get("template"),
            "templateNeedsConvert": str(legacy_doc) if legacy_doc else None,
            "guide": sources.get("guide"),
            "requirements": sources.get("requirements"),
            "samples": samples,
            "guideChars": len((cfg.book / "guide.md").read_text(encoding="utf-8")) if (cfg.book / "guide.md").exists() else 0,
            "requirementsChars": len(rules_text or ""),
            "sectionsCreated": created,
            "meta": meta,
            "metaSource": meta_source,
            "metaFilled": meta_filled,
            "metaMissing": missing,
            "warnings": warnings,
            "nextStep": "先读 book/guide.md 与 book/requirements.md，把十节大纲与公式/图表清单写进 sections/*.md，再跑 labreport_compute",
        },
    }


def _rules_from(text: str) -> list[dict]:
    """把书写要求正文切成条目，供 check 用（保持原文，不做语义改写）。"""
    rules: list[dict] = []
    for raw in (text or "").splitlines():
        line = raw.strip().lstrip("·•*-").strip()
        if len(line) < 8:
            continue
        if any(k in line for k in ("要求", "必须", "不得", "禁止", "应", "需要", "表题", "图题", "公式", "页", "提交")):
            rules.append({"level": "must", "text": line[:200]})
    return rules[:80]
