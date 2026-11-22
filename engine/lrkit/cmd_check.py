"""check：按《基础物理实验》书写要求做机械检查（只查能查的，不做审美判断）。"""
from __future__ import annotations

import re
from pathlib import Path

from . import docxbuild, mdlite, project
from .config import Config
from .util import char_count, read_text_auto

BULLET = re.compile(r"^\s*([-*•·▪]|\d+\))\s+")
PLACEHOLDER = ("待填写", "（待", "TODO", "todo", "XXX")
PRINCIPLE_LIMIT = 1900  # 2 页正文的保守字数
RAW_PHOTO = ("数据记录", "原始数据", "照片", "扫描", "手写")


def _issue(level: str, where: str, message: str) -> dict:
    return {"level": level, "where": where, "message": message}


def run(cfg: Config, req: dict) -> dict:
    issues: list[dict] = []
    report = project.load_report(cfg)
    if not report:
        return {
            "ready": False,
            "issues": [_issue("error", "report.yaml", "工程还没初始化：先跑 init")],
            "summary": "工程未初始化",
        }

    meta = report.get("meta") or {}
    for key, label in (
        ("experiment", "实验名称"),
        ("student", "姓名"),
        ("student_id", "学号"),
        ("teacher", "指导教师"),
        ("date", "实验日期"),
        ("place", "实验地点"),
    ):
        if not meta.get(key):
            issues.append(_issue("warn", "report.yaml", f"表头缺少「{label}」（meta.{key}）"))

    stats = {"sections": 0, "chars": 0, "tables": 0, "figures": 0, "formulas": 0}
    for sec in report.get("sections") or []:
        rel = sec.get("file") or ""
        path = cfg.resolve(rel)
        where = sec.get("title") or sec.get("id") or rel
        if not path.exists():
            issues.append(_issue("error", where, f"缺正文文件 {rel}"))
            continue
        text, _enc = read_text_auto(path)
        body = text.strip()
        stats["sections"] += 1
        stats["chars"] += char_count(body)
        if not body:
            issues.append(_issue("error", where, "这一节还是空的"))
            continue
        if any(p in body for p in PLACEHOLDER):
            issues.append(_issue("error", where, "还留着「待填写」占位"))

        for lineno, line in enumerate(body.splitlines(), 1):
            if BULLET.match(line):
                issues.append(_issue("error", where, f"第 {lineno} 行用了项目符号/编号，书写要求禁止（层级改用「1.1」「（1）」）"))

        try:
            blocks = mdlite.parse_blocks(body)
            # @table from= / @figure from= 的题注和内容在 calc/*.json 里，先展开再检查
            blocks = mdlite.expand(blocks, cfg.calc)
        except (ValueError, FileNotFoundError) as exc:
            issues.append(_issue("error", where, str(exc)))
            continue

        text_chars = 0
        for kind, *rest in blocks:
            if kind in ("para", "heading"):
                text_chars += len(str(rest[0]) if rest else "")
            if kind == "table":
                item = rest[0]
                stats["tables"] += 1
                cap = str(item.get("caption") or "")
                if not cap.startswith("表"):
                    issues.append(_issue("error", where, f"表题「{cap}」应以「表n」开头"))
                if "from" not in item:
                    cols = item.get("columns") or []
                    if any(not str(c).strip() for c in cols):
                        issues.append(_issue("warn", where, f"表「{cap}」有空的列标题"))
                    for row in item.get("rows") or []:
                        if len(row) != len(cols):
                            issues.append(_issue("error", where, f"表「{cap}」有一行字段数与表头不符"))
                            break
            elif kind == "figure":
                stats["figures"] += 1
                item = rest[0]
                cap = str(item.get("caption") or "")
                if cap and not cap.startswith("图"):
                    issues.append(_issue("error", where, f"图题「{cap}」应以「图n」开头"))
                if sec.get("id") != "handout" and any(k in cap for k in RAW_PHOTO):
                    issues.append(_issue("error", where, "不能直接贴原始数据照片，必须重新录入成电子表格"))
                p = item.get("path")
                if p and "from" not in item:
                    fp = Path(p)
                    if not fp.is_absolute():
                        fp = cfg.resolve(fp)
                    if not fp.exists():
                        issues.append(_issue("error", where, f"图文件不存在：{p}"))
            elif kind == "formula":
                stats["formulas"] += 1

        if sec.get("maxPages") and text_chars > PRINCIPLE_LIMIT and sec.get("id") == "principle":
            issues.append(_issue("warn", where, f"正文 {text_chars} 字，可能超过 {sec['maxPages']} 页（书写要求：实验原理 ≤2 页，公式与电路图另计）"))

        if sec.get("id") in ("data", "processing"):
            if not re.search(r"相对误差|相对不确定度|不确定度", body):
                issues.append(_issue("warn", where, "数据处理里没找到不确定度/相对误差的表述"))
        if sec.get("id") == "handout" and stats["figures"] == 0 and not re.search(r"!\[", body):
            issues.append(_issue("warn", where, "预习报告一节建议放一张教师签字的扫描件"))

    if stats["tables"] == 0:
        issues.append(_issue("warn", "全篇", "没有任何表格：原始数据必须录成电子表"))
    if stats["formulas"] == 0:
        issues.append(_issue("warn", "全篇", "没有公式：实验原理应有公式（用 @formula 写 LaTeX）"))
    if report.get("legacyProcessing"):
        issues.append(
            _issue(
                "warn",
                "report.yaml",
                f"旧版十节结构已合并，但 {report['legacyProcessing']} 尚未并入合并后的正文——"
                "把它的内容按分实验拆开、接在对应表格之后，再删掉该文件",
            )
        )

    out_dir = cfg.out
    target = None
    if req.get("file"):
        target = Path(str(req["file"]))
        if not target.is_absolute():
            target = cfg.resolve(target)
        if not target.exists():
            issues.append(_issue("error", "产物", f"要检查的文件不存在：{target}"))
    if target is None and out_dir.exists():
        target = max(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, default=None)
    if target is not None and target.exists() and target.suffix.lower() == ".pdf":
        if not re.match(r"^\d{2}-\d{2}-\d+-.+-\d{8}-.+\.pdf$", target.name):
            issues.append(
                _issue(
                    "info",
                    "文件名",
                    f"提交文件名建议「班级号-分组号-座位号-姓名-实验名称缩写-日期-指导教师.pdf」，当前是 {target.name}",
                )
            )

    strict = bool(req.get("strict"))
    errors = [i for i in issues if i["level"] == "error"]
    warns = [i for i in issues if i["level"] == "warn"]
    ready = not errors and not (strict and warns)
    head = "可以交" if ready else (f"还有 {len(errors)} 处必须修" if errors else f"严格模式下 {len(warns)} 条警告也算失败")
    summary = (
        f"{head}：{stats['sections']} 节 / {stats['chars']} 字 / 表 {stats['tables']} / 图 {stats['figures']} / 公式 {stats['formulas']}；"
        f"警告 {len(warns)} 条"
    )
    shown = [i for i in issues if i["level"] in ("error", "warn")][:8]
    if shown:
        summary += "\n" + "\n".join(f"- [{i['level']}] {i['where']}：{i['message']}" for i in shown)
    return {
        "ok": True,
        "ready": ready,
        "issues": issues,
        "stats": stats,
        "summary": summary,
        "data": {"ready": ready, "issues": issues, "stats": stats},
        "nextStep": "都通过就 build 出 docx/pdf 交作业" if ready else "按 issues 逐条修正 sections/*.md 后重跑 check",
    }
