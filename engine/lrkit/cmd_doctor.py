"""doctor：环境自检（纯标准库，不装依赖、不 re-exec）。"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

from . import bootstrap
from .config import Config


def _check(name: str, ok: bool, detail: str = "", fix: str = "") -> dict:
    return {"name": name, "ok": bool(ok), "detail": detail, "fix": fix}


def _version(python: str) -> str:
    try:
        r = subprocess.run([python, "-c", "import sys;print(sys.version.split()[0])"], capture_output=True, text=True, timeout=60)
        return (r.stdout or r.stderr).strip().splitlines()[0] if r.returncode == 0 else ""
    except Exception:
        return ""


def run(cfg: Config, request: dict) -> dict:
    checks: list[dict] = []

    checks.append(_check("基础 Python", Path(cfg.base_python).exists() or shutil.which(cfg.base_python) is not None,
                         f"{cfg.base_python}  {_version(cfg.base_python)}",
                         "在插件 config.pythonPath 指向一个 Python 3.10+"))

    vin = bootstrap.describe(cfg)
    checks.append(_check("工程 venv", vin.get("exists", False), str(cfg.venv_dir),
                         "重跑 labreport_init 会自动创建"))
    missing = vin.get("missing") or []
    checks.append(_check("文档依赖", not missing, "缺：" + "、".join(missing) if missing else "python-docx / lxml / pymupdf / PyYAML 就绪",
                         f'"{bootstrap.venv_python(cfg)}" -m pip install {" ".join(bootstrap.PIP_PACKAGES)}'))

    node = shutil.which("node") or "node"
    node_ok = bool(shutil.which("node"))
    checks.append(_check("Node.js", node_ok, node if node_ok else "PATH 里找不到 node", "安装 Node.js 20+"))

    temml = cfg.plugin_root / "vendor" / "temml.mjs"
    checks.append(_check("TeX→MathML（temml）", temml.exists(), str(temml), "从 温度测量\\_nodevendor\\node_modules\\temml\\dist 复制 temml.mjs 到 vendor/"))

    word_ok = bool(cfg.word) and Path(cfg.word).exists()
    checks.append(_check("Microsoft Word", word_ok, cfg.word or "未配置", "在插件 config.wordPath 指定 WINWORD.EXE（PDF 导出与 .doc 转换需要）"))

    xsl_ok = bool(cfg.mml2omml) and Path(cfg.mml2omml).exists()
    checks.append(_check("MML2OMML.XSL", xsl_ok, cfg.mml2omml or "未配置", "通常随 Office 安装在 root\\Office16\\ 下"))

    font = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "simsun.ttc"
    checks.append(_check("中文字体", font.exists(), str(font), "缺少宋体，图表中文会变方框"))

    writable = True
    try:
        cfg.project_dir.mkdir(parents=True, exist_ok=True)
        probe = cfg.project_dir / ".write-probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except Exception as exc:  # noqa: BLE001
        writable = False
        checks.append(_check("工程目录可写", False, f"{cfg.project_dir}: {exc}", "检查目录权限"))
    else:
        checks.append(_check("工程目录可写", writable, str(cfg.project_dir)))

    bad = [c for c in checks if not c["ok"]]
    lines = ["环境自检：" + ("全部通过 ✓" if not bad else f"{len(bad)} 项需要处理")]
    for c in checks:
        lines.append(f"  {'✓' if c['ok'] else '✗'} {c['name']}：{c['detail']}")
        if not c["ok"] and c["fix"]:
            lines.append(f"     修复：{c['fix']}")
    return {
        "ok": not bad,
        "summary": "\n".join(lines),
        "data": {"checks": checks, "venv": vin, "project": str(cfg.project_dir), "python": sys.version.split()[0]},
    }
