"""pdfinfo：插件侧导出 PDF 之后，回来确认文件并数页数。"""
from __future__ import annotations

from pathlib import Path

from . import pdfout
from .config import Config


def run(cfg: Config, req: dict) -> dict:
    raw = req.get("pdf")
    if not raw:
        return {"ok": False, "summary": "没有给 pdf 路径", "data": {"exists": False, "pages": 0}}
    path = Path(str(raw))
    if not path.is_absolute():
        path = cfg.resolve(path)
    if not path.exists():
        return {
            "ok": False,
            "summary": f"PDF 还没生成：{path}",
            "data": {"exists": False, "pages": 0, "pdf": str(path)},
        }
    pages = pdfout.page_count(path)
    size = path.stat().st_size
    return {
        "ok": True,
        "summary": f"PDF 已生成：{path.name}（{pages} 页 / {size // 1024} KB）",
        "data": {"exists": True, "pages": pages, "pdf": str(path), "bytes": size},
    }
