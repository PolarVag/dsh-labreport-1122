"""子命令分发。"""
from __future__ import annotations

from typing import Any

from .config import Config


def dispatch(kind: str, cfg: Config, request: dict[str, Any]) -> dict[str, Any]:
    if kind == "init":
        from . import cmd_init

        return cmd_init.run(cfg, request)
    if kind == "status":
        from . import cmd_status

        return cmd_status.run(cfg, request)
    if kind == "data":
        from . import cmd_data

        return cmd_data.run(cfg, request)
    if kind == "compute":
        from . import cmd_compute

        return cmd_compute.run(cfg, request)
    if kind == "build":
        from . import cmd_build

        return cmd_build.run(cfg, request)
    if kind == "check":
        from . import cmd_check

        return cmd_check.run(cfg, request)
    if kind == "pdfinfo":
        from . import cmd_pdfinfo

        return cmd_pdfinfo.run(cfg, request)
    return {"ok": False, "fatal": True, "message": f"未知子命令：{kind}"}
