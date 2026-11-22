"""labreport 引擎入口。

约定：每个子命令只在 stdout 输出**一行** JSON：
    {"ok": bool, "summary": str, "data": {...}}         成功
    {"ok": false, "fatal": true, "message": str}         失败（Node 侧会转成工具错误）
其余诊断信息一律走 stderr。

Node 侧始终用"基础 Python"启动本文件；本文件负责把 venv 准备好并 re-exec 到 venv 解释器。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

ENGINE_DIR = Path(__file__).resolve().parent
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))


def emit(obj: dict) -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    from lrkit.util import jsonable

    sys.stdout.write(json.dumps(jsonable(obj), ensure_ascii=False, allow_nan=False) + "\n")
    sys.stdout.flush()


def fail(message: str, **extra) -> None:
    emit({"ok": False, "fatal": True, "message": message, "summary": message, **extra})
    sys.exit(1)


def parse_args(argv: list[str]) -> tuple[str, Path]:
    if len(argv) < 2:
        fail("用法: cli.py <doctor|init|status|compute|build|check> --request <json>")
    kind = argv[1]
    req_path = None
    if "--request" in argv:
        req_path = argv[argv.index("--request") + 1]
    if not req_path or not Path(req_path).exists():
        fail("缺少 --request <json 文件>")
    return kind, Path(req_path)


def main() -> None:
    kind, req_path = parse_args(sys.argv)
    request = json.loads(req_path.read_text(encoding="utf-8-sig")).get("request", {})
    request.setdefault("requestPath", str(req_path))

    from lrkit import bootstrap
    from lrkit.config import Config

    cfg = Config.from_env(request)

    # doctor 只做诊断，不装依赖、不 re-exec
    if kind == "doctor":
        from lrkit import cmd_doctor

        emit(cmd_doctor.run(cfg, request))
        return

    want_bootstrap = request.get("bootstrap", True) and cfg.auto_bootstrap
    ok, note = bootstrap.ensure(cfg, allow=want_bootstrap)
    if not ok and not cfg.soft_bootstrap:
        fail(f"Python 依赖未就绪：{note}\n修复后重试。")

    in_venv = os.environ.get("LR_IN_VENV") == "1"
    if not in_venv and ok:
        vp = str(bootstrap.venv_python(cfg))
        if Path(vp).resolve() != Path(sys.executable).resolve():
            child = subprocess.run(
                [vp, str(Path(__file__).resolve()), kind, "--request", str(req_path)],
                env={**os.environ, "LR_IN_VENV": "1", "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
            )
            sys.exit(child.returncode)

    from lrkit import commands

    try:
        emit(commands.dispatch(kind, cfg, request))
    except Exception as exc:  # noqa: BLE001
        import traceback

        tb = traceback.format_exc()
        sys.stderr.write(tb)
        fail(f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
