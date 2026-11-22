"""venv 引导：建 venv（复用系统科学栈）并安装文档处理依赖。

只用标准库，所以基础 Python 也能跑这一段。
本机 pip 不可用（沙箱禁止 tempfile.mkdtemp 目录），因此优先用
lrkit.pyfetch 自己下载 wheel 解包，pip 仅作兜底。
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from .config import Config

# (pip 包名, 需要能 import 的模块名)
GROUPS: list[tuple[str, list[str]]] = [
    ("lxml", ["lxml.etree"]),
    ("typing_extensions", ["typing_extensions"]),
    ("python-docx", ["docx"]),
    ("PyYAML", ["yaml"]),
    ("pymupdf", ["fitz"]),
]
DEPS = [(mods[0], pkg) for pkg, mods in GROUPS]
PIP_PACKAGES = [pkg for pkg, _ in GROUPS]
PROBE = (
    "import docx, lxml.etree, yaml\n"
    "try:\n"
    "    import fitz\n"
    "except ImportError:\n"
    "    import pymupdf as fitz\n"
)


def venv_python(cfg: Config) -> Path:
    if os.name == "nt":
        return cfg.venv_dir / "Scripts" / "python.exe"
    return cfg.venv_dir / "bin" / "python"


def _run(argv: list[str], timeout: int = 900) -> subprocess.CompletedProcess:
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    env.pop("PYTHONPATH", None)
    return subprocess.run(argv, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env)


def missing_deps(cfg: Config) -> list[str]:
    vp = venv_python(cfg)
    if not vp.exists():
        return [pkg for pkg, _ in GROUPS]
    r = _run([str(vp), "-c", PROBE], timeout=180)
    return [] if r.returncode == 0 else [pkg for pkg, _ in GROUPS]


def create_venv(cfg: Config) -> str:
    cfg.venv_dir.parent.mkdir(parents=True, exist_ok=True)
    # --without-pip：本机 ensurepip 会因临时目录被沙箱拒绝而失败，pip 由 pyfetch 取代
    r = _run([cfg.base_python, "-m", "venv", "--system-site-packages", "--without-pip", str(cfg.venv_dir)], timeout=600)
    if r.returncode != 0 or not venv_python(cfg).exists():
        raise RuntimeError(f"创建 venv 失败：\n{(r.stderr or r.stdout)[-1500:]}")
    return "created"


def _pip_install(cfg: Config) -> None:
    vp = venv_python(cfg)
    r = _run([str(vp), "-m", "pip", "install", "--disable-pip-version-check", "--no-input", *PIP_PACKAGES], timeout=900)
    if r.returncode != 0:
        raise RuntimeError(f"{(r.stderr or r.stdout)[-2000:]}")


def install_deps(cfg: Config) -> str:
    from . import pyfetch

    errors: list[str] = []
    try:
        pyfetch.ensure_wheels(cfg, GROUPS)
    except Exception as exc:  # noqa: BLE001
        errors.append(f"wheel 直装失败：{exc}")
    if missing_deps(cfg):
        try:
            _pip_install(cfg)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"pip 兜底失败：{exc}")
    if missing_deps(cfg):
        vp = venv_python(cfg)
        raise RuntimeError(
            "安装文档依赖失败：\n" + "\n".join(errors) + "\n"
            "可手动执行（本机 pip 可能同样失败）：\n"
            f'  "{vp}" -m pip install {" ".join(PIP_PACKAGES)}\n'
        )
    return "installed"


def ensure(cfg: Config, allow: bool = True) -> tuple[bool, str]:
    """确保 venv 与依赖可用。返回 (ok, note)。"""
    vp = venv_python(cfg)
    if not vp.exists():
        if not allow:
            return False, f"venv 不存在：{cfg.venv_dir}（未允许自动安装）"
        try:
            create_venv(cfg)
        except Exception as exc:  # noqa: BLE001
            return False, str(exc)
    if missing_deps(cfg):
        if not allow:
            return False, "Python 依赖缺失（未允许自动安装）：" + "、".join(PIP_PACKAGES)
        try:
            install_deps(cfg)
        except Exception as exc:  # noqa: BLE001
            return False, str(exc)
    return True, f"venv 就绪：{vp}"


def describe(cfg: Config) -> dict:
    vp = venv_python(cfg)
    info = {"venv": str(cfg.venv_dir), "python": str(vp), "exists": vp.exists()}
    if vp.exists():
        try:
            info["missing"] = missing_deps(cfg)
        except Exception as exc:  # noqa: BLE001
            info["missing"] = [str(exc)]
    else:
        info["missing"] = [pkg for pkg, _ in GROUPS]
    info["base_python"] = cfg.base_python
    info["python_version"] = sys.version.split()[0]
    return info
