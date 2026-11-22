"""纯标准库的 wheel 下载 + 解包安装。

为什么不直接用 pip：
  本机 DSH 沙箱会把 tempfile.mkdtemp() 建出的临时目录判为不可访问
  （连 icacls 都 Access denied），pip 的 unpack/install 阶段必然失败：
    PermissionError: [Errno 13] ... \\pip-unpack-xxxx\\xxx.whl.metadata
  而普通 os.makedirs 建出的目录读写正常。所以这里自己下载 wheel、
  自己解包到 site-packages，全程只用 uri 请求 + zipfile，不建临时目录。
"""
from __future__ import annotations

import json
import os
import sys
import urllib.request
import zipfile
from pathlib import Path

INDEX = "https://pypi.org/pypi/{project}/json"
USER_AGENT = "dsh-labreport/0.1 (+python-stdlib)"


def _py_tag() -> tuple[str, int]:
    return (f"cp{sys.version_info.major}{sys.version_info.minor}", sys.version_info.major * 100 + sys.version_info.minor)


def _fetch_json(project: str, timeout: int = 60) -> dict:
    req = urllib.request.Request(INDEX.format(project=project), headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _vkey(version: str) -> tuple:
    out = []
    for part in version.replace("-", ".").split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        out.append(int(digits) if digits else 0)
    return tuple(out)


def _compatible_wheel(filename: str) -> tuple[bool, int]:
    """返回 (是否可用, 优先级)。优先级越大越优先：本机原生 > abi3 > 纯 python。"""
    if not filename.endswith(".whl"):
        return False, 0
    parts = filename[:-4].split("-")
    if len(parts) < 5:
        return False, 0
    pytags, abitags, plats = parts[-3].split("."), parts[-2].split("."), parts[-1].split(".")
    here = sys.platform == "win32"
    bits = 64 if sys.maxsize > 2**32 else 32
    wanted = {"win_amd64": 64, "win32": 32, "win_arm64": 64} if here else {}
    flat = [p for p in plats if p == "any" or wanted.get(p) == bits]
    if not flat:
        return False, 0
    rank = 0
    if not any(p == "any" for p in flat):
        rank += 5  # 本机原生平台优先于纯 python
    tag, pyver = _py_tag()
    rank = 0
    for abi in abitags:
        if abi == tag:
            if tag in pytags:
                rank = max(rank, 40)
        elif abi == "none":
            for pyt in pytags:
                if pyt == tag:
                    rank = max(rank, 30)
                elif pyt.startswith("py3") or pyt == "py2.py3":
                    rank = max(rank, 10)
        elif abi == "abi3":
            for pyt in pytags:
                if pyt.startswith("cp3"):
                    minor = int("".join(ch for ch in pyt[3:] if ch.isdigit()) or 0)
                    if minor <= pyver - 300:
                        rank = max(rank, 20)
    return (rank > 0), rank


def pick_wheel(project: str, data: dict) -> dict | None:
    version = (data.get("info") or {}).get("version") or ""
    best = None
    for f in data.get("urls") or []:
        ok, rank = _compatible_wheel(f.get("filename") or "")
        if ok and (best is None or rank > best["rank"]):
            best = {"rank": rank, "filename": f["filename"], "url": f["url"], "version": version}
    return best


def download(url: str, dest: Path, timeout: int = 300) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as fh:
        while True:
            chunk = resp.read(1 << 20)
            if not chunk:
                break
            fh.write(chunk)
    return dest


def site_packages(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Lib" / "site-packages"
    return venv_dir / "lib" / f"python{sys.version_info.major}.{sys.version_info.minor}" / "site-packages"


def _data_map(venv_dir: Path, sp: Path) -> dict[str, Path]:
    return {
        "purelib": sp,
        "platlib": sp,
        "scripts": venv_dir / ("Scripts" if os.name == "nt" else "bin"),
        "data": venv_dir,
        "headers": venv_dir / "include",
    }


def extract(wheel: Path, venv_dir: Path) -> list[str]:
    sp = site_packages(venv_dir)
    sp.mkdir(parents=True, exist_ok=True)
    mapping = _data_map(venv_dir, sp)
    written: list[str] = []
    with zipfile.ZipFile(wheel) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            target: Path
            if ".data/" in name:
                prefix, _, rest = name.partition(".data/")
                scheme, _, rel = rest.partition("/")
                base = mapping.get(scheme)
                if base is None or not rel:
                    continue
                target = base / rel
            else:
                target = sp / name
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as dst:
                while True:
                    chunk = src.read(1 << 20)
                    if not chunk:
                        break
                    dst.write(chunk)
            written.append(str(target))
    return written


def install(project: str, venv_dir: Path, wheels_dir: Path, module: str = "") -> dict:
    data = _fetch_json(project)
    chosen = pick_wheel(project, data)
    if not chosen:
        raise RuntimeError(f"{project} 没有适用于 CPython {sys.version_info.major}.{sys.version_info.minor}/本机平台的 wheel")
    wheel = download(chosen["url"], wheels_dir / chosen["filename"])
    files = extract(wheel, venv_dir)
    return {"project": project, "version": chosen["version"], "wheel": str(wheel), "files": len(files)}


def _probe(venv_python: Path, modules: list[str]) -> list[str]:
    import subprocess

    code = "import importlib,sys\nbad=[]\nfor m in sys.argv[1:]:\n  try: importlib.import_module(m)\n  except Exception: bad.append(m)\nprint('|'.join(bad))\n"
    r = subprocess.run([str(venv_python), "-c", code, *modules], capture_output=True, text=True, timeout=180)
    if r.returncode != 0:
        return list(modules)
    text = (r.stdout or "").strip()
    return [m for m in text.split("|") if m]


def ensure_wheels(cfg, groups: list[tuple[str, list[str]]]) -> list[dict]:
    """groups: [(pip 包名, [需要能 import 的模块名...]), ...]"""
    from .bootstrap import venv_python  # 避免循环导入

    vp = venv_python(cfg)
    wheels_dir = Path(cfg.venv_dir).parent / "wheels"
    done = []
    for project, modules in groups:
        if modules and not _probe(vp, modules):
            continue
        done.append(install(project, Path(cfg.venv_dir), wheels_dir))
    return done
