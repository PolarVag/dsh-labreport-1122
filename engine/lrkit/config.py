"""引擎公共配置：全部来自环境变量（由 Node 侧注入）与 request。"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _bool(name: str, default: bool) -> bool:
    v = os.environ.get(name)
    if v is None:
        return default
    return v not in ("0", "", "false", "False")


@dataclass
class Config:
    plugin_root: Path
    engine_dir: Path
    venv_dir: Path
    base_python: str
    word: str
    mml2omml: str
    fig_width_cm: float
    dpi: int
    auto_bootstrap: bool
    soft_bootstrap: bool
    keep_logs: int
    dir: Path
    project_dir: Path
    request: dict = field(default_factory=dict)

    @classmethod
    def from_env(cls, request: dict) -> "Config":
        plugin_root = Path(os.environ.get("LR_PLUGIN_ROOT") or Path(__file__).resolve().parents[2])
        engine_dir = plugin_root / "engine"
        venv_dir = Path(os.environ.get("LR_VENV_DIR") or (plugin_root.parent / "_labreport" / ".venv"))
        base_python = os.environ.get("LR_PYTHON") or "python"
        dir_raw = request.get("dir") or "."
        project_raw = request.get("projectDir") or str(Path(dir_raw) / "_lr")
        word = os.environ.get("LR_WORD", "") or _auto_word()
        mml2omml = os.environ.get("LR_MML2OMML", "") or _auto_xsl(word)
        return cls(
            plugin_root=plugin_root,
            engine_dir=engine_dir,
            venv_dir=venv_dir,
            base_python=base_python,
            word=word,
            mml2omml=mml2omml,
            fig_width_cm=float(os.environ.get("LR_FIG_WIDTH_CM") or 12),
            dpi=int(os.environ.get("LR_DPI") or 300),
            auto_bootstrap=_bool("LR_AUTO_BOOTSTRAP", True),
            soft_bootstrap=_bool("LR_SOFT_BOOTSTRAP", True),
            keep_logs=int(os.environ.get("LR_KEEP_LOGS") or 30),
            dir=Path(dir_raw).resolve(),
            project_dir=Path(project_raw).resolve(),
            request=request,
        )

    # 工程内常用路径
    # 工程分两个文件夹：
    #   report-making/  一切中间产物（书、素材、正文、计算、图、日志）
    #   output/         最终交付物（docx、pdf、实验数据 xlsx）
    @property
    def making(self) -> Path:
        return self.project_dir / "report-making"

    @property
    def book(self) -> Path:
        return self.making / "book"

    @property
    def media(self) -> Path:
        return self.making / "media"

    @property
    def data_dir(self) -> Path:
        return self.making / "data"

    @property
    def sections(self) -> Path:
        return self.making / "sections"

    @property
    def calc(self) -> Path:
        return self.making / "calc"

    @property
    def figures(self) -> Path:
        return self.making / "figures"

    @property
    def tmp(self) -> Path:
        return self.making / ".tmp"

    @property
    def out(self) -> Path:
        return self.project_dir / "output"

    @property
    def logs(self) -> Path:
        return self.making / "logs"

    # 这些子目录一律住在 report-making/ 下（旧工程里它们直接在 _lr/ 下）
    WORK_DIRS = ("book", "media", "data", "sections", "calc", "figures", "logs", ".tmp")

    def resolve(self, rel: str | Path) -> Path:
        """把 md/report.yaml 里写的相对路径解析成真实路径。

        规则：`figures/x.png` 这类工作目录相对路径落到 report-making/ 下；
        `../media/x.jpg`、`report-making/...`、绝对路径原样解析。
        """
        p = Path(rel)
        if p.is_absolute():
            return p
        parts = p.parts
        if parts and parts[0] in self.WORK_DIRS:
            return self.making / p
        cand = self.project_dir / p
        if cand.exists():
            return cand
        alt = self.making / p
        return alt if alt.exists() else cand

    @property
    def report_yaml(self) -> Path:
        return self.project_dir / "report.yaml"

    @property
    def venv_python(self) -> Path:
        if os.name == "nt":
            return self.venv_dir / "Scripts" / "python.exe"
        return self.venv_dir / "bin" / "python"


_WORD_GUESSES = [
    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
]


def _auto_word() -> str:
    for cand in _WORD_GUESSES:
        if Path(cand).exists():
            return cand
    return ""


def _auto_xsl(word: str) -> str:
    cands = []
    if word:
        cands.append(Path(word).with_name("MML2OMML.XSL"))
    cands += [Path(p).with_name("MML2OMML.XSL") for p in _WORD_GUESSES]
    for cand in cands:
        if cand.exists():
            return str(cand)
    return ""