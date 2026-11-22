"""docx → pdf：走本机 Word 的 ExportAsFixedFormat。

实测：本机没有 LibreOffice / pandoc，Word COM 是唯一保真路径；
但 Word 自动化在 DSH 沙箱里会挂在 Documents.Open（窗口/UI 被限制），
在放宽权限的普通 PowerShell 里可以正常导出。
所以这里做成"尽力而为"：先试，超时就杀掉 Word 并留一个可直接运行的 .ps1 给用户。
"""
from __future__ import annotations

import subprocess
from pathlib import Path

from .config import Config

_PS = r"""$ErrorActionPreference = 'Stop'
$src = '{src}'
$dst = '{dst}'
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open($src, $false, $true)
    $doc.ExportAsFixedFormat($dst, 17)
}} finally {{
    if ($doc) {{ $doc.Close(0) }}
    if ($word) {{ $word.Quit() }}
}}
if (-not (Test-Path $dst)) {{ throw 'ExportAsFixedFormat 没有产出 PDF' }}
Write-Output ('PDFOK ' + $dst)
"""


def write_script(docx_path: Path, pdf_path: Path) -> Path:
    script = docx_path.with_name(docx_path.stem + "-转PDF.ps1")
    script.write_text(
        _PS.format(src=str(docx_path).replace("'", "''"), dst=str(pdf_path).replace("'", "''")),
        encoding="utf-8-sig",
    )
    return script


def _kill_word() -> None:
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=30)


def prepare_script(cfg: Config, docx_path: Path, pdf_path: Path) -> tuple[Path, Path]:
    """写好转换脚本但不在 Python 里执行 Word。

    实测：由 Python 进程去起 Word 时，Word 会卡死在 Documents.Open（父进程是 Python 时
    COM 打开的文档一直不返回）；同样一段脚本由 Node/pwsh 直接起就 10 秒正常导出。
    所以导出动作交给插件侧的 lib/pdf.js 执行，Python 只负责生成脚本。
    """
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"找不到 docx：{docx_path}")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        pdf_path.unlink()
    return write_script(docx_path, pdf_path), pdf_path


def page_count(pdf_path: Path) -> int:
    try:
        import pymupdf

        with pymupdf.open(str(pdf_path)) as doc:
            return doc.page_count
    except Exception:
        return 0
