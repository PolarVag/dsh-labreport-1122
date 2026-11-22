"""通用工具：编码探测、JSON、时间戳、路径。"""
from __future__ import annotations

import json
import math
import re
from datetime import datetime
from pathlib import Path
from typing import Any

TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "utf-16")


def jsonable(obj: Any) -> Any:
    """把任意对象变成严格 JSON 可表示的形式。

    算数脚本很容易把 NaN/inf 写进结果（np.polyfit 拟合失败、除零都很常见），而
    `json.dumps` 默认会输出裸 `NaN`——那不是合法 JSON，Node 侧 JSON.parse 直接炸。
    这里统一把非有限浮点变成 None，让"数值算不出来"显式地表现为 null。
    """
    if isinstance(obj, float):
        return obj if math.isfinite(obj) else None
    if isinstance(obj, dict):
        return {str(k): jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [jsonable(v) for v in obj]
    if isinstance(obj, Path):
        return str(obj)
    return obj


def read_text_auto(path: Path) -> tuple[str, str]:
    """读取文本并返回 (文本, 编码)。

    本机实测 学号.txt / 注意事项.txt 是 GBK，不能一律按 UTF-8 读。
    顺序：BOM → UTF-8 → GB18030 → UTF-16。UTF-8 解码成功即认为可信。
    """
    raw = Path(path).read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "utf-16"
    for enc in ("utf-8", "gb18030"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8-replace"


def write_text(path: Path, text: str) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8", newline="\n")
    return p


def read_json(path: Path, default: Any = None) -> Any:
    p = Path(path)
    if not p.exists():
        return default
    return json.loads(p.read_text(encoding="utf-8"))


def write_json(path: Path, obj: Any) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(jsonable(obj), ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
    return p


def stamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def today_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def safe_name(name: str, fallback: str = "report") -> str:
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", (name or "").strip())
    s = re.sub(r"\s+", "", s).strip("._")
    return s or fallback


def truncate(text: str, limit: int) -> str:
    if text is None:
        return ""
    return text if len(text) <= limit else text[: limit - 1] + "…"


def rel(path: Path, base: Path) -> str:
    try:
        return str(Path(path).resolve().relative_to(Path(base).resolve())).replace("\\", "/")
    except Exception:
        return str(path)


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024.0
    return f"{n:.1f}GB"


def char_count(text: str) -> int:
    """按中文习惯计字数：非空白字符数。"""
    return len(re.sub(r"\s", "", text or ""))
