r"""md-lite：报告正文的唯一书写格式。

刻意只支持写实验报告需要的几样东西，避免 Agent 写出 Word 里不存在的花活。
块类型：
  ('heading', level, text)          # / ## / ###
  ('para', text)                    普通段落（一行一段，保留 **粗体** 与 $行内公式$）
  ('table', dict)                   caption/columns/rows/note/widths
  ('figure', dict)                  path/caption/width
  ('formula', dict)                 tex/number
  ('pagebreak',)
子块指令（整行）：
  @table 表1 不同温度下的电阻值        后面紧跟 | 开头的行
  @table from=tab1                    从 calc/tables/tab1.json 取（compute 产出）
  ![图1 曲线](figures/fig1.png){width=12}
  @figure from=fig1
  @formula \bar{x}=\frac{1}{n}\sum x_i
  $$ ... $$
  @pagebreak
"""
from __future__ import annotations

import json
import re
from pathlib import Path

_HEADING = re.compile(r"^(#{1,3})\s+(.*\S)\s*$")
_IMAGE = re.compile(r"^!\[(?P<cap>[^\]]*)\]\((?P<path>[^)]+)\)\s*(?:\{(?P<opts>[^}]*)\})?\s*$")
_TABLE_FROM = re.compile(r"^@table\s+from=(?P<id>[\w.\-]+)\s*$")
_FIGURE_FROM = re.compile(r"^@figure\s+from=(?P<id>[\w.\-]+)\s*$")
_SEP_CELL = re.compile(r"^:?-{3,}:?$")


def _opts(text: str) -> dict:
    out = {}
    for part in (text or "").split(","):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def _split_row(line: str) -> list[str]:
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|"):
        body = body[:-1]
    return [c.strip() for c in body.split("|")]


def parse_blocks(text: str) -> list[tuple]:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks: list[tuple] = []
    i = 0
    n = len(lines)
    while i < n:
        raw = lines[i]
        line = raw.strip()
        if not line:
            i += 1
            continue

        m = _HEADING.match(line)
        if m:
            blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        if line == "@pagebreak":
            blocks.append(("pagebreak",))
            i += 1
            continue

        m = _TABLE_FROM.match(line)
        if m:
            blocks.append(("table", {"from": m.group("id")}))
            i += 1
            continue

        if line.startswith("@table"):
            caption = line[len("@table"):].strip()
            rows: list[list[str]] = []
            i += 1
            while i < n and lines[i].strip().startswith("|"):
                cells = _split_row(lines[i])
                if not all(_SEP_CELL.match(c) for c in cells if c) or not any(cells):
                    rows.append(cells)
                i += 1
            if not rows:
                raise ValueError(f"@table {caption!r} 后面没有以 | 开头的数据行")
            header = rows[0]
            body = [r for r in rows[1:]]
            width = len(header)
            body = [(r + [""] * width)[:width] for r in body]
            blocks.append(("table", {"caption": caption, "columns": header, "rows": body}))
            continue

        m = _FIGURE_FROM.match(line)
        if m:
            blocks.append(("figure", {"from": m.group("id")}))
            i += 1
            continue

        m = _IMAGE.match(line)
        if m:
            opts = _opts(m.group("opts"))
            item = {"path": m.group("path").strip(), "caption": m.group("cap").strip()}
            if "width" in opts:
                item["width"] = float(opts["width"])
            blocks.append(("figure", item))
            i += 1
            continue

        if line.startswith("@formula"):
            tex = line[len("@formula"):].strip()
            number = None
            tail = re.search(r"#\((\w+)\)\s*$", tex)
            if tail:
                number = tail.group(1)
                tex = tex[: tail.start()].strip()
            blocks.append(("formula", {"tex": tex, "number": number}))
            i += 1
            continue

        if line.startswith("$$") and line.endswith("$$") and len(line) > 4:
            blocks.append(("formula", {"tex": line[2:-2].strip(), "number": None}))
            i += 1
            continue

        blocks.append(("para", line))
        i += 1
    return blocks


def expand(blocks: list[tuple], calc_dir: Path) -> list[tuple]:
    """把 from=<id> 解析成 compute 写出的 JSON。"""
    out: list[tuple] = []
    for kind, *rest in blocks:
        if kind in ("table", "figure") and rest and "from" in rest[0]:
            ident = rest[0]["from"]
            sub = "tables" if kind == "table" else "figures"
            path = calc_dir / sub / f"{ident}.json"
            if not path.exists():
                raise FileNotFoundError(
                    f"缺少 {path}：计算脚本要先写 calc/{sub}/{ident}.json，或用 @table/@figure 直接写内容"
                )
            data = json.loads(path.read_text(encoding="utf-8-sig"))
            data.setdefault("id", ident)
            out.append((kind, data))
        else:
            out.append((kind, *rest))
    return out
