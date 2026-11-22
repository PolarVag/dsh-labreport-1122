"""data：把（手抄/整理好的）原始数据落成 xlsx，放进 output/。

用途：实验只给了原始记录表（照片、扫描件、手写表格）时，先把读数抄成结构化数据，
生成一份 xlsx 交给用户确认；确认之后再写 compute.py 直接读这份 xlsx。

请求格式：
    {"kind": "data", "request": {
        "sheets": [
            {"name": "热电偶", "columns": ["t/℃", "E/mV"], "rows": [[32, 0.77], ...],
             "note": "冷端 0 ℃"}
        ],
        "filename": "可选，默认 <学生>-<实验简称>-实验数据.xlsx"
    }}

返回：{ok, summary, data:{xlsx, sheets:[{name, rows, columns}], totalRows}}
"""
from __future__ import annotations

from pathlib import Path

from . import project
from .config import Config
from .util import safe_name

_BAD_SHEET_CHARS = set('[]:*?/\\')


def _sheet_name(raw: str, used: set[str], index: int) -> str:
    name = "".join("_" if c in _BAD_SHEET_CHARS else c for c in (raw or "").strip())[:31]
    name = name or f"表{index + 1}"
    base, n = name, 2
    while name in used:
        suffix = f"-{n}"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


def _cell_value(value):
    """尽量把读数还原成数字，方便后续 openpyxl / pandas 直接算。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    text = str(value if value is not None else "").strip()
    if not text:
        return None
    cleaned = text.replace("−", "-").replace("＋", "+").replace(",", "")
    try:
        return int(cleaned)
    except ValueError:
        pass
    try:
        return float(cleaned)
    except ValueError:
        return text


def run(cfg: Config, req: dict) -> dict:
    project.prepare(cfg)
    sheets = req.get("sheets") or []
    if not sheets:
        raise RuntimeError("没有可写的数据表：请把每张表的 name / columns / rows 传进来")

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    out_sheets = []
    total_rows = 0
    for index, sheet in enumerate(sheets):
        columns = [str(c) for c in (sheet.get("columns") or [])]
        rows = sheet.get("rows") or []
        if not columns:
            raise RuntimeError(f"第 {index + 1} 张表缺少 columns")
        ws = wb.create_sheet(_sheet_name(str(sheet.get("name") or ""), used, index))
        ws.append(columns)
        for row in rows:
            ws.append([_cell_value(v) for v in row])
            total_rows += 1
        head_font = Font(bold=True)
        thin = Side(style="thin", color="999999")
        for cell in ws[1]:
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(top=thin, bottom=thin)
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 4 for c in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(26, max(10, width))
        ws.freeze_panes = "A2"
        if sheet.get("note"):
            note_row = ws.max_row + 2
            ws.cell(row=note_row, column=1, value=str(sheet["note"])).font = Font(italic=True, color="666666")
        out_sheets.append({"name": ws.title, "columns": columns, "rows": len(rows)})

    meta = dict((project.load_report(cfg) or {}).get("meta") or {})
    default_name = "-".join(x for x in (meta.get("student"), meta.get("short") or meta.get("experiment"), "实验数据") if x)
    filename = str(req.get("filename") or f"{default_name}.xlsx")
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    target = Path(filename)
    if not target.is_absolute():
        target = cfg.out / safe_name(target.name)
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(target))

    summary = f"实验数据已写成 xlsx：{target.name}（{len(out_sheets)} 张表 / {total_rows} 行）"
    return {
        "ok": True,
        "summary": summary,
        "data": {"xlsx": str(target), "sheets": out_sheets, "totalRows": total_rows},
    }
