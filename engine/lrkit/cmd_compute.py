"""compute：跑 calc/compute.py（每个实验自己写的一小段算数脚本）。

约定（写进文档，也是 check 的依据）：
  calc/compute.py 用标准库 + numpy/scipy 算数，必须写 calc/result.json；
  需要进报告的表格写到 calc/tables/<id>.json：{caption, columns, rows, note?}
  需要进报告的图写到 calc/figures/<id>.json：{path, caption, width?}
  path 相对工程目录（一般是 ../figures/<name>.png，即工程 figures/ 下）。
"""
from __future__ import annotations

import math
import os
import subprocess
import sys
from pathlib import Path

from .config import Config
from .util import jsonable, read_json, write_json

TEMPLATE = '''"""compute.py —— 本实验的算数脚本（由 Agent 按实验内容填写）。

只做"数据 → 数字"这一段：读数整理、逐差/最小二乘、A/B 类不确定度、相对误差。
图表别在这里画，交给 figures 目录里的作图脚本；但表和图的数据可以在这里生成。
结果写 calc/result.json，供 build 和 check 使用。
"""

import json
from pathlib import Path

CALC = Path(__file__).resolve().parent
PROJ = CALC.parent
DATA = PROJ / "data"


def main() -> dict:
    result = {"summary": "（待填写）", "values": {}, "notes": []}

    # 例子：
    # rows = parse_csv(DATA / "raw.csv")
    # mean = sum(rows) / len(rows); result["values"]["mean"] = mean

    (CALC / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    print(json.dumps(main(), ensure_ascii=False))
'''


def run(cfg: Config, req: dict) -> dict:
    script = Path(req.get("script") or (cfg.calc / "compute.py"))
    if not script.is_absolute():
        script = cfg.resolve(script)
    if not script.exists():
        if req.get("scaffold"):
            script.parent.mkdir(parents=True, exist_ok=True)
            script.write_text(TEMPLATE, encoding="utf-8")
            return {
                "created": str(script),
                "summary": f"已生成 {script.name} 模板，按本实验数据补齐算数后重新 compute",
                "nextStep": "填写 calc/compute.py（数据在 data/，参考 book/guide.md 的公式）",
            }
        raise FileNotFoundError(
            f"没有 {script}。labreport_compute 需要一段本实验的算数脚本：\n"
            f"  - 传 scaffold=true 可先生成模板\n"
            f"  - 或自己写 calc/compute.py（结果写 calc/result.json）"
        )

    python = str(cfg.venv_python) if cfg.venv_python.exists() else sys.executable
    env = dict(os.environ)
    env.update(
        {
            "PYTHONUTF8": "1",
            "PYTHONIOENCODING": "utf-8",
            "PYTHONPATH": os.pathsep.join([str(cfg.engine_dir), env.get("PYTHONPATH", "")]).strip(os.pathsep),
            "LR_PROJECT_DIR": str(cfg.project_dir),
            "LR_DATA_DIR": str(cfg.data_dir),
            "LR_CALC_DIR": str(cfg.calc),
            "LR_FIGURES_DIR": str(cfg.figures),
        }
    )
    r = subprocess.run(
        [python, str(script)],
        cwd=str(cfg.project_dir),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        timeout=int(req.get("timeoutSec") or 300),
    )
    log = (r.stdout or "") + (("\n[stderr]\n" + r.stderr) if r.stderr else "")
    (cfg.logs / "compute.log").parent.mkdir(parents=True, exist_ok=True)
    (cfg.logs / "compute.log").write_text(log, encoding="utf-8")

    if r.returncode != 0:
        tail = log.strip()[-2000:]
        raise RuntimeError(f"{script.name} 运行失败（exit {r.returncode}）：\n{tail}")

    result_path = cfg.calc / "result.json"
    if not result_path.exists():
        raise RuntimeError(
            f"{script.name} 跑完了但没写 calc/result.json。\n"
            "算数脚本必须把结果写成 JSON：{summary, values, notes, ...}"
        )
    result = read_json(result_path)
    if not isinstance(result, dict):
        raise RuntimeError("calc/result.json 必须是对象")
    bad = [k for k, v in result.items() if isinstance(v, float) and not math.isfinite(v)]
    result = jsonable(result)

    tables, figures, missing = [], [], []
    if bad:
        missing.append("结果里有算不出来的数（NaN/inf，已置为 null）：" + "、".join(bad))
        write_json(result_path, result)  # 顺手把非法 NaN 清掉，别让后面的命令再读一次坏 JSON
    for sub, bucket in (("tables", tables), ("figures", figures)):
        folder = cfg.calc / sub
        for path in sorted(folder.glob("*.json")) if folder.exists() else []:
            try:
                item = read_json(path)
            except Exception as exc:
                missing.append(f"{sub}/{path.name} 读不了：{exc}")
                continue
            item.setdefault("id", path.stem)
            bucket.append({"id": item["id"], "caption": item.get("caption", "")})
            if sub == "figures":
                fig_path = Path(item.get("path") or "")
                if not fig_path.is_absolute():
                    fig_path = cfg.resolve(fig_path)
                if not fig_path.exists():
                    missing.append(f"图 {item['id']} 指向的文件不存在：{fig_path}")
    write_json(cfg.logs / "compute-summary.json", {"result": result, "tables": tables, "figures": figures})

    head = str(result.get("summary") or "算数完成")[:400]
    bits = [head]
    counts = []
    if tables:
        counts.append(f"表 {len(tables)}")
    if figures:
        counts.append(f"图 {len(figures)}")
    if counts:
        bits.append(f"（{' / '.join(counts)} 已就绪）")
    keys = [k for k in result.keys() if k not in ("summary", "notes")][:8]
    if keys:
        bits.append("；结果字段：" + "、".join(keys))
    if missing:
        bits.append("；需注意：" + "；".join(missing[:4]))

    return {
        "ok": True,
        "summary": "".join(bits),
        "data": {
            "result": result,
            "tables": tables,
            "figures": figures,
            "issues": missing,
            "log": log.strip().splitlines()[-8:],
        },
        "nextStep": "把 result 里的数写进 sections/06_processing.md（表格用 @table，公式用 @formula），确认无误后 build",
    }
