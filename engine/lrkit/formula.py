"""公式：TeX → MathML（temml / node）→ OMML（MML2OMML.XSL）。

Word 里只认 OMML，PNG 会被判为"公式截图"不合规，所以整条链必须是矢量的。
"""
from __future__ import annotations

import importlib.util
import shutil
import subprocess
from pathlib import Path

from lxml import etree

from .config import Config
from .util import read_json, write_json

_XSLT_CACHE: dict[str, etree.XSLT] = {}


def _node() -> str:
    found = shutil.which("node")
    if not found:
        raise RuntimeError("找不到 node，无法渲染公式（TeX→MathML）")
    return found


def mathml_batch(cfg: Config, items: list[dict]) -> dict[str, str]:
    """items: [{id, tex, display}] → {id: mathml}；任何一条失败都抛错。"""
    tmp = cfg.tmp
    tmp.mkdir(parents=True, exist_ok=True)
    inp, outp = tmp / "mathml-in.json", tmp / "mathml-out.json"
    write_json(inp, items)
    script = cfg.engine_dir / "tools" / "tex2mathml.mjs"
    if not script.exists():
        raise RuntimeError(f"缺少 {script}")
    r = subprocess.run(
        [_node(), str(script), str(inp), str(outp)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
    )
    if r.returncode != 0 or not outp.exists():
        raise RuntimeError(f"temml 渲染失败：{(r.stderr or r.stdout)[-1500:]}")
    res = read_json(outp)
    bad = {x["id"]: x.get("error") for x in res if not x.get("ok")}
    if bad:
        detail = "\n".join(f"  {k}: {v}" for k, v in bad.items())
        raise RuntimeError("以下公式 TeX 写法有误：\n" + detail)
    return {x["id"]: x["mathml"] for x in res}


def _transformer(cfg: Config) -> etree.XSLT:
    key = str(cfg.mml2omml)
    if key in _XSLT_CACHE:
        return _XSLT_CACHE[key]
    if not cfg.mml2omml or not Path(cfg.mml2omml).exists():
        raise RuntimeError(
            "找不到 MML2OMML.XSL，无法把 MathML 转成 Word 原生公式。\n"
            "它随 Office 安装（通常在 C:\\Program Files\\Microsoft Office\\root\\Office16\\MML2OMML.XSL），"
            "可在插件 config.mml2ommlXsl 指定。"
        )
    trans = etree.XSLT(etree.parse(cfg.mml2omml))
    _XSLT_CACHE[key] = trans
    return trans


def mathml_to_omml(cfg: Config, mathml: str) -> etree._Element:
    root = etree.fromstring(mathml.encode("utf-8"), etree.XMLParser(resolve_entities=False, huge_tree=True))
    out = _transformer(cfg)(root)
    elem = out.getroot()
    if elem is None:
        raise RuntimeError("MathML→OMML 转换结果为空")
    return elem


def omml_string(cfg: Config, mathml: str) -> str:
    return etree.tostring(mathml_to_omml(cfg, mathml), encoding="unicode")


def build_formulas(cfg: Config, formulas: list[dict]) -> list[dict]:
    """formulas: [{id, tex, display}] → 写 calc/formulas/<id>.omml.xml 并回填 omml 字符串。"""
    if not formulas:
        return []
    mathmls = mathml_batch(cfg, [{"id": f["id"], "tex": f.get("tex", ""), "display": f.get("display", True)} for f in formulas])
    out_dir = cfg.calc / "formulas"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = []
    for f in formulas:
        xml = omml_string(cfg, mathmls[f["id"]])
        path = out_dir / f"{f['id']}.omml.xml"
        path.write_text(xml, encoding="utf-8")
        out.append({**f, "omml": xml, "path": str(path)})
    return out


def have_lxml() -> bool:
    return importlib.util.find_spec("lxml") is not None
