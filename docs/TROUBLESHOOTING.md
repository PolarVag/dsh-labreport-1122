# 排障手册

按「症状 → 原因 → 处理」排列，都是实际踩过的坑。

## 装不上 / 装上不生效

**`install_bundle` 报 failed，或工具表里没有 `labreport_*`**

- 看返回里的 `warnings`：常见是 `cannot get property "xxx" without inject`。
  插件里读可选服务必须用 `ctx.get('name')`，不能直接 `ctx.xxx`。
- 改完要重新加载：`set_bundle enabled=false` → `enabled=true`。只重启会话不算重载。
- 装的是**目录**不是包名：`target` 传 clone 下来的绝对路径。

## 路径被判越界

```
dir越界：...\温度测量 允许的根目录：C:\Users\<你>\.dsh\profiles\desktop
```

插件加载时还不知道当前会话目录，所以配置里的 `workspaceRoot` 可能是 profile 目录。
现在会在每次调用时用 `workingDirectory.ensure(agent)` 取真实会话目录重新绑定；
若你手动配了 `workspaceRoot`，它会被当成唯一允许的根——确认那是你想限制的目录。

## Python / 依赖

**`No module named pip`、`PermissionError: pip-unpack-...`**

某些沙箱把 `tempfile.mkdtemp()` 建的目录判为不可访问，pip 必然失败。插件已改为
`pyfetch.py`：直接按平台/ABI 挑 wheel，从 pypi 下载后解压进 venv，全程不用 pip。
若仍报错，说明 wheel 挑不出来（比如没有对应 win_amd64 wheel），把 `labreport_init` 的报错原文贴出来。

**`AttributeError: 'Config' object has no attribute ...`**

引擎版本与插件版本不匹配：仓库更新后请整体更新（`git pull`），不要只替换单个文件。

**compute 报 `ModuleNotFoundError: numpy`**

基础解释器没有科学栈。装机时 `python -m pip install numpy scipy matplotlib openpyxl`，
或删掉 `<会话目录>\_labreport\.venv` 后重新 init（会让插件重挑解释器）。
也可以在 `cordis.patch.yml` 里显式指定 `pythonPath`。

## Word / PDF

**PDF 没生成，output 里只有 `*-转PDF.ps1`**

- 先双击那个 `.ps1`（右键「使用 PowerShell 运行」），能出 PDF 就是自动化被挡了，不影响交付。
- 父进程若是 Python 起 Word，`Documents.Open()` 永久挂住——插件已经把导出放在 Node 侧，
  别把导出逻辑挪回 Python。
- Word 弹「受保护视图」：模板带 MOTW。右键文件 → 属性 → 解除锁定，或删
  `Zone.Identifier` 备用数据流。
- Word 弹崩溃恢复：删注册表 `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency` 下
  的 `DocumentRecovery`，并确认没有 `~$xxx.docx` 锁文件。

**`.doc` 模板转换失败**

Node 起的是 Windows PowerShell 5.1，脚本必须带 UTF-8 BOM 写盘（中文路径），
`lib/doc2docx.js` 已处理；若仍失败，先用 Word 手动另存为 `.docx`，再用 `template` 参数指定。

## 排版

**图片被页顶裁掉 / 图片重叠**

正文是固定 20 磅行距，内联图片比行高大就会被顶出页面。插件把「图 + 图题」放进 1×1 无框表格
（`cantSplit`），并按 `maxHeightCm` 等比缩图——如果你手工改了样式，注意别把段落行距改回固定值。

**表头字号不对（比如变成 18 pt）**

模板里填空位的 run 自带大字号，填值时不能沿用首 run 格式。`fill_header` 用「标签 run 的字号 →
段落样式 → 10.5」定字号；自定义模板请让标签与填空位同字号。

**怎么确认排版是对的**

用 pymupdf 取坐标，不要靠视觉模型：

```python
import fitz
doc = fitz.open(r"output\xxx.pdf")
for i, page in enumerate(doc, 1):
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                print(i, round(s["bbox"][1], 1), round(s["bbox"][2], 1), s["text"][:30])
```

**原理节总超过 2 页**

13 条行间公式 + 电路图本身就占版面，样例报告同样 3 页以上。`check` 只统计正文文字
（公式与图另计），超了给 warn 不阻断。

## 数据

**照片里的数据怎么进报告**

`data/` 里只有照片时：模型先 `read_image` 逐格抄成数值 → `labreport_data` 落成
`output/<学生>-<实验>-实验数据.xlsx` → **把表格贴给用户确认（单位 / 量程 / 有效数字 / 缺位）**
→ 确认后才写 `compute.py`，脚本直接读这份 xlsx。

**数据里有异常点**

不要自己删。先问用户：保留并说明理由，还是标注后剔除；两种口径的数值都算一遍，
在正文与讨论里写清楚取了哪条、为什么。
