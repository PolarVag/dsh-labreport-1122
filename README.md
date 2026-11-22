# dsh-labreport 实验报告工坊

[![ci](https://github.com/PolarVag/dsh-labreport-1122/actions/workflows/ci.yml/badge.svg)](https://github.com/PolarVag/dsh-labreport-1122/actions/workflows/ci.yml)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

把「实验报告要求 PDF + 实验指导书 PDF + 实验数据/图片 + 报告头模板」变成一份可直接交的
Word 报告（并在本机有 Word 时顺带导出 PDF）。排版细节（页边距 2 cm、宋体小四、行距 20 磅、
三线表、表题在上图题在下、公式用 Word 原生 OMML、页脚「第 N 页」）都按书写要求写死在引擎里，
不需要每次重新摸索。

## 安装与升级

先把仓库 clone 到本机任意位置，然后在 DSH 里装成本地 bundle（安装即软链，改源码立即生效）：

```
git clone https://github.com/PolarVag/dsh-labreport-1122.git
plugin_manager install_bundle  target=<clone 下来的绝对路径>
```

装完让模型跑一次 `/labreport doctor` 或 `labreport_init`，插件会自己建 Python venv
（`<会话工作目录>\_labreport\.venv`）并装 python-docx / lxml / PyYAML / pymupdf。

**环境要求**：Windows + Microsoft Word（2016 及以上，用于 .doc 转换与导出 PDF）+
Python 3.10–3.14（Anaconda 或自带 numpy/scipy/matplotlib/openpyxl 的解释器优先）+ Node ≥ 20。
插件会自动找 Python；找不到或选错时，在 profile 的 `cordis.patch.yml` 里给 `labreport`
配 `pythonPath`。

改了 `index.js` / `lib/*.js` 之后需要重新加载模块（host 插件不会热更新）：

```
plugin_manager set_bundle enabled=false
plugin_manager set_bundle enabled=true
```

只改 `engine/` 下的 Python 不用重载——每次工具调用都会新起进程。

更多见 [`docs/INSTALL.md`](docs/INSTALL.md)、[`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md)、
[`docs/EXPERIMENTS.md`](docs/EXPERIMENTS.md)。

## 目录：一个实验一个工程

```
温度测量\
  _lr\                      ← 工具工作的目录，删掉即可重来
    report.yaml             元数据（表头信息、章节清单）
    report-making\          中间产物都在这儿
      book\                 模板与要求（template.docx / guide.pdf / requirements.md / sample\）
      data\  media\         实验数据与图片（init 时按后缀自动分流：数据进 data，图片进 media）
      sections\01_purpose.md … 九节正文（第五节把"数据表格"和"数据处理"合成一节，按分实验交替写）
      calc\compute.py       算数脚本；产出 result.json、tables/*.json、figures/*.json
      figures\              作图产出的 PNG
      logs\                 每次调用的请求与日志
    output\                 最终交付：报告 docx、报告 pdf、实验数据 xlsx
```

## 五个工具 / 一个命令

| 工具 | 干什么 |
| --- | --- |
| `labreport_init` | 建工程：挑模板（旧版 .doc 自动转 docx）、抽 PDF 文字、复制数据图片、样例入库、生成九节空文件 |
| `labreport_status` | 看进度：哪节还空着、有几张表图、上次构建是什么 |
| `labreport_data` | 把原始记录表（照片里逐格抄出来的数值）落成 `output\<姓名>-<实验>-实验数据.xlsx`，先给用户确认单位与有效数字 |
| `labreport_compute` | 跑 `calc/compute.py`（numpy/scipy 随便用），读数、生成表图 JSON |
| `labreport_build` | 在模板原有表头上原地填空后排版出 docx，正文紧接表头从第 1 页开始（`pdf=true` 时再导 PDF） |
| `labreport_check` | 交前体检：空节、项目符号、表图题、缺不确定度、文件命名 |

`/labreport init|status|data|compute|build|check|doctor [目录]` 是同一套动作的命令行入口
（`doctor` 检查 python / Word / 公式链 / 字体是否齐全）。

## 正文怎么写（md-lite 方言）

只支持实验报告需要的几样东西，写别的也不会生效：

```markdown
## 1.1 冷却定律            ← 二级标题，编号会自动重排成「章号.序号」

普通段落一行一段，**粗体**可用，行内公式写 $T=T_0 e^{-kt}$。

@formula \frac{dQ}{dt}=-k(T-T_0) #(1)     ← 行间公式，末尾 #(n) 是编号

@table 表1 不同时刻的温度
| t/s | 0 | 40 | 80 |
| --- | --- | --- | --- |
| T/℃ | 90.1 | 85.3 | 80.7 |

@table from=tab1          ← 也可以直接引用 compute 写出的表

![图1 冷却曲线](figures/fig1.png){width=12}   ← 图题在前，宽度单位 cm

@pagebreak                ← 强制分页（如正文结束、附原始数据扫描）
```

规矩：**不要写项目符号**（书写要求明确禁止，用「（1）」「1.1」编号）；表题必须以「表」
开头并写在表格上方；图题以「图」开头写在图片下方；原始数据照片只能放在最后一节
（预习报告/原始数据），正文数据必须重新录入表格。

图的高度会被自动限制在约 15.5 cm 内并等比缩宽，且「图 + 图题」放在无框表格里整块挪页——
较高的图不会被页顶裁掉，也不会和图题分居两页。

### 第五节：数据表格与处理按分实验交替

默认结构是**九节**，其中第五节「获取数据表格与数据处理」把原来的「获取数据表格」和
「数据处理及结果」合成一节，写作顺序是"一个分实验一组"：

```markdown
## 5.1 拉伸法（CCD 法）
### 5.1.1 数据表格
@table from=t_stretch_raw
@table from=t_stretch_dim
### 5.1.2 逐差与计算
（处理过程、代入计算、结果）
### 5.1.3 不确定度与结果表达
![图1 砝码质量与读数的关系](figures/fig_stretch.png){width=11}
## 5.2 霍尔法
### 5.2.1 数据表格
…
## 5.4 结果汇总与误差分析
@table from=t_result
```

要点：**每个分实验的表格后面紧跟它自己的处理、拟合图与不确定度**，全部分实验做完后再给
汇总表与综合误差分析；不要"所有表格摆一遍、最后统一处理一遍"。三级标题会自动编号成
`5.1.1`、`5.1.2`，不会和二级标题抢编号。

旧版十节工程（`05_rawdata.md` + `06_processing.md` 分家）在读取 `report.yaml` 时会自动合并成一个
`data` 节，但被并掉的那份文件内容不会自动搬迁——`check` 会提示先手工把它的内容按分实验拆开、
接在对应表格后面，再删掉旧文件。

## 模板与样例

- **模板**：目录里有 `*报告头*.docx`/`*模板*.docx` 就直接用；只有旧版 `.doc` 时，插件会由
  Node 起 Word 把它转成 `book/template-from-doc.docx` 再进流程（Python 侧不碰 Word）。
- **表头**：build 打开 `book/template.docx` 原文件，在原有四段表头里**逐段原地填空**——只替换
  下划线空白 run 的文字，字体、字号、下划线、制表位、行距全部沿用模板；填完删掉表头后的空段落，
  正文（实验题目 + 一、目的及要求 …）紧接表头从**第 1 页**开始，没有封面页、不插分页符。
  模板缺失时才退回自绘表头。
- **样例**：实验目录下的 `样例\`、`sample\`、`范文\` 子目录会被自动收进 `book/sample/`，docx/pdf
  同时抽成同名 `.md`（段落 + 表格）。写正文前读这些 `.md` 学结构、图表格套和口吻，
  **不要抄其中的数据、结论、心得**。也可用 `sample` 参数显式指定单个文件。

## 配置

默认值在 `config.js` 的 `DEFAULTS` 里，要改就在 profile 的 `cordis.patch.yml` 里给
`labreport` 的 `config` 加字段：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `pythonPath` | 自动探测 | 装依赖用的基础解释器；按 `DSH_LABREPORT_PYTHON` → 常见安装路径（优先自带 numpy/scipy/matplotlib/openpyxl 的）→ PATH 顺序找 |
| `venvDir` | `_labreport\.venv` | 依赖装这里（相对会话工作目录；多个工程共享） |
| `wordPath` / `mml2ommlXsl` | 自动探测 | 公式与 PDF 要用；找不到就按常见 Office 路径找 `WINWORD.EXE` 与同目录 `MML2OMML.XSL` |
| `figureWidthCm` / `dpi` | 12 / 300 | 插图默认宽度与出图分辨率 |
| `workspaceRoot` | 当前会话目录 | 允许操作的根目录 |

## 已知坑（都已在代码里处理，别绕过）

- **依赖不走 pip**：某些沙箱环境会把 `tempfile.mkdtemp()` 建的目录判为不可访问，pip 解包必失败；
  `engine/lrkit/pyfetch.py` 直接从 pypi.org 下 wheel 解压到 venv，venv 用
  `--system-site-packages --without-pip` 建。（所以引擎里别用 `import tempfile`，临时文件写
  `工程/report-making/.tmp/`。）
- **Word 只能从 Node 起**：Python 进程里 `Word.Application` 能创建，但
  `Documents.Open()` 会永久挂住；插件侧 `lib/pdf.js` 从 Node 调 powershell 跑
  `output\*-转PDF.ps1`，约 10 秒出 PDF。导出失败不阻断 build，脚本会留在 output/，双击即可。
- **docx 带 MOTW 会卡死在受保护视图**：`init` 会删 `Zone.Identifier`；手动放了新的模板也
  建议右键「解除锁定」。
- **文本编码**：`学号.txt`、`注意事项.txt` 这类文件可能是 GBK；引擎统一用 `read_text_auto`
  （UTF-8 BOM → UTF-8 → GB18030 → UTF-16）读，不要按 UTF-8 硬读。
- **排版别靠肉眼模型判断**：用 pymupdf 取 span 坐标核对（中心、右边距、表宽）；视觉模型会
  把居中的公式说成偏右。

## 测试

```
_labreport\.venv\Scripts\python.exe -m unittest discover -s engine\tests -v
```

覆盖 md-lite 方言解析、init 的模板/样例/表头识别、wheel 挑选（后者的两个历史 bug：
cp314-cp314 被判不兼容、32 位 wheel 抢过 64 位）。
