# 换一个实验怎么写

插件与具体实验无关：实验之间不同的只有三样——**表头从哪来**、**数据怎么算**、**正文怎么写**。
其余（模板比对、九节结构、公式渲染、三线表、图题页码、PDF 导出）都是通用的。

## 一、目录里放什么

```
我的实验/
├─ 实验报告报告头-word.docx      # 模板（.doc 也行，会自动转换）
├─ 实验指导书.pdf                # 关键词：指导书 / 实验指导
├─ 书写要求.pdf                  # 关键词：书写要求 / 要求
├─ 学号.txt / 注意事项.txt      # 表头信息（GBK 也行），可选
├─ 实验数据汇总.xlsx 或 记录表.jpg
├─ 原理图.png / 波形.jpg
└─ 样例\*.docx                   # 可选，学结构口吻，不抄数据
```

文件名只要含对应关键词就能被自动挑中；挑错了就用参数显式指定
（`template` / `guide` / `requirements` / `sample`）。

## 二、表头从哪来

`init` 会依次尝试：实验目录下的文本说明文件（`学号.txt`、`注意事项.txt`、
`Extra Information.txt` 等，键值行如 `指导教师：苏恩曦`）→ 已有的 `report.yaml` → 你显式传的
`student` / `experiment`。**日期若是"今天"会被识别为默认值**，检测到的真实实验日期优先。

九个表头字段（`report.yaml` 的 `meta`）：
`experiment / short / date / place / teacher / makeup / student / student_id / group / grade`。
`init` 会打印它填了什么、来源是哪个文件、还缺什么。

## 三、数据怎么算：`report-making/calc/compute.py`

约定（`labreport_compute` 直接跑它）：

1. 读 `report-making/data/*.xlsx|csv`（Excel 里每张表下面若有一行说明，`read_table` 要能跳过）；
2. 写 `report-making/calc/result.json`：`{"summary": "一句话结论", "notes": [...], ...}`；
3. 要进正文的表写 `report-making/calc/tables/<id>.json`：
   `{"caption": "表3 串联电路幅频与相频数据", "columns": ["f/kHz", "U_pp/V"], "rows": [[...], ...]}`；
4. 要进正文的图写 `report-making/calc/figures/<id>.json`：
   `{"path": "figures/fig_series.png", "caption": "图5 串联电路相频特性", "width": 12}`，
   图片本身存到 `report-making/figures/`（300 dpi + 中文字体）；
5. 用 `os.environ["LR_CALC_DIR"]` / `LR_FIGURES_DIR` / `LR_PROJECT_DIR` 定位，别写死路径；
6. 不确定度：A 类 `s/√n`、B 类 `Δ仪/√3`、方和根合成，最后给**相对误差**。

第一次可以让 `labreport_compute(scaffold=true)` 生成骨架再改。

## 四、正文怎么写：`report-making/sections/*.md`

九节（前四节是说明，后五节是内容）：目的及要求 / 仪器用具 / 实验原理 / 实验内容 /
**获取数据表格与数据处理** / 讨论 / 总结 / 思考题 / 预习报告与原始数据。

**第五节按分实验交替写**（最关键的一条）：

```markdown
## 5.1 拉伸法

### 5.1.1 原始数据

@table from=t_stretch_raw

### 5.1.2 逐差与计算

E = ... 这里写推导

@formula Y=\frac{4MgL}{\pi d^{2}\Delta l} #(1)

### 5.1.3 不确定度

![图1 载荷-伸长量拟合](../figures/fig_stretch.png){width=11}

## 5.2 弯曲法
...
## 5.4 结果汇总与误差分析
```

方言（`engine/lrkit/mdlite.py` 实现，别的写法不认）：

| 写法 | 作用 |
| --- | --- |
| `## 5.1 标题` / `### 5.1.1 标题` | 二、三级标题，编号引擎自动重排 |
| `@table 表1 仪器清单` + 紧跟的 `\|` 管道行（首行表头） | 手写三线表，表题在上 |
| `@table from=t_xxx` | 引用 `calc/tables/t_xxx.json` |
| `![图1 说明](../figures/x.png){width=11}` | 插图，图题在下 |
| `@figure from=fig_xxx` | 引用 `calc/figures/fig_xxx.json` |
| `@formula <LaTeX> #(3)` | 行间编号公式（真 Word 公式，可双击编辑） |
| `$行内公式$` | 行内公式 |
| `@pagebreak` | 分页 |

规则：不用项目符号（用 `（1）` / `1.1`），表格必须有表头行与单位，图题在下表题在上，
原理节控制在 2 页内（公式与图另计），第九节只放预习报告与原始记录的扫描件。

## 五、产物

```
我的实验/_lr/
├─ report.yaml
├─ report-making/            # 中间产物，全部在这
│  ├─ book/                  # 模板、指导书、书写要求、样例、校验清单
│  ├─ media/ data/           # 图与数据
│  ├─ sections/              # 九节正文
│  ├─ calc/ figures/ logs/   # 算数脚本、表/图 JSON、日志
└─ output/                   # 只放交付物：docx、pdf、实验数据 xlsx
```

交付名按书写要求：`班级-分组-座位-姓名-实验名称缩写-实验日期-教师.pdf`。

## 六、参考实例

| 实验 | 特点 |
| --- | --- |
| 温度测量 | 三种温度计（热电偶 / 铜电阻 / 热敏电阻）+ 非平衡电桥设计 + AI 快速测温；异常点保留并说明 |
| 杨氏模量 | 三种方法（拉伸 / 霍尔 / 动态），逐差、作图、外推；静态法系统性偏高 +48% 的误差分析 |
| RLC 谐振与暂态 | 相频过零定 f₀、带宽法求 Q、回路电阻反推；440 mV 异常点两种口径都算 |
| 虚拟仪器 | 原始数据是 `t(ms)-电压(V).txt` 曲线，没有表格照片 |
| 光路偏振 | 数据来自 JPG 总表 + 多个 xlsx |
