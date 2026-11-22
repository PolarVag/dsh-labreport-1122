# 安装与首次运行

## 1. 环境要求

| 项 | 要求 | 说明 |
| --- | --- | --- |
| 系统 | Windows 10/11 | 依赖 Word COM 与 Windows PowerShell 5.1 |
| DSH | 已装 DeepSeek Harness（`desktop` profile） | 插件以 host bundle 形式装入 |
| Node.js | ≥ 20 | 插件运行时；`node` 需在 PATH 里 |
| Microsoft Word | 2016 及以上 | 导出 PDF、把 `.doc` 模板转 `.docx` |
| Python | 3.10–3.14 | 只用来建 venv；**优先自带 numpy / scipy / matplotlib / openpyxl 的解释器**（Anaconda、Miniconda 都行） |

Python 的科学计算栈（numpy / scipy / matplotlib / openpyxl）是**算数脚本**要用的。
venv 用 `--system-site-packages` 创建，会继承基础解释器的 site-packages，所以插件会优先挑
自带这些包的解释器；如果你机器上只有一个「裸」Python，请先装它们：

```powershell
python -m pip install numpy scipy matplotlib openpyxl
```

## 2. 获取代码

```powershell
cd D:\
git clone https://github.com/PolarVag/dsh-labreport-1122.git
```

没有 git 就下载 ZIP 解压，效果一样——插件读的是本地目录。

## 3. 装入 DSH

在 DSH 会话里让模型执行（或自己在工具面板调用）：

```
plugin_manager  action=install_bundle  target=D:\dsh-labreport-1122
```

`target` 必须是 **clone/解压后的绝对路径**。安装是软链：之后改仓库里的源码，插件立即用新代码
（JS 改动需重新加载模块，见第 5 节）。

安装成功后：

- 工具表里出现 `labreport_init` / `labreport_status` / `labreport_compute` /
  `labreport_build` / `labreport_check` / `labreport_data`；
- 斜杠命令 `/labreport` 可用；
- 系统提示里出现「labreport-workflow」流程段。

## 4. 体检

```
/labreport doctor
```

逐项确认：基础 Python、venv、文档依赖（docx / lxml / yaml / pymupdf）、Node、temml、
Word、MML2OMML.XSL、中文字体、工程目录可写。前几项如果是 ✗，第一次跑 `labreport_init`
会自动建 venv 并下载 wheel（约 1–3 分钟，视网速）。

## 5. 改代码后重新加载

Host 插件不会热更新 `index.js` / `lib/*.js`：

```
plugin_manager  action=set_bundle  target=dsh-labreport  enabled=false
plugin_manager  action=set_bundle  target=dsh-labreport  enabled=true
```

只改 `engine/` 下的 Python 不需要重载（每次调用都新起进程）。

## 6. 第一次跑一份报告

在一个实验目录里放好：模板（`.doc`/`.docx`）、指导书 PDF、书写要求 PDF、数据文件、图片，
然后：

```
labreport_init(dir=<实验目录绝对路径>)
```

`init` 会自动发现模板 / 指导书 / 书写要求 / 样例 / 表头信息文件（`学号.txt`、
`注意事项.txt`、`Extra Information.txt` 之类），打印它选中的来源与表头，让你确认。
之后按 `/labreport` 提示的流程走：大纲确认 → 数据（必要时先转 xlsx）确认 → 出稿。

## 7. 卸载

```
plugin_manager  action=remove_bundle  target=dsh-labreport
```
