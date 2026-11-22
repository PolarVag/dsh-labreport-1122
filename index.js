// dsh-labreport — 基础物理实验报告自动生成（host-only 插件）
//
// 分工：Node 侧只做"路径校验 → 起 Python 引擎 → 渲染结果"，
// 一切文档/数据/公式处理都在 engine/ 里，便于单测与脱离 DSH 复用。
import { mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join, resolve as resolvePath, isAbsolute, dirname } from 'node:path';
import { resolveConfig, forSession } from './config.js';
import { resolveUserDir, projectDirOf, requireDir, requireFile, ensureInside, allowedRoots } from './lib/paths.js';
import { runCli, ENGINE_DIR } from './lib/run.js';
import { exportPdf } from './lib/pdf.js';
import { convertDoc } from './lib/doc2docx.js';

export const name = 'labreport';
export const inject = ['tools', 'subprocess'];

const OUTPUT = {
  schema: {
    type: 'object',
    additionalProperties: true,
    required: ['ok', 'summary'],
    properties: {
      ok: { type: 'boolean', description: '是否成功' },
      summary: { type: 'string', description: '给模型/用户看的文本摘要' },
      data: { type: 'object', additionalProperties: true, description: '结构化结果' },
    },
  },
  render: (_args, value) => [
    { type: 'text', text: String(value?.summary || (value?.ok === false ? '失败（无摘要）' : '完成')) },
  ],
};

// ---------------------------------------------------------------------------
// 参数归一
// ---------------------------------------------------------------------------

function asString(v) {
  return typeof v === 'string' && v.trim() ? v.trim().replace(/^["']|["']$/g, '') : '';
}

function optionalPath(cfg, baseDir, v, label) {
  const s = asString(v);
  if (!s) return null;
  const abs = isAbsolute(s) ? resolvePath(s) : resolvePath(join(baseDir, s));
  return requireFile(ensureInside(cfg, abs, label), label);
}

function optionalPathList(cfg, baseDir, list, label) {
  if (!Array.isArray(list) || !list.length) return [];
  return list.map((v, i) => optionalPath(cfg, baseDir, v, `${label}[${i}]`));
}

function serviceOf(cfg, name) {
  try {
    return cfg && cfg._ctx ? cfg._ctx.get?.(name) ?? null : null;
  } catch {
    return null;
  }
}

async function prepareProject(cfg, dirRaw, exec) {
  // 会话目录只有到调用时才知道（插件 apply 是加载期跑的），这里拿一次再重算允许根目录。
  let bound = cfg;
  let sessionDir = null;
  const wd = serviceOf(cfg, 'workingDirectory');
  if (wd && exec?.agent) {
    try {
      sessionDir = await wd.ensure(exec.agent, exec.signal);
      bound = forSession(cfg, sessionDir);
    } catch {
      /* 拿不到就退回加载期配置 */
    }
  }
  const expDir = requireDir(resolveUserDir(bound, dirRaw, 'dir'), '实验目录');
  const projectDir = ensureInside(bound, projectDirOf(expDir), '工程目录');
  // 共用 venv 优先复用已存在的那个（会话目录 / 实验目录的上一层），避免每次重新装依赖。
  if (bound._venvRelative) {
    const candidates = [sessionDir && join(sessionDir, '_labreport', '.venv'), join(dirname(expDir), '_labreport', '.venv')].filter(Boolean);
    const found = candidates.find((p) => existsSync(join(p, 'Scripts', 'python.exe')));
    bound = { ...bound, venvDir: found || candidates[candidates.length - 1] };
  }
  await mkdir(join(projectDir, 'logs'), { recursive: true });
  return { cfg: bound, expDir, projectDir };
}

function coerce(res) {
  const ok = res?.ok !== false;
  const summary = typeof res?.summary === 'string' && res.summary ? res.summary : ok ? '完成' : '失败（引擎未给摘要）';
  return { ok, summary, data: res?.data && typeof res.data === 'object' ? res.data : {} };
}

// ---------------------------------------------------------------------------
// 工具
// ---------------------------------------------------------------------------

function initTool(ctx, cfg) {
  return {
    name: 'labreport_init',
    description:
      '初始化（或刷新）一份基础物理实验报告工程：在实验目录下建 _lr/，自动挑出报告模板、实验指导书 PDF、书写要求 PDF，' +
      '把它们复制进 _lr/book/ 并把 PDF 全文抽成 book/guide.md、book/requirements.md；旧版 .doc 模板会自动用 Word 转成 .docx；' +
      '样例报告（样例\\ 目录或 sample 参数）复制进 _lr/book/sample/ 并抽成文本，供写作时参照写法；' +
      '数据/图片/补充要求复制进 _lr/media/；生成 report.yaml（表头 + 九节 + 校验清单）。' +
      '首次运行会自动建 Python venv 并装依赖（已装好会跳过）。只需要实验目录的绝对路径，其余自动发现。写报告的第一步就是调它。',
    parameters: {
      type: 'object',
      properties: {
        dir: { type: 'string', description: '实验目录绝对路径（如 D:\\物理基础实验\\温度测量）' },
        template: { type: 'string', description: '报告模板 .doc/.docx；省略则在实验目录里按"报告头/模板"自动挑' },
        guide: { type: 'string', description: '实验指导书 PDF；省略自动挑"指导书/实验指导"' },
        requirements: { type: 'string', description: '报告书写要求 PDF；省略自动挑"书写要求/要求"' },
        sample: {
          type: 'string',
          description: '样例报告（docx/pdf/md）；省略则自动找实验目录下的 样例\\ 或 sample\\ 子目录',
        },
        data: { type: 'array', items: { type: 'string' }, description: '数据文件（xlsx/xls/csv/txt/tsv），可多份' },
        images: { type: 'array', items: { type: 'string' }, description: '已有图片（原理图/电路图/实拍），可多份' },
        extras: { type: 'array', items: { type: 'string' }, description: '补充要求（文本或图片，如"进一步的实验要求.png"），可多份' },
        student: {
          type: 'object',
          additionalProperties: true,
          description: '表头学生信息 {name, id, classGroupSeat}；省略则沿用已有 report.yaml（初建时可留空由用户补）',
        },
        experiment: {
          type: 'object',
          additionalProperties: true,
          description: '实验信息 {name, short, date, location, instructor, makeUp}；省略则用目录名与当前日期',
        },
        force: { type: 'boolean', description: '重建 report.yaml（默认保留已有工程里的修改）' },
        bootstrap: { type: 'boolean', description: '是否允许首次运行自动安装 Python 依赖（默认允许）' },
      },
      required: ['dir'],
    },
    timeoutMs: 900000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      const meta = {};
      const student = args.student && typeof args.student === 'object' ? args.student : null;
      if (student) {
        if (asString(student.name)) meta.student = asString(student.name);
        if (asString(student.id)) meta.student_id = asString(student.id);
        if (asString(student.classGroupSeat)) meta.classGroupSeat = asString(student.classGroupSeat);
      }
      const experiment = args.experiment && typeof args.experiment === 'object' ? args.experiment : null;
      if (experiment) {
        const map = {
          name: 'experiment',
          short: 'short',
          date: 'date',
          location: 'place',
          instructor: 'teacher',
          makeUp: 'makeup',
        };
        for (const [from, to] of Object.entries(map)) {
          const v = asString(experiment[from]);
          if (v) meta[to] = v;
        }
      }
      const copy = [
        ...optionalPathList(c, expDir, args.data, 'data'),
        ...optionalPathList(c, expDir, args.images, 'images'),
        ...optionalPathList(
          c,
          expDir,
          Array.isArray(args.extras) ? args.extras : args.extras ? [args.extras] : [],
          'extras',
        ),
      ].filter(Boolean);
      const sample = optionalPath(c, expDir, args.sample, 'sample');
      const request = {
        dir: expDir,
        projectDir,
        template: optionalPath(c, expDir, args.template, 'template'),
        guide: optionalPath(c, expDir, args.guide, 'guide'),
        requirements: optionalPath(c, expDir, args.requirements, 'requirements'),
        sample: sample || undefined,
        meta,
        copy,
        force: args.force === true,
        bootstrap: args.bootstrap !== false && cfg.autoBootstrap !== false,
        roots: allowedRoots(cfg),
      };
      let result = coerce(await runCli(ctx, exec, c, 'init', projectDir, request, { timeoutSec: 900 }));

      // 模板是旧版 .doc：先由 Node 起 Word 转成 docx，再用转换结果重跑一次 init
      const legacy = result?.data?.templateNeedsConvert;
      if (legacy && !request.template) {
        const outDocx = join(projectDir, 'book', 'template-from-doc.docx');
        const conv = await convertDoc(ctx, exec, legacy, outDocx);
        if (conv.ok) {
          const again = { ...request, template: outDocx };
          result = coerce(await runCli(ctx, exec, c, 'init', projectDir, again, { timeoutSec: 900 }));
          result.data = result.data || {};
          result.data.warnings = (result.data.warnings || []).filter((w) => !String(w).includes('旧版'));
          result.summary = `${result.summary}\n模板已从 ${legacy.split(/[\\/]/).pop()} 转成 docx`.replace(/\n警告：.*旧版.*$/m, '');
          result.data.templateConvertedFrom = legacy;
        } else {
          result.data = result.data || {};
          result.data.warnings = [...(result.data.warnings || []), conv.note];
          result.summary = `${result.summary}\n警告：.doc 模板未转换 —— ${conv.note}`;
        }
      }
      return result;
    },
  };
}

function statusTool(ctx, cfg) {
  return {
    name: 'labreport_status',
    description:
      '查看报告工程进度：各节 markdown 是否已写、字数、表格/图/公式数量、上次构建产物、还缺什么。' +
      '不确定下一步做什么时先调它。',
    parameters: {
      type: 'object',
      properties: { dir: { type: 'string', description: '实验目录绝对路径（或工程 _lr 目录）' } },
      required: ['dir'],
    },
    timeoutMs: 120000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      return coerce(await runCli(ctx, exec, c, 'status', projectDir, { dir: expDir, projectDir }, { timeoutSec: 120 }));
    },
  };
}

function computeTool(ctx, cfg) {
  return {
    name: 'labreport_compute',
    description:
      '运行 _lr/calc/compute.py —— 一段你按本实验数据写的算数脚本（读数整理、逐差/最小二乘拟合、A/B 类不确定度、相对误差），' +
      '脚本约定：结果写 calc/result.json（对象），要进报告的表写 calc/tables/<id>.json {caption, columns, rows}，' +
      '图写 calc/figures/<id>.json {path, caption, width}（path 相对工程目录，一般 ../figures/xx.png）。' +
      '返回 result.json 内容、表格/图清单、脚本输出尾巴与缺图告警，用来让用户确认数据结果。' +
      '首次调用若没有 compute.py，传 scaffold=true 会生成模板。',
    parameters: {
      type: 'object',
      properties: {
        dir: { type: 'string', description: '实验目录绝对路径' },
        script: { type: 'string', description: '自定义脚本路径（默认 _lr/calc/compute.py）' },
        scaffold: { type: 'boolean', description: '没有脚本时生成模板' },
        timeoutSec: { type: 'number', description: '超时秒数，默认 300' },
      },
      required: ['dir'],
    },
    timeoutMs: 900000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      const request = {
        dir: expDir,
        projectDir,
        script: optionalPath(c, expDir, args.script, 'script'),
        scaffold: args.scaffold === true,
        timeoutSec: Number(args.timeoutSec) > 0 ? Number(args.timeoutSec) : 300,
      };
      return coerce(await runCli(ctx, exec, c, 'compute', projectDir, request, { timeoutSec: request.timeoutSec + 60 }));
    },
  };
}

function dataTool(ctx, cfg) {
  return {
    name: 'labreport_data',
    description:
      '把整理好的实验数据落成 xlsx 交用户确认：每张表一个 sheet，写进 _lr/output/<学生>-<实验简称>-实验数据.xlsx。' +
      '实验只给了原始记录表（照片/扫描件/手写表）时，先用 read_image 看清读数、逐格抄成 sheets 传进来，' +
      '把生成的表格贴给用户确认单位与有效数字，确认后再写 compute.py 直接读这份 xlsx。' +
      '返回 xlsx 路径与每张表的行列数。',
    parameters: {
      type: 'object',
      properties: {
        dir: { type: 'string', description: '实验目录绝对路径' },
        sheets: {
          type: 'array',
          description: '数据表清单，每张表一个 sheet',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string', description: '工作表名（如 热电偶、铜电阻）' },
              columns: { type: 'array', items: { type: 'string' }, description: '表头（带单位，如 t/℃）' },
              rows: {
                type: 'array',
                description: '数据行，每行一个数组；数字直接写数字',
                items: { type: 'array', items: {} },
              },
              note: { type: 'string', description: '可选，表下注（如 冷端 0 ℃）' },
            },
            required: ['name', 'columns', 'rows'],
          },
        },
        filename: { type: 'string', description: '可选输出文件名（默认 <学生>-<实验简称>-实验数据.xlsx）' },
      },
      required: ['dir', 'sheets'],
    },
    timeoutMs: 180000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      const request = {
        dir: expDir,
        projectDir,
        sheets: Array.isArray(args.sheets) ? args.sheets : [],
        filename: asString(args.filename) || null,
      };
      return coerce(await runCli(ctx, exec, c, 'data', projectDir, request, { timeoutSec: 120 }));
    },
  };
}

function buildTool(ctx, cfg) {
  return {
    name: 'labreport_build',
    description:
      '把工程组装成最终 Word 报告：在模板原有的表头上原地填空（字体、字号、下划线、位置一律沿用模板，正文紧接表头从第一页开始）、' +
      '按 report.yaml 顺序渲染 report-making/sections/*.md 的正文（默认九节）、插入三线表/图与图题/可编辑公式（Word 原生 OMML）、' +
      '加页脚页码，可选导出 PDF；成品写进 _lr/output/。返回各节字数与页数、图表公式计数与警告。',
    parameters: {
      type: 'object',
      properties: {
        dir: { type: 'string', description: '实验目录绝对路径' },
        docxOut: { type: 'string', description: '输出 docx 路径（默认 output/<名称>.docx）' },
        pdf: { type: 'boolean', description: '是否同时导出 PDF（默认 true）' },
        strict: { type: 'boolean', description: '严格模式：缺图/缺公式/空节直接失败' },
      },
      required: ['dir'],
    },
    timeoutMs: 600000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      const request = {
        dir: expDir,
        projectDir,
        docxOut: asString(args.docxOut) || null,
        pdf: args.pdf !== false,
        strict: args.strict === true,
      };
      const res = coerce(await runCli(ctx, exec, c, 'build', projectDir, request, { timeoutSec: 600 }));
      const data = res.data || {};
      if (request.pdf && data.pdfScript && data.pdfTarget && !data.pdf) {
        const got = await exportPdf(ctx, exec, data.pdfScript, data.pdfTarget);
        if (got.ok) {
          data.pdf = data.pdfTarget;
          let pdfLine = '';
          try {
            const info = coerce(
              await runCli(ctx, exec, c, 'pdfinfo', projectDir, { dir: expDir, projectDir, pdf: data.pdfTarget }, { timeoutSec: 120 }),
            );
            if (info.data?.pages) data.pages = info.data.pages;
            pdfLine = info.summary || '';
          } catch {
            /* 页数拿不到不影响交付 */
          }
          res.summary = `${res.summary}${pdfLine ? '\n' + pdfLine : '；PDF 已导出'}`;
        } else if (got.note) {
          data.warnings = [...(data.warnings || []), got.note];
          res.summary = `${res.summary}\nPDF 未导出：${got.note}`;
        }
      }
      return res;
    },
  };
}

function checkTool(ctx, cfg) {
  return {
    name: 'labreport_check',
    description:
      '对照书写要求自检报告：各节齐全、表题在上/图题在下、单位、项目符号、公式是否为可编辑 OMML、' +
      '原理节页数、与指导书照搬检测、相对误差是否给出、文件名规范、Word 打开是否需要修复等。' +
      '返回 pass/fail/warn 清单，交付前必跑。',
    parameters: {
      type: 'object',
      properties: {
        dir: { type: 'string', description: '实验目录绝对路径' },
        file: { type: 'string', description: '要检查的 docx（默认取最近一次 build 的产物）' },
        strict: { type: 'boolean', description: '严格模式：WARN 也计入失败' },
      },
      required: ['dir'],
    },
    timeoutMs: 300000,
    async execute(args, exec) {
      const { cfg: c, expDir, projectDir } = await prepareProject(cfg, args.dir, exec);
      const request = {
        dir: expDir,
        projectDir,
        file: optionalPath(c, expDir, args.file, 'file'),
        strict: args.strict === true,
      };
      return coerce(await runCli(ctx, exec, c, 'check', projectDir, request, { timeoutSec: 300 }));
    },
  };
}

// ---------------------------------------------------------------------------
// 斜杠命令
// ---------------------------------------------------------------------------

const HELP = [
  '实验报告流水线（/labreport）',
  '',
  '  /labreport doctor            检查 Python/venv/Word/公式转换 是否就绪',
  '  /labreport init [目录]       建工程：抽取指导书/要求/样例，生成 report.yaml',
  '  /labreport status [目录]     看进度（各节字数、图表公式、缺口）',
  '  /labreport data [目录]       把原始记录表整理成 output/ 里的实验数据 xlsx',
  '  /labreport compute [目录]    跑 compute.py，出表格/拟合/误差/图',
  '  /labreport build [目录]      组装 docx（默认同时导 PDF）',
  '  /labreport check [目录]      对照书写要求自检',
  '',
  '目录省略时用当前会话工作目录。也可以直接说"帮我写 XX 实验的报告"，模型会按同一流程调用工具。',
].join('\n');

function parseCommandArgs(rawInput, cfg) {
  const parts = String(rawInput || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  const known = ['doctor', 'init', 'status', 'data', 'compute', 'build', 'check'];
  let sub = 'status';
  if (parts.length && known.includes(parts[0].toLowerCase())) sub = parts.shift().toLowerCase();
  const dir = parts.length ? parts.join(' ') : cfg.workspaceRoot;
  return { sub, dir };
}

function formatStatus(res) {
  const d = res?.data || {};
  const lines = [];
  if (res?.summary) lines.push(res.summary);
  if (Array.isArray(d.sections) && d.sections.length) {
    lines.push('');
    lines.push('章节：');
    for (const s of d.sections) {
      lines.push(`  ${s.state === 'ok' ? '✓' : s.state === 'missing' ? '·' : '!'} ${s.title}${s.chars ? ` (${s.chars} 字)` : ''}`);
    }
  }
  if (d.nextStep) {
    lines.push('');
    lines.push(`下一步：${d.nextStep}`);
  }
  return lines.join('\n') || HELP;
}

function registerCommand(ctx, cfg) {
  // commands 是可选服务：只能走 ctx.get，直接读 ctx.commands 在未 inject 时会抛。
  const commands = ctx.get ? ctx.get('commands') : null;
  if (!commands?.register) {
    ctx.logger?.warn?.('[dsh-labreport] commands 服务不可用，跳过 /labreport');
    return null;
  }
  return commands.register({
    name: 'labreport',
    description: '实验报告流水线：init / status / data / compute / build / check / doctor',
    input: { hint: 'init|status|data|compute|build|check|doctor [实验目录]', attachments: false },
    async handler(invocation) {
      const { sub, dir } = parseCommandArgs(invocation?.rawInput, cfg);
      if (!String(invocation?.rawInput || '').trim()) return { kind: 'success', text: HELP };
      try {
        const { expDir, projectDir } = await prepareProject(cfg, dir);
        const res = await runCli(ctx, { signal: invocation?.signal }, cfg, sub === 'doctor' ? 'doctor' : 'status', projectDir, {
          dir: expDir,
          projectDir,
          bootstrap: true,
        }, { timeoutSec: sub === 'doctor' ? 900 : 120 });
        if (sub === 'doctor') return { kind: 'success', text: res?.summary || 'doctor 完成' };
        if (sub === 'status') return { kind: 'success', text: formatStatus(res) };
        const map = {
          init: `请让模型调用 labreport_init（dir=${expDir}）来完成 init。`,
          data: `请让模型调用 labreport_data（dir=${expDir}，sheets=[...]）。`,
          compute: `请让模型调用 labreport_compute（dir=${expDir}）。`,
          build: `请让模型调用 labreport_build（dir=${expDir}）。`,
          check: `请让模型调用 labreport_check（dir=${expDir}）。`,
        };
        return { kind: 'success', text: `${map[sub] || HELP}\n\n当前状态：\n${formatStatus(res)}` };
      } catch (err) {
        return { kind: 'error', text: `/labreport 失败：${err?.message || err}` };
      }
    },
  });
}

// ---------------------------------------------------------------------------
// 系统提示：让模型知道什么时候、按什么顺序用这套工具
// ---------------------------------------------------------------------------

const PROMPT = [
  '## 基础物理实验报告（dsh-labreport）',
  '',
  '用户要写/改基础物理实验报告时，用 labreport_* 工具，不要自己拼 python-docx 脚本、也不要反复试错 docx 二进制。',
  '工程结构：实验目录下 `_lr/`，只有一份 report.yaml 和两个文件夹：',
  '- `_lr/report-making/`：全部中间产物 —— book/（模板、指导书 md、书写要求 md、样例）、media/（图片副本）、data/（数据文件）、',
  '  sections/（正文 md，默认九节）、calc/（compute.py、result.json、tables/、figures/）、figures/、logs/。',
  '- `_lr/output/`：交付物 —— 最终 docx、pdf，以及实验数据 xlsx。',
  '',
  '流程（用户要求"分步确认"，几道确认门必须真的问）：',
  '1. labreport_init：建工程、抽指导书与书写要求全文、生成 report.yaml；要用的数据/图片用 copy 传进去（xlsx/csv/jpg/png 的路径）。',
  '   表头信息（姓名/学号/日期/地点/指导教师/调课）优先从实验目录的说明文本自动读（学号.txt、Extra Information.txt、注意.txt 等），',
  '   读到就用，不要再问用户要一遍；缺项才问。样例报告自动进 book/sample/ 并抽成同名 .md（旧版 .doc 模板会自动转 docx）。',
  '2. 读 book/requirements.md（书写要求）与 book/guide.md（实验指导书），把各节要点写进 sections/*.md 草稿，并在 report.yaml 登记表/图/公式。',
  '   若 book/sample/ 有样例，读它的 .md：只学章节层次、图表格套与口吻，绝不要抄它的数据、结论与心得。',
  '3. 【确认门 1：大纲】用 ask_user_question 把"九节大纲 + 公式清单 + 图表清单"给用户确认，用户点头前不要开始算数据写正文。',
  '4. 【数据】如果实验目录只给了原始记录表（照片/扫描件/手写表格），**先不要算**：用 read_image 看清每一格读数，按',
  '   {name, columns, rows} 传给 labreport_data，生成 _lr/output/ 里的实验数据 xlsx；然后把表格（每张 sheet 的表头与数值）贴给用户，',
  '   请用户确认单位、量程、有效数字、有没有漏位的数。用户确认（或改正）之后，才写 compute.py，且 compute.py 直接读这份 xlsx。',
  '   如果用户本来就给了 xlsx/csv，跳过这一步。绝不 OCR 猜数，也绝不跳过确认直接算。',
  '5. 写 _lr/report-making/calc/compute.py（用 venv 里的 numpy/scipy，环境变量给好了工程路径），调 labreport_compute 跑；',
  '   算数结论写 calc/result.json；要进报告的表写 calc/tables/<id>.json，图写 calc/figures/<id>.json。',
  '   结果里算不出来的数（NaN/inf）会被置成 null 并在返回里点名——先回头核对原始数据，再问用户是不是抄错了，别硬凑。',
  '6. 【确认门 2：数据结果】把表格数值、拟合参数、不确定度/相对误差给用户看，确认无误再写进正文。',
  '7. 补全正文 → labreport_build（表头按模板原格式原地填空、正文从第一页就开始，成品进 output/，同时尝试导 PDF）→ labreport_check，',
  '   按 issues 改完再交付 output/ 里的 docx 与 pdf。',
  '',
  '表头（重要）：build 直接打开模板 docx、在原有的下划线填空位里填值，字体、字号、行距、下划线、左右位置都沿用模板，',
  '正文紧接表头从第 1 页开始写，不插封面页/分页符。不要自己重画表头，也不要给表头另起一页。',
  '',
  '第五节的写法（重要）：这一节把"数据表格"和"数据处理"合在一起，按分实验交替——每个分实验先给它的数据表格，',
  '紧接着给这个分实验自己的处理（逐差/拟合中间量、代入计算、结果、拟合图、该分实验的不确定度），再进入下一个分实验；',
  '全部分实验做完后再给一张结果汇总表和综合误差分析。严禁"所有表格摆一遍、最后统一处理一遍"。',
  '推荐层次：`## 5.1 拉伸法（CCD 法）` 下先 `@table`，接着 `### 5.1.1 逐差与计算`、`### 5.1.2 不确定度`、图1；`## 5.2 …` 换下一个分实验；',
  '末尾 `## 5.4 结果汇总与误差分析`。',
  '章节 md 方言（引擎只认这些，别写别的）：',
  '- 节标题不要写在 md 里：report.yaml 的 sections[].title 会渲染成"一、目的及要求"。子标题用 "## 1.1 热电偶测温原理"。',
  '- 层级段落以 "（1）""（2）" 或 "1.1）" 开头；严禁 "- ""* " 项目符号（书写要求明确不要，check 会报错）。',
  '- 行间公式：一行 `@formula \\frac{U}{I} #(1)`（#(n) 是编号，可省略）；行内公式写 `$...$`；`$$...$$` 也会被识别。',
  '- 表格：一行 `@table 表1 热电偶标定数据`，紧跟若干 `|` 开头的管道行（第一行是表头，不要写 |---| 分隔行，引擎会自动丢掉）。',
  '- 图：`![图1 电动势-温差拟合曲线](figures/emf_fit.png){width=12}`（路径相对工程目录 _lr/：数据图写 `figures/x.png`，实验目录里的图写 `../x.PNG`；宽度单位 cm，可省）。',
  '- 分页：单独一行 `@pagebreak`。',
  '- 表题必须以"表"开头、图题必须以"图"开头（check 会查）；表题在上、图题在下由引擎排版，你只管写题注文字。',
  '',
  '硬规则：',
  '- 手写数据只用 read_image 看清后抄成数值，走 labreport_data 落 xlsx 并经用户确认；绝不把照片/扫描件贴进正文（check 会拦）。',
  '- 数据异常点（末点漏位、跳变）先问用户，按用户答复处理并在正文注明理由，不要自己编。',
  '- 数据必须重新计算并给相对误差或不确定度，不能抄指导书里的结果；实验原理不得整段照抄指导书，且 ≤2 页。',
  '- 思考题与问题讨论必须自己推理，不套模板。',
  '- 拿到旧版十节工程（第五、六节分家）时，按上面的交替写法把内容重排进 05_data.md，再删掉旧的 05_rawdata/06_processing，别把两节混着渲染。',
  '- build 会在插件侧自动调 Word 导出 PDF（约 10 秒）；若仍失败，output/ 里会留 `-转PDF.ps1`，让用户双击即可，不要为了 PDF 反复重试 build。',
].join('\n');

function registerPrompt(ctx) {
  const sp = ctx.get ? ctx.get('systemPrompt') : null;
  if (!sp?.section) return null;
  try {
    return sp.section({ name: 'labreport-workflow', order: 50, text: PROMPT, interpolate: false });
  } catch (err) {
    ctx.logger?.warn?.(`[dsh-labreport] 注册提示段失败：${err?.message || err}`);
    return null;
  }
}

// ---------------------------------------------------------------------------
// apply
// ---------------------------------------------------------------------------

export function apply(ctx, userConfig) {
  const cfg = { ...resolveConfig(ctx, userConfig), _ctx: ctx };
  const tools = ctx.tools || (ctx.get ? ctx.get('tools') : null);
  if (!tools || typeof tools.register !== 'function') {
    ctx.logger?.warn?.('[dsh-labreport] tools 服务不可用，插件未注册');
    return;
  }
  const registerAll = () => {
    const disposers = [];
    for (const def of [
      initTool(ctx, cfg),
      statusTool(ctx, cfg),
      dataTool(ctx, cfg),
      computeTool(ctx, cfg),
      buildTool(ctx, cfg),
      checkTool(ctx, cfg),
    ]) {
      disposers.push(tools.register({ ...def, output: OUTPUT }));
    }
    const cmdDisposer = registerCommand(ctx, cfg);
    if (cmdDisposer) disposers.push(cmdDisposer);
    const promptDisposer = registerPrompt(ctx);
    if (promptDisposer) disposers.push(promptDisposer);
    ctx.logger?.info?.(`[dsh-labreport] 已注册 6 个工具 + /labreport；engine=${ENGINE_DIR}`);
    return () => {
      for (const d of disposers) {
        try {
          typeof d === 'function' && d();
        } catch {
          /* ignore */
        }
      }
    };
  };
  if (typeof ctx.effect === 'function') ctx.effect(registerAll, 'dsh-labreport:tools');
  else registerAll();
}

export { DEFAULTS } from './config.js';
