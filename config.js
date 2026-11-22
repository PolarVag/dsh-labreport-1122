// 配置：全部可被 profile 的 cordis.patch.yml 里的 config 覆盖。
// 这里不用 schemastery（避免给 bundle 引入依赖），只做普通默认值合并。
import { join, isAbsolute } from 'node:path';
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { spawnSync } from 'node:child_process';

const BASE_VENV = join('_labreport', '.venv');

// 算数脚本要用到这些包；venv 用 --system-site-packages 建，所以基础解释器最好自带
const NEEDED = 'import numpy, scipy, matplotlib, openpyxl';

function hasScienceStack(python) {
  try {
    const r = spawnSync(python, ['-c', NEEDED], { timeout: 8000, stdio: 'ignore', windowsHide: true });
    return r.status === 0;
  } catch {
    return false;
  }
}

/**
 * 找一个能用的基础 Python：只用来 `python -m venv` 与首次引导，装完就 re-exec 到 venv。
 * 顺序：显式环境变量 → 常见安装目录（优先自带 numpy/scipy/matplotlib 的那个）→ PATH 里的 python。
 */
export function detectPython() {
  const explicit = ['DSH_LABREPORT_PYTHON', 'LABREPORT_PYTHON', 'PYTHON']
    .map((k) => process.env[k])
    .filter(Boolean);
  if (explicit.length) return explicit[0];

  const cands = [];
  const local = process.env.LOCALAPPDATA;
  if (local) {
    for (const v of ['314', '313', '312', '311', '310']) {
      cands.push(join(local, 'Programs', 'Python', `Python${v}`, 'python.exe'));
    }
  }
  for (const v of ['314', '313', '312', '311', '310']) {
    cands.push(`C:\\Python${v}\\python.exe`);
  }
  cands.push(
    join(homedir(), 'anaconda3', 'python.exe'),
    join(homedir(), 'miniconda3', 'python.exe'),
    join(homedir(), 'Anaconda3', 'python.exe'),
    'C:\\ProgramData\\Anaconda3\\python.exe',
    'C:\\ProgramData\\miniconda3\\python.exe',
    'C:\\Anaconda\\python.exe',
    'C:\\Anaconda3\\python.exe'
  );

  const found = [];
  for (const c of cands) {
    try {
      if (c && existsSync(c)) found.push(c);
    } catch {
      /* 忽略非法路径 */
    }
  }
  // 优先选自带科学计算栈的解释器：venv 继承它的 site-packages，compute 才能直接 import
  for (const p of found) {
    if (hasScienceStack(p)) return p;
  }
  if (found.length) return found[0];
  const fallback = process.platform === 'win32' ? 'python' : 'python3';
  return hasScienceStack(fallback) ? fallback : fallback;
}

export const DEFAULTS = {
  // 基础 Python（只用 stdlib 引导 venv；插件会自己 re-exec 到 venv）。null = 自动探测
  pythonPath: null,
  // 跨工程复用的 venv 目录；相对路径按会话工作目录解析
  venvDir: BASE_VENV,
  // 允许写入的根目录；null = 自动取会话工作目录
  workspaceRoot: null,
  // 额外允许的根目录（例如把工程放到别的盘）
  extraRoots: [],
  // Word 可执行文件，用于 .doc→.docx 与 docx→pdf。null = 让引擎按常见安装位置自动找
  wordPath: null,
  // MML2OMML.XSL（Office 自带）。null = 取 Word 同目录下的那个
  mml2ommlXsl: null,
  // 成图默认宽度与 dpi
  figureWidthCm: 12,
  dpi: 300,
  // 单次 python 任务超时
  timeoutSec: 600,
  // 每个工程保留多少份日志
  keepLogs: 30,
  // 首次 init 自动装依赖（python-docx lxml pymupdf）
  autoBootstrap: true,
  // 依赖安装失败后是否仍允许继续（抽不出 PDF 文本时会降级）
  softBootstrap: true,
};

function pickWorkspaceRoot(ctx, user) {
  if (user.workspaceRoot) return user.workspaceRoot;
  try {
    const wd = ctx.workingDirectory;
    const dir = wd?.defaultDirectory;
    if (typeof dir === 'string' && dir) return dir;
  } catch {
    /* ignore */
  }
  if (process.env.DSH_WORKSPACE) return process.env.DSH_WORKSPACE;
  return process.cwd();
}

/** 合并默认值，并把相对 venvDir 解析成绝对路径。 */
export function resolveConfig(ctx, userConfig) {
  const user = userConfig && typeof userConfig === 'object' ? userConfig : {};
  const cfg = { ...DEFAULTS, ...user };
  const userRoot = typeof user.workspaceRoot === 'string' && user.workspaceRoot ? user.workspaceRoot : null;
  cfg.workspaceRoot = pickWorkspaceRoot(ctx, user);
  cfg._venvRelative = !isAbsolute(cfg.venvDir);
  cfg._userRoot = userRoot;
  if (!isAbsolute(cfg.venvDir)) cfg.venvDir = join(cfg.workspaceRoot, cfg.venvDir);
  if (!Array.isArray(cfg.extraRoots)) cfg.extraRoots = [];
  if (!cfg.pythonPath) cfg.pythonPath = detectPython();
  return cfg;
}

/**
 * 把配置绑到"本次会话真正的工作目录"上。
 * 插件的 apply 只在加载时跑一次，那时会话目录还不可知；工具调用时才知道 exec.agent 对应的目录，
 * 所以 allowedRoots 与 venv 位置必须在这里重算，否则会被判成"越界"。
 */
export function forSession(cfg, sessionDir) {
  if (!sessionDir) return cfg;
  const venvDir = cfg._venvRelative ? join(sessionDir, BASE_VENV) : cfg.venvDir;
  return {
    ...cfg,
    workspaceRoot: cfg._userRoot || sessionDir,
    venvDir,
    sessionDir,
  };
}

/** venv 里的 python（不存在时返回基础 python）。 */
export function venvPython(cfg) {
  const p =
    process.platform === 'win32'
      ? join(cfg.venvDir, 'Scripts', 'python.exe')
      : join(cfg.venvDir, 'bin', 'python');
  return existsSync(p) ? p : cfg.pythonPath;
}

/** 引擎目录（插件包内的 engine/）。 */
export const ENGINE_DIR_FALLBACK = 'engine';
