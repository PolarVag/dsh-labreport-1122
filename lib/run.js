// 起 Python 引擎：写 request JSON → spawn → 读 stdout 的最后一行 JSON。
// 所有重活都在 engine/cli.py，Node 侧只做编排与渲染。
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const PLUGIN_ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
export const ENGINE_DIR = join(PLUGIN_ROOT, 'engine');

const MAX_OUT = 8 * 1024 * 1024;

function stamp() {
  const d = new Date();
  const p = (n, w = 2) => String(n).padStart(w, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

function tail(text, n = 4000) {
  if (!text) return '';
  return text.length <= n ? text : '…' + text.slice(-n);
}

function spawner(ctx) {
  if (ctx?.subprocess?.spawn) return (spec) => ctx.subprocess.spawn(spec);
  const svc = ctx?.get?.('subprocess');
  if (svc?.spawn) return (spec) => svc.spawn(spec);
  throw new Error('subprocess 服务不可用：无法启动 Python 引擎');
}

function readAll(handle, name) {
  const r = handle?.collected?.[name];
  if (!r || typeof r.readFrom !== 'function') return '';
  try {
    return r.readFrom(0).text || '';
  } catch {
    return '';
  }
}

function parsePayload(stdout) {
  const lines = String(stdout || '')
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    if (!line.startsWith('{') || !line.endsWith('}')) continue;
    try {
      return JSON.parse(line);
    } catch {
      /* 继续往前找 */
    }
  }
  return null;
}

/**
 * 运行一个引擎子命令。
 * @param {object} ctx plugin ctx
 * @param {object} exec ToolRunContext（用它的 signal 支持取消）
 * @param {object} cfg 已解析配置
 * @param {string} kind init|doctor|compute|build|check|status
 * @param {string} projectDir 工程根（<实验目录>\_lr）
 * @param {object} request 传给引擎的参数（含 dir 等）
 * @param {{timeoutSec?:number}} opts
 */
export async function runCli(ctx, exec, cfg, kind, projectDir, request, opts = {}) {
  const logsDir = join(projectDir, 'report-making', 'logs');
  await mkdir(logsDir, { recursive: true });
  const reqPath = join(logsDir, `req-${kind}-${stamp()}.json`);
  await writeFile(reqPath, JSON.stringify({ kind, request }, null, 2), 'utf8');

  const timeoutMs = Math.max(10, Number(opts.timeoutSec || cfg.timeoutSec || 600)) * 1000;
  const argv = [cfg.pythonPath, join(ENGINE_DIR, 'cli.py'), kind, '--request', reqPath];
  const env = {
    ...process.env,
    PYTHONUTF8: '1',
    PYTHONIOENCODING: 'utf-8',
    PYTHONPATH: ENGINE_DIR,
    LR_PLUGIN_ROOT: PLUGIN_ROOT,
    LR_VENV_DIR: cfg.venvDir,
    LR_PYTHON: cfg.pythonPath,
    LR_WORD: cfg.wordPath || '',
    LR_MML2OMML: cfg.mml2ommlXsl || '',
    LR_FIG_WIDTH_CM: String(cfg.figureWidthCm ?? 12),
    LR_DPI: String(cfg.dpi ?? 300),
    LR_AUTO_BOOTSTRAP: cfg.autoBootstrap === false ? '0' : '1',
    LR_SOFT_BOOTSTRAP: cfg.softBootstrap === false ? '0' : '1',
    LR_KEEP_LOGS: String(cfg.keepLogs ?? 30),
  };

  const controller = new AbortController();
  const onAbort = () => controller.abort();
  try {
    exec?.signal?.addEventListener?.('abort', onAbort);
  } catch {
    /* ignore */
  }

  const spawn = spawner(ctx);
  let handle;
  try {
    handle = spawn({
      argv,
      cwd: projectDir,
      env,
      stdio: {
        stdin: 'ignore',
        stdout: { maxBytes: MAX_OUT, spill: { maxBytes: MAX_OUT } },
        stderr: { maxBytes: 512 * 1024, spill: { maxBytes: 512 * 1024 } },
      },
      graceMs: 2000,
      signal: controller.signal,
    });
  } catch (err) {
    throw new Error(`启动 Python 引擎失败：${err?.message || err}\npython=${cfg.pythonPath}`);
  }

  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    try {
      handle.terminate();
    } catch {
      /* ignore */
    }
    controller.abort();
  }, timeoutMs);

  const outcome = await Promise.resolve(handle.done).catch((e) => ({ exitCode: -1, error: e }));
  clearTimeout(timer);
  try {
    exec?.signal?.removeEventListener?.('abort', onAbort);
  } catch {
    /* ignore */
  }

  const stdout = readAll(handle, 'stdout');
  const stderr = readAll(handle, 'stderr');

  if (timedOut) {
    throw new Error(
      `引擎命令 ${kind} 超时（${Math.round(timeoutMs / 1000)}s）\n` +
        `可用 config.timeoutSec 调大，或给工具传 timeoutSec。\n` +
        `stderr 末尾：\n${tail(stderr)}`
    );
  }

  const payload = parsePayload(stdout);
  if (payload) {
    // 引擎自己给的错误也带 ok:false，这里原样交给调用方决定
    if (payload.ok === false && payload.fatal) {
      throw new Error(payload.message || `${kind} 失败`);
    }
    return payload;
  }

  const exitCode = outcome?.exitCode;
  throw new Error(
    `引擎命令 ${kind} 未产出 JSON（exit=${exitCode}）\n` +
      `stdout 末尾：\n${tail(stdout, 2000)}\nstderr 末尾：\n${tail(stderr)}`
  );
}
