import { existsSync, rmSync } from 'node:fs';
import { dirname } from 'node:path';

// 由插件（Node）直接起 PowerShell 跑 Word 导出。
// 实测：同一个脚本如果由 Python 进程去起，Word 会卡死在 Documents.Open；
// 由 Node 直接起则约 10 秒正常导出，故导出动作放在这一层。
export async function exportPdf(ctx, exec, scriptPath, pdfPath, timeoutMs = 120000) {
  const sub = typeof ctx?.get === 'function' ? ctx.get('subprocess') : ctx?.subprocess;
  if (!sub?.spawn) {
    return { ok: false, note: '拿不到 subprocess 服务，PDF 未导出；可双击 out/ 里的 -转PDF.ps1。' };
  }
  if (existsSync(pdfPath)) {
    try {
      rmSync(pdfPath);
    } catch {
      /* 删不掉也让 Word 覆盖 */
    }
  }

  let handle;
  try {
    handle = sub.spawn({
      argv: ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath],
      cwd: dirname(scriptPath),
      stdio: {
        stdin: 'ignore',
        stdout: { maxBytes: 200000 },
        stderr: { maxBytes: 200000 },
      },
      graceMs: 2000,
      signal: exec?.signal,
    });
  } catch (err) {
    return { ok: false, note: `启动 Word 导出失败：${err?.message || err}` };
  }

  let timer;
  const timeout = new Promise((resolve) => {
    timer = setTimeout(() => resolve('timeout'), timeoutMs);
    if (typeof timer?.unref === 'function') timer.unref();
  });
  const outcome = await Promise.race([handle.done.then(() => 'done').catch(() => 'done'), timeout]);
  if (timer) clearTimeout(timer);

  if (outcome === 'timeout') {
    try {
      handle.terminate();
    } catch {
      /* ignore */
    }
    killWord(ctx);
    return {
      ok: false,
      note: `PDF 导出超时（${Math.round(timeoutMs / 1000)}s），Word 可能被别的窗口卡住；docx 已生成，可双击 -转PDF.ps1 重试。`,
    };
  }

  const settled = await Promise.race([
    new Promise((resolve) => setTimeout(() => resolve(false), 3000)),
    Promise.resolve(true),
  ]);
  void settled;
  if (existsSync(pdfPath)) return { ok: true, note: '' };
  killWord(ctx);
  return { ok: false, note: 'Word 没有产出 PDF（docx 已生成）；可双击 -转PDF.ps1 重试。' };
}

function killWord(ctx) {
  try {
    const sub = typeof ctx?.get === 'function' ? ctx.get('subprocess') : ctx?.subprocess;
    if (!sub?.spawn) return;
    const h = sub.spawn({
      argv: ['taskkill', '/F', '/IM', 'WINWORD.EXE'],
      cwd: process.cwd(),
      stdio: { stdin: 'ignore', stdout: { maxBytes: 4096 }, stderr: { maxBytes: 4096 } },
      graceMs: 1000,
    });
    void h.done?.catch?.(() => {});
  } catch {
    /* ignore */
  }
}
