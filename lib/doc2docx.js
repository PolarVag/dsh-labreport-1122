import { existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

// 旧版 .doc 模板先转成 .docx 再进流程（写报告的其余环节只认 OOXML）。
// 与 PDF 导出同理：Word 必须由 Node 起 PowerShell 去驱动，Python 起会卡在 Documents.Open。
export async function convertDoc(ctx, exec, srcDoc, outDocx, timeoutMs = 120000) {
  const sub = typeof ctx?.get === 'function' ? ctx.get('subprocess') : ctx?.subprocess;
  if (!sub?.spawn) return { ok: false, note: '拿不到 subprocess 服务，.doc 模板未转换' };
  if (!existsSync(srcDoc)) return { ok: false, note: `模板不存在：${srcDoc}` };

  // 脚本放工程 logs/，不污染 book/（book 里只放模板与抽取出来的原文）
  const logDir = join(dirname(dirname(outDocx)), 'logs');
  mkdirSync(logDir, { recursive: true });
  const script = join(logDir, 'doc2docx.ps1');
  // 必须带 BOM：Node 起的 powershell 是 Windows PowerShell 5.1，无 BOM 的 UTF-8 会被当 ANSI，
  // 脚本里的中文路径全部变成乱码（和 pdfout.py 写 utf-8-sig 是同一个原因）。
  writeFileSync(script, '\uFEFF' + ps1(srcDoc, outDocx), 'utf8');
  if (existsSync(outDocx)) {
    try {
      rmSync(outDocx);
    } catch {
      /* 让 Word 覆盖 */
    }
  }

  let handle;
  try {
    handle = sub.spawn({
      argv: ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script],
      cwd: dirname(script),
      stdio: { stdin: 'ignore', stdout: { maxBytes: 100000 }, stderr: { maxBytes: 100000 } },
      graceMs: 2000,
      signal: exec?.signal,
    });
  } catch (err) {
    return { ok: false, note: `启动 Word 转换失败：${err?.message || err}` };
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
    await killWord(sub);
    return { ok: false, note: `Word 转换超时（${Math.round(timeoutMs / 1000)}s）` };
  }
  if (existsSync(outDocx)) return { ok: true, note: '', docx: outDocx };
  await killWord(sub);
  return { ok: false, note: 'Word 没有产出 .docx（模板可能是加密/受保护文档）' };
}

function ps1(src, dst) {
  const q = (p) => `'${String(p).replace(/'/g, "''")}'`;
  return `$ErrorActionPreference = 'Stop'
$src = ${q(src)}
$dst = ${q(dst)}
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3
  $doc = $word.Documents.Open($src, $false, $true)
  $doc.SaveAs2($dst, 16)
  $doc.Close($false)
  if (-not (Test-Path -LiteralPath $dst)) { throw 'saveas produced nothing' }
  Write-Output "DOC2DOCX OK $dst"
} finally {
  if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
  if ($word -ne $null) { try { $word.Quit() } catch {} }
}
`;
}

async function killWord(sub) {
  try {
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
