// 路径边界：所有工程目录必须落在 workspaceRoot（或 extraRoots）内。
// 拒绝越界，避免模型传错路径时把文件写到别处。
import { sep, join, resolve as resolvePath, isAbsolute, basename } from 'node:path';
import { realpathSync, statSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';

const WIN = process.platform === 'win32';
const norm = (p) => (WIN ? p.toLowerCase() : p);

function realish(p) {
  try {
    return realpathSync.native(p);
  } catch {
    return p;
  }
}

export function stripTrailingSep(p) {
  if (p.length > 3 && (p.endsWith(sep) || p.endsWith('/'))) return p.replace(/[\\/]+$/, '');
  return p;
}

export function allowedRoots(cfg) {
  // 用户目录是保底根：会话目录拿不到时（例如加载期解析），也不至于把正常的实验目录判成越界。
  // 真正的写入限制由 DSH 的 file policy 负责，这里只拦手滑传错的路径。
  return [cfg.workspaceRoot, homedir(), ...(cfg.extraRoots || [])]
    .filter(Boolean)
    .map((r) => stripTrailingSep(realish(resolvePath(r))));
}

export function isInside(child, root) {
  const c = norm(stripTrailingSep(child));
  const r = norm(stripTrailingSep(root));
  return c === r || c.startsWith(r + (WIN ? '\\' : '/')) || c.startsWith(r + '/');
}

export function ensureInside(cfg, target, label = '路径') {
  const abs = stripTrailingSep(realish(resolvePath(target)));
  const roots = allowedRoots(cfg);
  if (!roots.some((r) => isInside(abs, r))) {
    throw new Error(
      `${label}越界：${abs}\n允许的根目录：${roots.join('、')}\n` +
        `（可在插件 config.workspaceRoot / config.extraRoots 里追加）`
    );
  }
  return abs;
}

export function isDirectory(p) {
  try {
    return statSync(p).isDirectory();
  } catch {
    return false;
  }
}

/** 把用户/模型给的 dir 解析成"实验目录"（不含 _lr 时补 _lr 的父层语义由调用方决定）。 */
export function resolveUserDir(cfg, raw, label = 'dir') {
  if (typeof raw !== 'string' || !raw.trim()) throw new Error(`${label} 不能为空，请传本机绝对路径`);
  const p = raw.trim().replace(/^["']|["']$/g, '');
  const abs = isAbsolute(p) ? resolvePath(p) : resolvePath(join(cfg.workspaceRoot, p));
  return ensureInside(cfg, abs, label);
}

/** 工程根：<实验目录>\_lr；若传入的就是 _lr 目录则原样返回。 */
export function projectDirOf(experimentDir) {
  return basename(stripTrailingSep(experimentDir)).toLowerCase() === '_lr'
    ? stripTrailingSep(experimentDir)
    : join(stripTrailingSep(experimentDir), '_lr');
}

export function requireDir(p, label = '目录') {
  if (!existsSync(p)) throw new Error(`${label}不存在：${p}`);
  if (!isDirectory(p)) throw new Error(`${label}不是目录：${p}`);
  return p;
}

export function requireFile(p, label = '文件') {
  if (!existsSync(p)) throw new Error(`${label}不存在：${p}`);
  if (!statSync(p).isFile()) throw new Error(`${label}不是文件：${p}`);
  return p;
}

/** 把 dir 里给的相对路径解析为绝对路径（相对实验目录）。 */
export function resolveAgainst(baseDir, p) {
  if (!p) return p;
  return isAbsolute(p) ? resolvePath(p) : resolvePath(join(baseDir, p));
}
