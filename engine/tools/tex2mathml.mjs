// TeX -> MathML，供公式转 OMML 使用。用法：node tex2mathml.mjs <in.json> <out.json>
// in:  [{ "id": "eq01", "tex": "a=b", "display": true }, ...]
// out: [{ "id": "eq01", "ok": true, "mathml": "<math ...>" } | { "id": ..., "ok": false, "error": "..." }]
import { readFileSync, writeFileSync } from 'node:fs';
import temml from '../../vendor/temml.mjs';

const [, , inPath, outPath] = process.argv;
if (!inPath || !outPath) {
  console.error('usage: node tex2mathml.mjs <in.json> <out.json>');
  process.exit(2);
}

const items = JSON.parse(readFileSync(inPath, 'utf8').replace(/^\uFEFF/, ''));
const out = items.map((item) => {
  try {
    const mathml = temml.renderToString(String(item.tex ?? ''), {
      displayMode: Boolean(item.display),
      throwOnError: true,
      xml: true,
    });
    return { id: item.id, ok: true, mathml };
  } catch (err) {
    return { id: item.id, ok: false, error: String(err && err.message ? err.message : err) };
  }
});
writeFileSync(outPath, JSON.stringify(out), 'utf8');
