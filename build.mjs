// מאחד את התבנית, המנוע והממשק לקובץ HTML אחד לפרסום כ-Artifact ב-claude.ai.
// הרצה: node build.mjs  →  dist/morning-plan.html
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const tpl = readFileSync(new URL('./app/index.html', import.meta.url), 'utf8');
const engine = readFileSync(new URL('./app/engine.js', import.meta.url), 'utf8');
const app = readFileSync(new URL('./app/app.js', import.meta.url), 'utf8');

const guard = (s) => s.replace(/<\/script/gi, '<\\/script');
const scripts = `<script>\n${guard(engine)}\n</script>\n<script>\n${guard(app)}\n</script>`;
const out = tpl.replace('<!--SCRIPTS-->', () => scripts);

mkdirSync(new URL('./dist/', import.meta.url), { recursive: true });
writeFileSync(new URL('./dist/morning-plan.html', import.meta.url), out);
console.log(`dist/morning-plan.html (${(out.length / 1024).toFixed(1)} KB)`);
