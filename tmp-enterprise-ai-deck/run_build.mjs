import fs from 'node:fs';

const sourcePath = new URL('./build_deck.mjs', import.meta.url);
const fixedPath = new URL('./build_deck_runtime.mjs', import.meta.url);
let text = fs.readFileSync(sourcePath, 'utf8');
text = text
  .replace(', faArrowRight, faTasks\n', ', faArrowRight\n')
  .replaceAll('faTasks', 'faListCheck')
  .replace("const svg = icon(fa, { styles: { color } }).html.join('');", "const svg = icon(fa, { styles: { color: '#' + color } }).html.join('');");
fs.writeFileSync(fixedPath, text, 'utf8');
await import(fixedPath.href + `?v=${Date.now()}`);
