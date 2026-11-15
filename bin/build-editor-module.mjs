import { build } from 'vite';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { readFile, writeFile } from 'node:fs/promises';
import { workspaceSourceAliases } from './workspace-aliases.mjs';
import { secureEditorHtml } from './editor-csp.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const outDir = resolve(root, process.argv[2] ?? 'dist-module');
await build({
  configFile: false,
  root,
  base: './',
  publicDir: resolve(root, 'public'),
  resolve: { alias: [{ find: '@', replacement: root }, ...workspaceSourceAliases(root)] },
  build: { outDir, rollupOptions: { input: resolve(root, 'module.html') } },
});
const html = resolve(outDir, 'module.html');
await writeFile(html, secureEditorHtml(await readFile(html, 'utf8')));
