import { copyFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';

const [destination, pluginsJs] = process.argv.slice(2);
if (!destination || !pluginsJs) {
  throw new Error('Usage: node bin/build-courseware-online.mjs <output directory> <plugins.js from the target Docs image>');
}
const root = fileURLToPath(new URL('../', import.meta.url));
const outDir = resolve(destination);
const modules = join(outDir, 'modules');
await build({
  root,
  configFile: join(root, 'vite.courseware-online.config.ts'),
  build: { outDir, emptyOutDir: false },
});
await build({
  root,
  configFile: false,
  publicDir: false,
  build: {
    outDir: modules,
    emptyOutDir: false,
    lib: {
      entry: join(root, 'lib/docs-community-plugin-entry.ts'),
      formats: ['es'],
      fileName: () => 'docs-community-plugin-entry.js',
    },
  },
});
for (const name of ['index.html', 'config.json']) {
  await copyFile(join(root, 'public/docs-community-plugin', name), join(modules, name));
}
await copyFile(join(root, 'public/docs-online-host.html'), join(outDir, 'docs-online-host.html'));
await copyFile(pluginsJs, join(modules, 'plugins.js'));
