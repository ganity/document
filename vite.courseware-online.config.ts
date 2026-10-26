import { defineConfig } from 'vite';

/** Build only the online entry; install alongside the matching existing Office host/plugin resources. */
export default defineConfig({
  base: './',
  publicDir: false,
  build: {
    outDir: 'dist-courseware-online-entry',
    emptyOutDir: true,
    rollupOptions: { input: 'courseware-online.html' },
  },
});
