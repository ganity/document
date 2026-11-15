import type { Plugin } from 'vite';

export function systemFontCatalog(
  source: string,
  directory: string,
): {
  source: string;
  available: Map<string, string>;
};

export function localSystemFonts(root: string): Plugin;
