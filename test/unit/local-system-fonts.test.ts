import { readFileSync, mkdtempSync, rmSync, copyFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import vm from 'node:vm';
import { describe, it, expect } from 'vitest';
import { systemFontCatalog, localSystemFonts } from '../../bin/local-system-fonts.mjs';
import { decode } from '../../bin/lib/selection-bin.mjs';

const source = readFileSync(resolve(__dirname, '../../public/sdkjs/common/AllFonts.js'), 'utf8');
const readCatalog = (text: string) => {
  const context = { window: {} as Record<string, any> };
  vm.runInNewContext(text, context);
  return context.window;
};

describe('private local font catalog', () => {
  it('retains all source mappings if host fonts are unavailable', () => {
    const folder = mkdtempSync(join(tmpdir(), 'office-fonts-'));
    try {
      const result = systemFontCatalog(source, folder);
      expect(result.available.size).toBe(0);
      expect(readCatalog(result.source).__fonts_infos).toEqual(readCatalog(source).__fonts_infos);
    } finally { rmSync(folder, { recursive: true }); }
  });

  it('rejects a file whose embedded family is not the requested installed face', () => {
    const folder = mkdtempSync(join(tmpdir(), 'office-fonts-'));
    try {
      copyFileSync(resolve(__dirname, '../../public/fonts/112'), join(folder, 'calibrib.ttf'));
      expect(systemFontCatalog(source, folder).available.size).toBe(0);
    } finally { rmSync(folder, { recursive: true }); }
  });

  it.skipIf(process.platform !== 'win32')('keeps shaping, painting and aliases bound to installed faces', () => {
    const result = systemFontCatalog(source, join(process.env.SystemRoot || 'C:/Windows', 'Fonts'));
    const catalog = readCatalog(result.source);
    const records = decode(catalog.g_fonts_selection_bin).records;
    for (const [name, filename] of [['Calibri', 'calibri.ttf'], ['SimSun', 'simsun.ttc'], ['SimHei', 'simhei.ttf'],
      ['KaiTi', 'simkai.ttf'], ['FangSong', 'simfang.ttf']]) {
      const row = catalog.__fonts_infos.find((entry: any[]) => entry[0] === name);
      if (!result.available.has(`desktop-system-${filename}`)) continue;
      expect(catalog.__fonts_files[row[1]]).toMatch(/^desktop-system-/);
      const selected = records.filter((r: any) => r.name === name);
      expect(selected.length).toBeGreaterThan(0);
      expect(selected.every((r: any) => r.path.startsWith('/desktop-system-fonts/'))).toBe(true);
    }
    expect(catalog.__fonts_infos.find((r: any[]) => r[0] === '宋体').slice(1)).toEqual(catalog.__fonts_infos.find((r: any[]) => r[0] === 'SimSun').slice(1));
    for (const [family, ...aliases] of [['KaiTi', '楷体', '楷体_GB2312'], ['FangSong', '仿宋', '仿宋_GB2312']]) {
      const row = catalog.__fonts_infos.find((r: any[]) => r[0] === family);
      for (const alias of aliases) expect(catalog.__fonts_infos.find((r: any[]) => r[0] === alias).slice(1)).toEqual(row.slice(1));
    }
  });

  it.skipIf(process.platform !== 'win32')('denies cross-origin and network-host access before reading fonts', () => {
    const previous = process.env.DECKWISE_SYSTEM_FONTS;
    process.env.DECKWISE_SYSTEM_FONTS = '1';
    try {
      let handler: any;
      const hook = localSystemFonts('not-a-real-directory').configureServer as (server: any) => void;
      hook({ middlewares: { use: (fn: any) => { handler = fn; } } });
      for (const headers of [{ host: 'evil.test' }, { host: '127.0.0.1:4382', 'sec-fetch-site': 'cross-site' }, { host: 'localhost:4382', origin: 'https://evil.test' }]) {
        let status = 0;
        handler({ url: '/sdkjs/common/AllFonts.js', headers, socket: { localAddress: '127.0.0.1' }, method: 'GET' },
          { writeHead: (code: number) => { status = code; }, end() {} }, () => { throw new Error('must not continue'); });
        expect(status).toBe(403);
      }
    } finally {
      if (previous === undefined) delete process.env.DECKWISE_SYSTEM_FONTS;
      else process.env.DECKWISE_SYSTEM_FONTS = previous;
    }
  });

  it.skipIf(process.platform !== 'win32')('serves installed font bytes only to the same-origin desktop editor', () => {
    const previous = process.env.DECKWISE_SYSTEM_FONTS;
    process.env.DECKWISE_SYSTEM_FONTS = '1';
    try {
      let handler: any;
      const hook = localSystemFonts(resolve(__dirname, '../..')).configureServer as (server: any) => void;
      hook({ middlewares: { use: (fn: any) => { handler = fn; } } });
      let status = 0;
      let body: Buffer | undefined;
      const headers: Record<string, string> = {};
      handler({ url: '/fonts/desktop-system-simsun.ttc', headers: { host: '127.0.0.1:4371', origin: 'http://127.0.0.1:4371', 'sec-fetch-site': 'same-origin' }, socket: { localAddress: '127.0.0.1' }, method: 'GET' },
        { writeHead: (code: number) => { status = code; }, setHeader: (key: string, value: string) => { headers[key] = value; }, end: (data: Buffer) => { body = data; } },
        () => { throw new Error('known system font request must be handled'); });
      expect(status).toBe(0);
      expect(headers['Cross-Origin-Resource-Policy']).toBe('same-origin');
      expect(body && body.length).toBeGreaterThan(1024 * 1024);
    } finally {
      if (previous === undefined) delete process.env.DECKWISE_SYSTEM_FONTS;
      else process.env.DECKWISE_SYSTEM_FONTS = previous;
    }
  });
});
