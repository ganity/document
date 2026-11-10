import fs from 'node:fs';
import path from 'node:path';
import { readNames, xorPrefix } from './lib/sfnt.mjs';
import { buildRecord, decode, encode } from './lib/selection-bin.mjs';

// Runtime-only host fonts. Never copy these into public/, dist/ or an export.
const FAMILIES = [
  { names: ['Calibri'], files: ['calibri.ttf', 'calibrii.ttf', 'calibrib.ttf', 'calibriz.ttf'] },
  { names: ['SimSun', '宋体'], files: ['simsun.ttc'] },
  { names: ['SimHei', '黑体'], files: ['simhei.ttf'] },
  { names: ['KaiTi', '楷体', '楷体_GB2312'], files: ['simkai.ttf'] },
  { names: ['FangSong', '仿宋', '仿宋_GB2312'], files: ['simfang.ttf'] },
  { names: ['Microsoft YaHei', '微软雅黑'], files: ['msyh.ttc', null, 'msyhbd.ttc'] },
];

/** Derive a private catalog without changing the redistributable catalog. */
export function systemFontCatalog(source, directory) {
  const arrays = {};
  for (const key of ['__fonts_files', '__fonts_infos']) {
    const marker = `window["${key}"] = [`;
    const start = source.indexOf(marker) + marker.length;
    if (start < marker.length) throw new Error(`Missing font registry: ${key}`);
    const end = source.indexOf('];', start);
    arrays[key] = { start, end, values: JSON.parse(`[${source.slice(start, end)}]`) };
  }
  const files = arrays.__fonts_files.values;
  const infos = arrays.__fonts_infos.values;
  const available = new Map();
  const records = [];
  const replaced = new Set();
  for (const family of FAMILIES) {
    const regular = path.join(directory, family.files[0]);
    if (!fs.existsSync(regular) || !family.names.includes(readNames(fs.readFileSync(regular))[1])) continue;
    const faces = family.files.map((file) => {
      if (!file) return [-1, -1];
      const filename = path.join(directory, file);
      if (!fs.existsSync(filename)) return [-1, -1];
      const bytes = fs.readFileSync(filename);
      const name = readNames(bytes)[1];
      if (!family.names.includes(name)) return [-1, -1];
      const id = `desktop-system-${file}`;
      const position = files.push(id) - 1;
      available.set(id, filename);
      const record = buildRecord(bytes, { path: `/desktop-system-fonts/${file}` });
      records.push(record);
      replaced.add(record.name);
      return [position, 0];
    });
    if (faces[0][0] < 0) continue;
    while (faces.length < 4) faces.push([-1, -1]);
    for (const name of family.names) {
      const row = infos.find((entry) => entry[0] === name);
      if (row) row.splice(1, 8, ...faces.flat());
    }
  }
  const selectionPattern = /window\["g_fonts_selection_bin"\] = "([^"]+)"/;
  const match = source.match(selectionPattern);
  if (!match) throw new Error('Missing font selection registry');
  const selection = decode(match[1]);
  selection.records = selection.records.filter((record) => !replaced.has(record.name)).concat(records);
  let result = source.replace(selectionPattern, `window["g_fonts_selection_bin"] = "${encode(selection)}"`);
  for (const key of ['__fonts_infos', '__fonts_files']) {
    const { start, end, values } = arrays[key];
    result = result.slice(0, start) + '\n' + values.map((row) => JSON.stringify(row)).join(',\n') + '\n' + result.slice(end);
  }
  return { source: result, available };
}

const isLoopbackHost = (host) => /^(127\.0\.0\.1|localhost|\[::1\])(?::\d+)?$/i.test(host || '');

export function localSystemFonts(projectRoot) {
  let catalog;
  const middleware = (root) => (req, res, next) => {
    if (process.env.DECKWISE_SYSTEM_FONTS !== '1' || process.platform !== 'win32') return next();
    const pathname = (req.url || '').split('?')[0];
    const isCatalog = pathname === '/sdkjs/common/AllFonts.js';
    const id = pathname.startsWith('/fonts/desktop-system-') ? pathname.slice('/fonts/'.length) : null;
    if (!isCatalog && !id) return next();
    // A private local adapter: never expose installed fonts to a network bind,
    // cross-origin website, arbitrary path or untrusted Host header.
    if (!isLoopbackHost(req.headers.host) ||
        !['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.localAddress) ||
        (req.headers['sec-fetch-site'] && !['same-origin', 'none'].includes(req.headers['sec-fetch-site'])) ||
        (req.headers.origin && req.headers.origin !== `http://${req.headers.host}`)) {
      res.writeHead(403); res.end(); return;
    }
    if (!['GET', 'HEAD'].includes(req.method || '')) { res.writeHead(405); res.end(); return; }
    try {
      catalog ??= systemFontCatalog(fs.readFileSync(path.join(root, 'sdkjs/common/AllFonts.js'), 'utf8'),
        path.join(process.env.SystemRoot || 'C:/Windows', 'Fonts'));
      const filename = id ? catalog.available.get(id) : null;
      if (id && !filename) { res.writeHead(404); res.end(); return; }
      const body = isCatalog ? catalog.source : xorPrefix(fs.readFileSync(filename));
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('Content-Type', isCatalog ? 'application/javascript' : 'application/octet-stream');
      res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
      res.end(req.method === 'HEAD' ? undefined : body);
    } catch (error) { next(error); }
  };
  return { name: 'desktop-local-system-fonts',
    configureServer(server) { server.middlewares.use(middleware(path.join(projectRoot, 'public'))); },
    configurePreviewServer(server) { server.middlewares.use(middleware(path.resolve(projectRoot, server.config.build.outDir))); },
  };
}
