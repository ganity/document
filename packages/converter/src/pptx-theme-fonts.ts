const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * Adapt MAIC extractThemeEastAsianFont: an explicit EA face wins; otherwise
 * use the source theme's Hans face. The offline x2t reader skips the script
 * collection and resolves an empty +mn-ea to the Latin font instead.
 * Only the transient import theme is normalized; no slide/run is rewritten.
 */
export function resolvePptxThemeEastAsianFonts(xml: string): string | null {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return null;
  let changed = false;
  for (const kind of ['majorFont', 'minorFont']) {
    for (const collection of Array.from(doc.getElementsByTagNameNS(DRAWING_NS, kind))) {
      const children = Array.from(collection.children);
      const ea = children.find((element) => element.namespaceURI === DRAWING_NS && element.localName === 'ea');
      if (!ea || ea.getAttribute('typeface')?.trim()) continue;
      const hans = children.find(
        (element) =>
          element.namespaceURI === DRAWING_NS &&
          element.localName === 'font' &&
          element.getAttribute('script') === 'Hans',
      );
      const face = hans?.getAttribute('typeface')?.trim();
      if (!face) continue;
      ea.setAttribute('typeface', face);
      changed = true;
    }
  }
  return changed ? new XMLSerializer().serializeToString(doc) : null;
}
