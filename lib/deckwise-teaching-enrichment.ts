// Wire contract mirrored from Desktop packages/shared/src/courseware-enrichment.ts.
/** Portable accepted teaching evidence. Raw attachments and local paths stay outside the PPT. */
export type CoursewareEnrichment = {
  schemaVersion: 'deckwise-teaching-enrichment/1';
  originalPages: Array<{ pageId: string; pageNumber: number }>;
  script?: {
    sha256: string;
    policy: 'preserve' | 'revise' | 'generate-missing';
    segments: Array<{
      id: string;
      text: string;
      sourceLine: number;
      sourcePageNumber?: number;
      pageId?: string;
      actionId?: string;
    }>;
    silentPageIds: string[];
  };
  plan?: {
    sha256: string;
    units: Array<{
      id: string;
      title: string;
      evidence: string;
      pageIds: string[];
      kind: 'teaching' | 'quiz' | 'checkpoint';
      activityId?: string;
    }>;
  };
};

const record = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);
const fields = (v: unknown, required: string[], optional: string[] = []): v is Record<string, unknown> =>
  record(v) &&
  required.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => required.includes(k) || optional.includes(k));
const id = (v: unknown): v is string => typeof v === 'string' && /^[\w.-]{1,128}$/u.test(v);
const hash = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/u.test(v);
const integer = (v: unknown, max: number) => Number.isSafeInteger(v) && Number(v) >= 1 && Number(v) <= max;
const text = (v: unknown, max: number) => typeof v === 'string' && !!v.trim() && v.length <= max;
const ids = (v: unknown): v is string[] =>
  Array.isArray(v) && v.length <= 100 && v.every(id) && new Set(v).size === v.length;

export function isCoursewareEnrichment(value: unknown): value is CoursewareEnrichment {
  if (
    !fields(value, ['schemaVersion', 'originalPages'], ['script', 'plan']) ||
    value.schemaVersion !== 'deckwise-teaching-enrichment/1' ||
    new TextEncoder().encode(JSON.stringify(value)).length > 2 * 1024 * 1024 ||
    !Array.isArray(value.originalPages) ||
    !value.originalPages.length ||
    value.originalPages.length > 100 ||
    !value.originalPages.every(
      (p) => fields(p, ['pageId', 'pageNumber']) && id(p.pageId) && integer(p.pageNumber, 100),
    ) ||
    new Set(value.originalPages.map((p) => p.pageId)).size !== value.originalPages.length ||
    new Set(value.originalPages.map((p) => p.pageNumber)).size !== value.originalPages.length
  )
    return false;
  if (value.script !== undefined) {
    const s = value.script;
    if (
      !fields(s, ['sha256', 'policy', 'segments', 'silentPageIds']) ||
      !hash(s.sha256) ||
      !['preserve', 'revise', 'generate-missing'].includes(String(s.policy)) ||
      !ids(s.silentPageIds) ||
      !Array.isArray(s.segments) ||
      !s.segments.length ||
      s.segments.length > 1000 ||
      !s.segments.every(
        (p) =>
          fields(p, ['id', 'text', 'sourceLine'], ['sourcePageNumber', 'pageId', 'actionId']) &&
          id(p.id) &&
          text(p.text, 4000) &&
          integer(p.sourceLine, 100000) &&
          (p.sourcePageNumber === undefined || integer(p.sourcePageNumber, 100)) &&
          (p.pageId === undefined || id(p.pageId)) &&
          (p.actionId === undefined || (id(p.actionId) && id(p.pageId))),
      ) ||
      new Set(s.segments.map((p) => p.id)).size !== s.segments.length ||
      new Set(s.segments.filter((p) => p.actionId).map((p) => `${p.pageId}:${p.actionId}`)).size !==
        s.segments.filter((p) => p.actionId).length ||
      s.segments.some((p) => Array.isArray(s.silentPageIds) && s.silentPageIds.includes(p.pageId))
    )
      return false;
  }
  if (value.plan !== undefined) {
    const p = value.plan;
    if (
      !fields(p, ['sha256', 'units']) ||
      !hash(p.sha256) ||
      !Array.isArray(p.units) ||
      !p.units.length ||
      p.units.length > 100 ||
      !p.units.every(
        (u) =>
          fields(u, ['id', 'title', 'evidence', 'pageIds', 'kind'], ['activityId']) &&
          id(u.id) &&
          text(u.title, 200) &&
          text(u.evidence, 8000) &&
          ids(u.pageIds) &&
          ['teaching', 'quiz', 'checkpoint'].includes(String(u.kind)) &&
          (u.activityId === undefined || id(u.activityId)),
      ) ||
      new Set(p.units.map((u) => u.id)).size !== p.units.length
    )
      return false;
  }
  return true;
}

function pageLabel(line: string): { number: number; end: number } | undefined {
  const normalized = line.replace(/[Ａ-Ｚａ-ｚ０-９]/gu, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  const match =
    /^(?:#{1,6}\s*)?(?:P(?:age)?\s*(\d+)|第\s*([零〇一二两三四五六七八九十百\d]+)\s*页)(?=$|[\s:：、．.]\D|[\s:：、．.]$|[:：])\s*[:：、．.]?\s*/iu.exec(
      normalized,
    );
  if (!match) return;
  const digits = (match[1] ?? match[2]).replace(/〇/gu, '零').replace(/两/gu, '二');
  let number = Number(digits);
  if (!/^\d+$/u.test(digits)) {
    let total = 0,
      digit = 0;
    for (const c of digits) {
      if (c === '十' || c === '百') {
        total += (digit || 1) * (c === '十' ? 10 : 100);
        digit = 0;
      } else digit = '零一二三四五六七八九'.indexOf(c);
    }
    number = total + digit;
  }
  if (!integer(number, 100)) throw new Error('Script page label is outside 1–100');
  return { number, end: match[0].length };
}

/** Strip only structural page headings. Spoken words, decimals and media references remain exact. */
export function parseCoursewareScript(value: string): NonNullable<CoursewareEnrichment['script']>['segments'] {
  if (typeof value !== 'string' || !value.trim() || value.length > 500000 || value.includes('\0'))
    throw new Error('Invalid script text');
  const segments: NonNullable<CoursewareEnrichment['script']>['segments'] = [];
  let page: number | undefined,
    lastPage = 0,
    lineNumber = 0;
  for (const line of value.split(/\r?\n/u)) {
    lineNumber++;
    const label = pageLabel(line);
    if (label) {
      if (label.number < lastPage) throw new Error('Script page labels are out of order');
      page = label.number;
      lastPage = page;
    }
    const body = label ? line.slice(label.end) : line;
    if (!body.trim()) continue;
    // Split long lines without losing punctuation, whitespace or surrogate pairs.
    for (let start = 0; start < body.length;) {
      let end = Math.min(start + 4000, body.length);
      if (end < body.length && /[\uD800-\uDBFF]/u.test(body[end - 1])) end--;
      segments.push({
        id: `segment-${segments.length + 1}`,
        text: body.slice(start, end),
        sourceLine: lineNumber,
        ...(page ? { sourcePageNumber: page } : {}),
      });
      start = end;
    }
  }
  if (!segments.length || segments.length > 1000) throw new Error('Script has no speech or exceeds the segment limit');
  return segments;
}

export type EnrichmentPage = { pageId: string; actions: Array<Record<string, unknown>> };
export function inspectCoursewareEnrichment(enrichment: CoursewareEnrichment | undefined, pages: EnrichmentPage[]) {
  const issues: Array<{ code: string; pageId?: string; sourceId?: string }> = [];
  if (!enrichment) return { ready: true, issues };
  if (!isCoursewareEnrichment(enrichment)) throw new Error('Unsupported or invalid teaching enrichment');
  const script = enrichment.script;
  for (const segment of script?.segments ?? []) {
    const page = pages.find((p) => p.pageId === segment.pageId);
    const action = page?.actions.find((a) => a.id === segment.actionId);
    if (!page || (script?.policy === 'revise' ? !page.actions.some((a) => a.type === 'speech') : !action))
      issues.push({
        code: 'SCRIPT_UNMATCHED',
        sourceId: segment.id,
        ...(segment.pageId ? { pageId: segment.pageId } : {}),
      });
    else if (script?.policy === 'preserve' && (action?.type !== 'speech' || action?.text !== segment.text))
      issues.push({ code: 'SCRIPT_CHANGED', pageId: page.pageId, sourceId: segment.id });
  }
  if (script?.policy === 'preserve')
    for (const original of enrichment.originalPages) {
      if (
        pages.some((p) => p.pageId === original.pageId) &&
        !script.silentPageIds.includes(original.pageId) &&
        !script.segments.some((s) => s.pageId === original.pageId)
      )
        issues.push({ code: 'SCRIPT_PAGE_MISSING', pageId: original.pageId });
    }
  if (script)
    for (const page of pages) {
      if (
        !enrichment.originalPages.some((p) => p.pageId === page.pageId) &&
        !enrichment.plan?.units.some((u) => u.pageIds.includes(page.pageId)) &&
        page.actions.some((a) => a.type === 'speech' && a.source === 'external')
      )
        issues.push({ code: 'SCRIPT_COPY_UNBOUND', pageId: page.pageId });
    }
  for (const pageId of script?.silentPageIds ?? []) {
    const page = pages.find((p) => p.pageId === pageId);
    if (!page || page.actions.some((a) => a.type === 'speech')) issues.push({ code: 'SILENCE_CHANGED', pageId });
  }
  for (const unit of enrichment.plan?.units ?? []) {
    if (!unit.pageIds.length || unit.pageIds.some((id) => !pages.some((p) => p.pageId === id)))
      issues.push({ code: 'PLAN_UNMATCHED', sourceId: unit.id });
    if (
      unit.kind !== 'teaching' &&
      (!unit.activityId ||
        !unit.pageIds.some((id) =>
          pages
            .find((p) => p.pageId === id)
            ?.actions.some((a) => a.id === unit.activityId && a.type === 'interaction_checkpoint'),
        ))
    )
      issues.push({ code: 'ACTIVITY_MISSING', sourceId: unit.id });
  }
  return { ready: !issues.length, issues };
}
