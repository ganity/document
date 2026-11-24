import { t } from '@ranuts/shared/i18n';

type Step = {
  type: string;
  text?: string;
  title?: string;
  question?: string;
  durationMs?: number;
  pauseAfterMs?: number;
  stepIndex?: number;
  checkpointKind?: string;
  visualCues?: Array<{ effect?: { type?: string } }>;
};

// Native Host stores elapsed milliseconds directly, unlike cloud-web's encoded field.
export function speechPauseMode(ms = 0): 'none' | 'pause' | 'timed' {
  return ms <= 0 ? 'none' : ms < 3000 ? 'pause' : 'timed';
}

export function speechPauseLabel(ms = 0): string {
  const mode = speechPauseMode(ms);
  return mode === 'none'
    ? ''
    : mode === 'pause'
      ? t('teachingAfterPause')
      : t('teachingAfterTimed', { seconds: ms / 1000 });
}

export function teachingStepTitle(step: Step): string {
  switch (step.type) {
    case 'speech':
      return t('teachingSpeech');
    case 'native_animation':
      return t('teachingAnimation', { step: (step.stepIndex ?? 0) + 1 });
    case 'audio_clip':
      return t('teachingAudio');
    case 'video_clip':
      return t('teachingVideo');
    case 'delay':
      return t('teachingLegacyDelay');
    case 'interaction_checkpoint':
      return t(step.checkpointKind === 'chapter_summary' ? 'teachingExistingQuestion' : 'teachingTeacherPause');
    case 'music_slot':
      return t('teachingMusicPending');
    default:
      return step.type;
  }
}

export function teachingStepSummary(step: Step): string {
  if (step.type === 'speech') return step.text?.trim() || t('teachingEmptySpeech');
  if (step.type === 'audio_clip' || step.type === 'video_clip')
    return (
      step.title ||
      (step.durationMs ? t('teachingMediaDuration', { seconds: step.durationMs / 1000 }) : teachingStepTitle(step))
    );
  if (step.type === 'interaction_checkpoint') return step.question || t('teachingTeacherPauseHint');
  if (step.type === 'delay') return `${(step.durationMs ?? 0) / 1000} s`;
  return teachingStepTitle(step);
}

const icons = {
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7h.01"/>',
  slide: '<path d="M3 3h18M5 3v13h14V3M12 16v5m-4 0h8M10 7l5 3-5 3z"/>',
  lesson: '<rect x="5" y="5" width="16" height="15" rx="2"/><path d="M2 15V2h15m-6 7 6 3-6 3z"/>',
  continue: '<path d="m5 5 9 7-9 7zM18 5v14"/>',
  previous: '<path d="m15 5-7 7 7 7"/>',
  next: '<path d="m9 5 7 7-7 7"/>',
  replay: '<path d="M3 12a9 9 0 1 0 2.6-6.4L3 8m0-5v5h5"/><path d="M12 8v4l3 2"/>',
  forward: '<path d="m3 5 9 7-9 7zm9 0 9 7-9 7z"/>',
  play: '<path d="m8 5 11 7-11 7z"/>',
  pause: '<path d="M8 5v14M16 5v14"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="1"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>',
  refresh: '<path d="M20 7v5h-5M4 17v-5h5M6 6a8 8 0 0 1 13 3M18 18a8 8 0 0 1-13-3"/>',
  undo: '<path d="m8 4-5 5 5 5M3 9h11a6 6 0 0 1 0 12"/>',
  redo: '<path d="m16 4 5 5-5 5M21 9H10a6 6 0 0 0 0 12"/>',
  up: '<path d="m6 14 6-6 6 6"/>',
  down: '<path d="m6 10 6 6 6-6"/>',
  edit: '<path d="m15 4 5 5M4 20l5-1L21 7l-5-5L4 14z"/>',
  delete: '<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13M10 10v6M14 10v6"/>',
  link: '<path d="m10 13 4-4m-6 6-2 2a3.5 3.5 0 0 1-5-5l4-4a3.5 3.5 0 0 1 5 0m4 1 2-2a3.5 3.5 0 0 0-5-5l-4 4a3.5 3.5 0 0 0 0 5"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  save: '<path d="M5 3h12l4 4v14H3V3zM7 3v6h10V3M7 21v-8h10v8"/>',
  canvas: '<rect x="3" y="4" width="18" height="14" rx="2"/><path d="M8 22h8M12 18v4m-2-14 5 3-5 3z"/>',
};

export function decorateTeachingButton(
  button: HTMLButtonElement,
  icon: keyof typeof icons,
  iconOnly = false,
  caption?: string,
) {
  const label = button.textContent ?? '';
  const text = button.ownerDocument.createElement('span');
  text.textContent = caption ?? label;
  const svg = button.ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.6');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = icons[icon];
  button.setAttribute('aria-label', label);
  button.title = label;
  button.replaceChildren(svg, text);
  if (iconOnly) {
    text.className = 'teaching-sr-only';
    button.classList.add('teaching-icon-button');
  }
  return button;
}

/** Resize through the native layout so the canvas never sits behind the teaching panel. */
export function createTeachingPanelLayout(doc: Document) {
  const win = doc.defaultView;
  let visible = false;
  let originalWidth: number | undefined;
  const sync = (open = visible) => {
    visible = open;
    const native = win as (Window & { PE?: { getController(name: string): any } }) | null;
    const right = native?.PE?.getController('RightMenu')?.getView('RightMenu');
    const menu = doc.getElementById('right-menu');
    const content = menu?.querySelector('.right-panel');
    if (!win || !menu || !content || !right?.setInnerWidth) return;
    const styles = win.getComputedStyle(content);
    const padding = (parseFloat(styles.paddingLeft) || 0) + (parseFloat(styles.paddingRight) || 0);
    if (visible && originalWidth === undefined)
      originalWidth = Math.max(220, menu.getBoundingClientRect().width - padding);
    const width = visible ? Math.max(220, Math.min(360, win.innerWidth - 520)) : originalWidth;
    if (width === undefined) return;
    if (Math.abs(menu.getBoundingClientRect().width - padding - width) > 1) {
      right.setInnerWidth(width);
      native?.PE?.getController('Viewport')?.onLayoutChanged('rightmenu');
    }
    if (!visible) originalWidth = undefined;
  };
  const resize = () => sync();
  win?.addEventListener('resize', resize);
  return {
    sync,
    dispose() {
      sync(false);
      win?.removeEventListener('resize', resize);
    },
  };
}

export const teachingFlowStyles = `
  .deckwise-teaching-sidebar {
    --teaching-accent:var(--ran-color-primary,#32735c);
    --teaching-line:var(--border-regular-control,#e2e7e4);
    --teaching-muted:var(--text-secondary,#7b8981);
    --teaching-surface:var(--background-toolbar,#f5f7f6);
    display:flex; flex-direction:column; gap:0; padding:0; overflow:hidden;
    font-size:12px; line-height:1.5; font-family:inherit;
  }
  .deckwise-teaching-sidebar * { box-sizing:border-box; }
  .deckwise-teaching-sidebar [hidden] { display:none!important; }
  .deckwise-teaching-sidebar .teaching-sr-only { position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip-path:inset(50%); white-space:nowrap; }
  .deckwise-teaching-sidebar header { display:flex; flex-wrap:nowrap; align-items:center; gap:2px; padding:18px 16px 4px; flex:none; }
  .deckwise-teaching-sidebar header strong { font-size:16px; letter-spacing:-.4px; flex:1; min-width:0; }
  .deckwise-teaching-sidebar button { display:inline-flex; align-items:center; justify-content:center; gap:6px; margin:0; padding:7px 10px; border:1px solid transparent; border-radius:7px; font:inherit; font-size:12px; line-height:18px; white-space:nowrap; background:transparent; transition:background .15s,border-color .15s; }
  .deckwise-teaching-sidebar button:hover:not(:disabled) { background:var(--teaching-surface); }
  .deckwise-teaching-sidebar button svg { width:16px; height:16px; flex:none; }
  .deckwise-teaching-sidebar button.teaching-icon-button { width:28px; height:28px; padding:6px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar button:disabled { opacity:.35; }
  .deckwise-teaching-sidebar .teaching-primary { background:var(--teaching-accent); color:var(--background-normal,#fff); border-color:var(--teaching-accent); font-weight:600; }
  .deckwise-teaching-sidebar .teaching-primary:hover:not(:disabled) { background:color-mix(in srgb,var(--teaching-accent) 88%,#000); }
  .deckwise-teaching-sidebar .teaching-statusbar { display:flex; align-items:center; gap:4px; padding:4px 16px 10px; flex:none; position:relative; }
  .deckwise-teaching-sidebar .teaching-statusbar [data-status] { flex:1; min-width:0; margin:0; font-size:11px; color:var(--teaching-muted); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .deckwise-teaching-sidebar .teaching-statusbar [data-status][role=alert] { color:#b54735; white-space:normal; }
  .deckwise-teaching-sidebar .teaching-history { display:flex; gap:0; }
  .deckwise-teaching-sidebar .teaching-toolbar { display:flex; align-items:center; gap:2px; flex:none; }
  .deckwise-teaching-sidebar .teaching-toolbar [data-preview-command=lesson] { order:-2; }
  .deckwise-teaching-sidebar .teaching-toolbar [data-preview-command=show] { order:-1; }
  .deckwise-teaching-sidebar .teaching-toolbar .teaching-primary { color:var(--teaching-accent); background:color-mix(in srgb,var(--teaching-accent) 9%,transparent); border-color:transparent; }
  .deckwise-teaching-sidebar .teaching-page-info { flex:none; }
  .deckwise-teaching-sidebar .teaching-page-info summary { list-style:none; display:flex; width:28px; height:28px; padding:6px; cursor:pointer; border-radius:7px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-page-info summary::-webkit-details-marker { display:none; }
  .deckwise-teaching-sidebar .teaching-page-info summary svg { width:16px; height:16px; }
  .deckwise-teaching-sidebar .teaching-page-info[open] summary { color:var(--teaching-accent); background:var(--teaching-surface); }
  .deckwise-teaching-sidebar .teaching-page-info p { position:absolute; z-index:5; top:100%; right:12px; left:12px; margin:0; padding:12px; border:1px solid var(--teaching-line); border-radius:8px; background:var(--background-normal,#fff); box-shadow:0 4px 12px #00000012; font-size:12px; overflow-wrap:anywhere; }
  .deckwise-teaching-sidebar .teaching-flow { overflow:auto; min-height:0; flex:1; padding:14px 12px 20px; background:var(--teaching-surface); border-block:1px solid var(--teaching-line); scrollbar-width:thin; }
  .deckwise-teaching-sidebar .teaching-flow article { position:relative; padding:0; margin:0 0 8px; border:1px solid transparent; border-radius:10px; }
  .deckwise-teaching-sidebar .teaching-flow article[data-selected] { background:var(--background-normal,#fff); border-color:color-mix(in srgb,var(--teaching-accent) 28%,var(--teaching-line)); box-shadow:0 2px 5px #00000004; }
  .deckwise-teaching-sidebar .teaching-flow article[data-playing] { box-shadow:inset 3px 0 var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-step { display:grid; grid-template-columns:22px minmax(0,1fr); align-items:start; gap:10px; width:100%; padding:12px 82px 12px 12px; border:0; border-radius:9px; text-align:left; }
  .deckwise-teaching-sidebar .teaching-step:hover:not(:disabled) { background:color-mix(in srgb,var(--teaching-accent) 4%,transparent); }
  .deckwise-teaching-sidebar .teaching-step-number { display:grid; place-items:center; width:22px; height:22px; border:1px solid var(--teaching-line); border-radius:50%; color:var(--teaching-muted); font-size:10px; background:var(--background-normal,#fff); }
  .deckwise-teaching-sidebar [data-selected] .teaching-step-number { color:var(--background-normal,#fff); background:var(--teaching-accent); border-color:var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-step-copy { min-width:0; padding-top:1px; }
  .deckwise-teaching-sidebar .teaching-step-copy strong { display:block; font-size:12px; font-weight:600; white-space:nowrap; text-overflow:ellipsis; overflow:hidden; }
  .deckwise-teaching-sidebar .teaching-step-summary { display:block; overflow:hidden; white-space:nowrap; text-overflow:ellipsis; color:var(--teaching-muted); font-size:11px; margin-top:4px; }
  .deckwise-teaching-sidebar .teaching-step-badge { display:inline-block; margin:5px 7px 0 0; color:var(--teaching-accent); font-size:10px; }
  .deckwise-teaching-sidebar article[data-selected] .teaching-step-summary,.deckwise-teaching-sidebar article[data-selected] .teaching-step-badge { display:none; }
  .deckwise-teaching-sidebar .teaching-node-controls { display:flex; position:absolute; right:5px; top:10px; gap:0; padding:0; }
  .deckwise-teaching-sidebar .teaching-node-controls button.teaching-icon-button { width:24px; height:24px; padding:4px; }
  .deckwise-teaching-sidebar article:not([data-selected]):not(:hover):not(:focus-within) .teaching-node-controls { visibility:hidden; }
  .deckwise-teaching-sidebar .teaching-node-controls button:last-child:hover:not(:disabled) { color:#b54735; }
  .deckwise-teaching-sidebar article[data-drop-target] { border-top:2px solid var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-inspector { padding:0 14px 16px 44px; overflow:visible; }
  .deckwise-teaching-sidebar textarea,.deckwise-teaching-sidebar input,.deckwise-teaching-sidebar select { border-color:var(--teaching-line); border-radius:6px; font:inherit; font-size:12px; padding:8px; }
  .deckwise-teaching-sidebar .teaching-inspector textarea { min-height:90px; margin:0 0 8px; padding:10px 11px; line-height:1.8; background:var(--teaching-surface); border-color:transparent; }
  .deckwise-teaching-sidebar .teaching-inspector .teaching-speech-text { field-sizing:content; min-height:58px; max-height:240px; height:auto; overflow-y:auto; resize:none; }
  .deckwise-teaching-sidebar .teaching-inspector textarea:focus { outline:1px solid color-mix(in srgb,var(--teaching-accent) 45%,transparent); background:var(--background-normal,#fff); }
  .deckwise-teaching-sidebar .teaching-voice-status { display:flex; align-items:center; justify-content:space-between; gap:8px; }
  .deckwise-teaching-sidebar .teaching-voice-status p { margin:0; font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-voice-status button { padding:2px 0; color:var(--teaching-accent); font-size:11px; }
  .deckwise-teaching-sidebar .teaching-after { display:flex; align-items:center; flex-wrap:wrap; gap:7px; margin:10px 0 0; padding-top:10px; border-top:1px solid var(--teaching-line); font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-after select { width:auto; max-width:110px; margin:0; padding:3px 6px; border:0; background:var(--teaching-surface); color:var(--text-normal,#263d31); }
  .deckwise-teaching-sidebar .teaching-after input { width:62px; padding:3px 6px; margin:0; }
  .deckwise-teaching-sidebar .teaching-after-hint { font-size:10px; line-height:1.6; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-attention { margin-top:14px; padding-top:12px; border-top:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar .teaching-attention > strong { font-weight:500; font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-animation-authoring { border-top:1px solid var(--teaching-line); margin-top:12px; padding-top:10px; }
  .deckwise-teaching-sidebar .teaching-animation-authoring form { background:var(--background-normal,#fff); border:1px solid var(--teaching-line); border-radius:8px; padding:10px; margin:8px 0; }
  .deckwise-teaching-sidebar .teaching-animation-authoring [hidden] { display:none!important; }
  .deckwise-teaching-sidebar .teaching-animation-picker { position:relative; margin:8px 0; }
  .deckwise-teaching-sidebar .teaching-animation-picker summary { cursor:pointer; border:1px solid var(--teaching-line); border-radius:6px; padding:8px; }
  .deckwise-teaching-sidebar .teaching-animation-picker [role=group] { max-height:220px; overflow:auto; border:1px solid var(--teaching-line); border-radius:6px; margin-top:4px; }
  .deckwise-teaching-sidebar .teaching-animation-picker label { display:grid; grid-template-columns:18px 1fr; align-items:center; gap:6px; padding:8px; border-bottom:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar .teaching-animation-picker input[type=checkbox] { width:15px; height:15px; margin:0; }
  .deckwise-teaching-sidebar .teaching-animation-picker small { grid-column:2; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-animation-row { display:flex; flex-wrap:wrap; gap:4px; padding:8px 0; border-bottom:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar .teaching-animation-row .teaching-animation-edit { display:flex; flex-direction:column; align-items:flex-start; gap:3px; min-width:0; padding:5px 8px; position:relative; }
  .deckwise-teaching-sidebar .teaching-animation-edit-label { font-size:10px; color:var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-animation-edit strong { font-size:11px; font-weight:500; color:var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-animation-edit span { display:block; max-width:100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-animation-authoring form > label,.deckwise-teaching-sidebar .teaching-animation-authoring form > div > label { display:block; font-size:11px; color:var(--teaching-muted); margin-top:8px; }
  .deckwise-teaching-sidebar .teaching-animation-authoring .teaching-animation-help { margin:4px 0 10px; font-size:11px; color:var(--teaching-muted); }

  .deckwise-teaching-sidebar .teaching-animation-row > button:first-child { flex:1; text-align:left; white-space:normal; }
  .deckwise-teaching-sidebar .teaching-animation-word { border-left:2px solid var(--teaching-accent); padding-left:8px; }
  .deckwise-teaching-sidebar .teaching-effect-actions { display:flex; flex-wrap:wrap; gap:6px; margin:6px 0 10px; }
  .deckwise-teaching-sidebar .teaching-effect-actions button { flex:1 0 auto; border-color:var(--teaching-line); padding:5px 7px; font-size:11px; }
  .deckwise-teaching-sidebar .teaching-effect-actions .teaching-primary { color:var(--teaching-accent); background:color-mix(in srgb,var(--teaching-accent) 7%,var(--background-normal,#fff)); border-color:color-mix(in srgb,var(--teaching-accent) 20%,var(--teaching-line)); }
  .deckwise-teaching-sidebar .teaching-effect-actions .teaching-primary:hover:not(:disabled) { background:color-mix(in srgb,var(--teaching-accent) 12%,var(--background-normal,#fff)); }
  .deckwise-teaching-sidebar .teaching-attention > p { font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-attention .teaching-attention-hint { display:none; }
  .deckwise-teaching-sidebar [data-cue-id] { display:grid; grid-template-columns:minmax(0,1fr) auto auto auto; align-items:center; gap:3px; padding:7px 0; border-top:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar [data-cue-id] > p { margin:0; font-size:10px; }
  .deckwise-teaching-sidebar [data-cue-id] > p:first-child { color:var(--teaching-accent); grid-column:1 / -1; }
  .deckwise-teaching-sidebar [data-cue-id] > p:nth-child(2) { min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar [data-cue-id] button.teaching-icon-button { width:22px; height:22px; padding:3px; font-size:10px; }
  .deckwise-teaching-sidebar [data-cue-form] { padding:10px; border:1px solid var(--teaching-line); border-radius:7px; background:var(--teaching-surface); margin-bottom:8px; }
  .deckwise-teaching-sidebar .teaching-cue-actions { display:flex; gap:8px; margin-top:12px; }
  .deckwise-teaching-sidebar .teaching-cue-actions button { min-width:64px; min-height:32px; padding:6px 16px; cursor:pointer; }
  .deckwise-teaching-sidebar .teaching-cue-actions .teaching-secondary { border-color:var(--teaching-line); background:var(--background-normal,#fff); }
  .deckwise-teaching-sidebar .teaching-object-choice { display:flex; flex-wrap:wrap; align-items:center; gap:5px; padding:7px 0; border-bottom:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar .teaching-object-choice .teaching-target { display:grid; grid-template-columns:16px 54px minmax(0,1fr); align-items:center; gap:7px; flex:1; min-width:0; padding:0; margin:0; border:0; cursor:pointer; }
  .deckwise-teaching-sidebar .teaching-object-choice .teaching-target > span:last-child { font-size:11px; line-height:1.5; overflow-wrap:anywhere; }
  .deckwise-teaching-sidebar .teaching-object-choice .teaching-target input { width:14px; height:14px; margin:0; }
  .deckwise-teaching-sidebar .teaching-object-choice:has(input:checked) { background:color-mix(in srgb,var(--teaching-accent) 7%,transparent); }
  .deckwise-teaching-sidebar .teaching-object-picture { display:grid; place-items:center; width:54px; height:46px; border:1px solid var(--teaching-line); border-radius:5px; background:#fff; overflow:hidden; }
  .deckwise-teaching-sidebar .teaching-object-picture img { width:100%; height:100%; object-fit:contain; }
  .deckwise-teaching-sidebar .teaching-object-choice > button { flex:none; padding:4px; font-size:10px; color:var(--teaching-accent); }
  .deckwise-teaching-sidebar .teaching-object-choice > small { flex-basis:100%; color:var(--teaching-muted); font-size:10px; }
  .deckwise-teaching-sidebar .teaching-attention fieldset { max-height:280px; padding:8px; }
  .deckwise-teaching-sidebar .teaching-attention fieldset > p { font-size:10px; color:var(--teaching-muted); margin:0 0 5px; }
  .deckwise-teaching-sidebar .teaching-target { font-size:11px; }
  .deckwise-teaching-sidebar .teaching-inspector > details { margin-top:12px; padding-top:10px; border-top:1px solid var(--teaching-line); font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar summary { cursor:pointer; }
  .deckwise-teaching-sidebar .teaching-native-effects { margin-top:12px; padding-top:10px; border-top:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar .teaching-native-effects > p { font-size:10px; color:var(--teaching-muted); line-height:1.6; }
  .deckwise-teaching-sidebar .teaching-native-effects > button { padding:5px 8px; border:1px solid var(--teaching-line); font-size:11px; }
  .deckwise-teaching-sidebar .teaching-media-preview > p { font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar video { width:100%; max-height:180px; background:#111; border-radius:6px; }
  .deckwise-teaching-sidebar audio { width:100%; height:32px; margin-top:8px; }
  .deckwise-teaching-sidebar footer { display:flex; flex-wrap:wrap; align-items:center; gap:10px 14px; flex:none; padding:12px 16px 14px; border:0; }
  .deckwise-teaching-sidebar footer > button { width:100%; border:1px dashed var(--teaching-line); color:var(--teaching-accent); }
  .deckwise-teaching-sidebar footer > details { margin:0; font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar footer > details[open] { flex-basis:100%; }
  .deckwise-teaching-sidebar :focus-visible { outline-color:var(--teaching-accent); }
  @media(prefers-reduced-motion:reduce) { .deckwise-teaching-sidebar button { transition:none; } }
`;
