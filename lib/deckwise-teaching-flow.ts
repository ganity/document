import { t } from '@ranuts/shared/i18n';

type Step = {
  type: string;
  text?: string;
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
      return t('teachingLegacyInteraction');
    default:
      return step.type;
  }
}

export function teachingStepSummary(step: Step): string {
  if (step.type === 'speech') return step.text?.trim() || t('teachingEmptySpeech');
  if (step.type === 'delay') return `${(step.durationMs ?? 0) / 1000} s`;
  return teachingStepTitle(step);
}

const icons = {
  play: '<path d="m8 5 11 7-11 7z"/>',
  pause: '<path d="M8 5v14M16 5v14"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="1"/>',
  refresh: '<path d="M20 7v5h-5M4 17v-5h5M6 6a8 8 0 0 1 13 3M18 18a8 8 0 0 1-13-3"/>',
  undo: '<path d="m8 4-5 5 5 5M3 9h11a6 6 0 0 1 0 12"/>',
  redo: '<path d="m16 4 5 5-5 5M21 9H10a6 6 0 0 0 0 12"/>',
  up: '<path d="m6 14 6-6 6 6"/>',
  down: '<path d="m6 10 6 6 6-6"/>',
  edit: '<path d="m15 4 5 5M4 20l5-1L21 7l-5-5L4 14z"/>',
  delete: '<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13M10 10v6M14 10v6"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  save: '<path d="M5 3h12l4 4v14H3V3zM7 3v6h10V3M7 21v-8h10v8"/>',
  canvas: '<rect x="3" y="4" width="18" height="14" rx="2"/><path d="M8 22h8M12 18v4m-2-14 5 3-5 3z"/>',
};

export function decorateTeachingButton(button: HTMLButtonElement, icon: keyof typeof icons, iconOnly = false) {
  const label = button.textContent ?? '';
  const text = button.ownerDocument.createElement('span');
  text.textContent = label;
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
  .deckwise-teaching-sidebar button { display:inline-flex; align-items:center; justify-content:center; gap:6px; margin:0; padding:7px 10px; border:1px solid transparent; border-radius:7px; font:inherit; font-size:12px; line-height:18px; background:transparent; transition:background .15s,border-color .15s; }
  .deckwise-teaching-sidebar button:hover:not(:disabled) { background:var(--teaching-surface); }
  .deckwise-teaching-sidebar button svg { width:16px; height:16px; flex:none; }
  .deckwise-teaching-sidebar button.teaching-icon-button { width:28px; height:28px; padding:6px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar button:disabled { opacity:.35; }
  .deckwise-teaching-sidebar .teaching-primary { background:var(--teaching-accent); color:var(--background-normal,#fff); border-color:var(--teaching-accent); font-weight:600; }
  .deckwise-teaching-sidebar .teaching-primary:hover:not(:disabled) { background:color-mix(in srgb,var(--teaching-accent) 88%,#000); }
  .deckwise-teaching-sidebar > [data-status] { flex:none; min-height:23px; margin:0; padding:2px 18px 8px; font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar > [data-status][role=alert] { color:#b54735; }
  .deckwise-teaching-sidebar .teaching-history { display:flex; gap:0; }
  .deckwise-teaching-sidebar .teaching-toolbar { display:flex; align-items:center; flex-wrap:wrap; gap:6px; padding:4px 16px 16px; flex:none; }
  .deckwise-teaching-sidebar .teaching-toolbar [data-preview-command=lesson] { order:-2; flex:1; }
  .deckwise-teaching-sidebar .teaching-toolbar [data-preview-command=show] { order:-1; border-color:var(--teaching-line); }
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
  .deckwise-teaching-sidebar .teaching-effect-actions { display:flex; gap:6px; margin:6px 0 10px; }
  .deckwise-teaching-sidebar .teaching-effect-actions button { flex:1; border-color:var(--teaching-line); padding:5px 7px; font-size:11px; }
  .deckwise-teaching-sidebar .teaching-effect-actions .teaching-primary { color:var(--teaching-accent); background:color-mix(in srgb,var(--teaching-accent) 7%,var(--background-normal,#fff)); border-color:color-mix(in srgb,var(--teaching-accent) 20%,var(--teaching-line)); }
  .deckwise-teaching-sidebar .teaching-effect-actions .teaching-primary:hover:not(:disabled) { background:color-mix(in srgb,var(--teaching-accent) 12%,var(--background-normal,#fff)); }
  .deckwise-teaching-sidebar .teaching-attention > p { font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar .teaching-attention .teaching-attention-hint { display:none; }
  .deckwise-teaching-sidebar [data-cue-id] { display:flex; align-items:center; flex-wrap:wrap; gap:2px; padding:4px 0; border-top:1px solid var(--teaching-line); }
  .deckwise-teaching-sidebar [data-cue-id] > p { margin:0; font-size:10px; }
  .deckwise-teaching-sidebar [data-cue-id] > p:first-child { color:var(--teaching-accent); flex:1; }
  .deckwise-teaching-sidebar [data-cue-id] > p:nth-child(2) { position:absolute; width:1px; height:1px; overflow:hidden; clip-path:inset(50%); }
  .deckwise-teaching-sidebar [data-cue-id] button.teaching-icon-button { width:22px; height:22px; padding:3px; font-size:10px; }
  .deckwise-teaching-sidebar [data-cue-form] { padding:10px; border:1px solid var(--teaching-line); border-radius:7px; background:var(--teaching-surface); margin-bottom:8px; }
  .deckwise-teaching-sidebar .teaching-target { font-size:11px; }
  .deckwise-teaching-sidebar .teaching-inspector > details { margin-top:12px; padding-top:10px; border-top:1px solid var(--teaching-line); font-size:11px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar summary { cursor:pointer; }
  .deckwise-teaching-sidebar audio { width:100%; height:32px; margin-top:8px; }
  .deckwise-teaching-sidebar footer { display:flex; flex-wrap:wrap; align-items:center; gap:10px 14px; flex:none; padding:12px 16px 14px; border:0; }
  .deckwise-teaching-sidebar footer > button { width:100%; border:1px dashed var(--teaching-line); color:var(--teaching-accent); }
  .deckwise-teaching-sidebar footer > details { margin:0; font-size:10px; color:var(--teaching-muted); }
  .deckwise-teaching-sidebar footer > details[open] { flex-basis:100%; }
  .deckwise-teaching-sidebar :focus-visible { outline-color:var(--teaching-accent); }
  @media(prefers-reduced-motion:reduce) { .deckwise-teaching-sidebar button { transition:none; } }
`;
