import { decorateTeachingButton } from './deckwise-teaching-flow';
import { appendTeachingObjectChoice, type TeachingObjectChoice } from './deckwise-teaching-object-choice';
import { t } from '@ranuts/shared/i18n';
import { animationTriggerLabel, type AnimationCue, type SpeechTrigger } from './deckwise-teaching-animation-bindings';

type NativeChoice = {
  index: number;
  objectId?: string;
  effect?: string;
  durationMs: number;
  supported: boolean;
  owner?: string;
};
export function createTeachingAnimationForm(input: {
  doc: Document;
  action: { id: string; text: string; audioAssetId?: string; durationMs?: number; visualCues?: AnimationCue[] };
  editable: boolean;
  selection(): { start: number; end: number } | undefined;
  objects: Array<
    { objectId: string; name: string } & Partial<Pick<TeachingObjectChoice, 'description' | 'thumbnail' | 'locate'>>
  >;
  native(): NativeChoice[];
  preview(): Promise<void>;
  position(): number | undefined;
  audioDuration?(): number | undefined;
  selectedObjects(): string[];
  save(cues: AnimationCue[], moveNative: string[]): void;
  error(message: string): void;
}) {
  const { doc, action } = input;
  const root = doc.createElement('section');
  root.className = 'teaching-animation-authoring';
  const make = <K extends keyof HTMLElementTagNameMap>(tag: K, text?: string) => {
    const node = doc.createElement(tag);
    if (text) node.textContent = text;
    return node;
  };
  const button = (label: string, run: () => void) => {
    const b = make('button', label);
    b.type = 'button';
    b.addEventListener('click', run);
    return b;
  };
  const objectLabel = (id: string | undefined) =>
    input.objects.find((object) => object.objectId === id)?.name ?? t('teachingMotionTarget');
  const effectLabel = (value: string | undefined) =>
    value === 'fadeIn'
      ? t('teachingMotionFadeIn')
      : value === 'fadeOut'
        ? t('teachingMotionFadeOut')
        : value === 'pulse'
          ? t('teachingMotionPulse')
          : value === 'zoomIn'
            ? t('teachingMotionZoomIn')
            : value === 'slideInLeft'
              ? t('teachingMotionSlideLeft')
              : value === 'slideInRight'
                ? t('teachingMotionSlideRight')
                : value === 'gentleRotate'
                  ? t('teachingMotionSpin')
                  : t('teachingNativeMotion');
  const labels = {
    start: t('teachingTriggerStart'),
    end: t('teachingTriggerEnd'),
    with: t('teachingTriggerWith'),
    after: t('teachingTriggerAfter'),
  };
  const summary = make('div');
  summary.className = 'teaching-animation-summary';
  const controls = make('div');
  controls.className = 'teaching-effect-actions';
  const form = make('form');
  form.hidden = true;
  form.dataset.animationForm = '';
  let editing: AnimationCue | undefined;
  let mode: 'add' | 'link' = 'add';
  let selectedText: { start: number; end: number } | undefined;
  let markedStart: number | undefined, markedEnd: number | undefined;
  let choices: NativeChoice[] = [];
  const position = make('select');
  position.setAttribute('aria-label', t('teachingTriggerPosition'));
  for (const [value, label] of [
    ['start', labels.start],
    ['end', labels.end],
    ['text', t('teachingTriggerText')],
  ])
    position.append(new Option(label, value));
  const word = make('div');
  word.className = 'teaching-animation-word';
  const quote = make('p');
  const phase = make('select');
  phase.setAttribute('aria-label', t('teachingTriggerPhase'));
  phase.append(new Option(labels.with, 'with'), new Option(labels.after, 'after'));
  const marks = make('p');
  marks.setAttribute('role', 'status');
  const target = make('input');
  target.type = 'hidden';
  target.value = input.objects[0]?.objectId ?? '';
  const objectPicker = make('details');
  objectPicker.className = 'teaching-animation-picker teaching-object-picker';
  const objectSummary = make('summary', objectLabel(target.value));
  const objectList = make('div');
  objectList.setAttribute('role', 'group');
  const loadObjects: Array<() => void> = [];
  const objectInputs = input.objects.map((object) => {
    const radio = make('input');
    radio.type = 'radio';
    radio.name = `object-${action.id}`;
    radio.value = object.objectId;
    radio.checked = object.objectId === target.value;
    loadObjects.push(
      appendTeachingObjectChoice(
        doc,
        objectList,
        radio,
        {
          ...object,
          description: object.description ?? object.name,
          thumbnail: object.thumbnail ?? (() => undefined),
          locate: object.locate ?? (() => {}),
        },
        input.error,
      ),
    );
    radio.addEventListener('change', () => {
      target.value = object.objectId;
      objectSummary.textContent = object.name;
      objectPicker.open = false;
    });
    return radio;
  });
  objectPicker.append(objectSummary, objectList);
  objectPicker.addEventListener('toggle', () => {
    if (objectPicker.open) loadObjects.forEach((load) => load());
  });
  const syncObject = () => {
    objectSummary.textContent = objectLabel(target.value);
    objectInputs.forEach((radio) => {
      radio.checked = radio.value === target.value;
    });
  };
  const effect = make('select');
  effect.setAttribute('aria-label', t('teachingAddAnimation'));
  effect.append(new Option(t('teachingMotionFadeIn'), 'fadeIn'), new Option(t('teachingMotionFadeOut'), 'fadeOut'));
  const duration = make('input');
  duration.type = 'number';
  duration.min = '.05';
  duration.max = '5';
  duration.step = '.05';
  duration.value = '.6';
  duration.setAttribute('aria-label', t('teachingAnimationDuration'));
  const field = (labelText: string, control: HTMLElement) => {
    const label = make('label', labelText);
    label.append(control);
    return label;
  };
  const hint = make('p', t('teachingAnimationWhenHint'));
  hint.className = 'teaching-animation-help';
  const addFields = make('div');
  addFields.append(
    make('strong', t('teachingAnimationTargetLabel')),
    target,
    objectPicker,
    button(t('teachingUseCanvasSelection'), () => {
      try {
        const ids = input.selectedObjects();
        if (ids.length !== 1) throw new Error(t('teachingAnimationOneObject'));
        target.value = ids[0];
        syncObject();
      } catch (error) {
        input.error(String(error));
      }
    }),
    field(t('teachingAnimationEffectLabel'), effect),
    field(t('teachingAnimationDurationLabel'), duration),
  );
  const picker = make('details');
  picker.className = 'teaching-animation-picker';
  const pickerLabel = make('summary', t('teachingSelectAnimations'));
  const list = make('div');
  list.setAttribute('role', 'group');
  list.setAttribute('aria-label', t('teachingSelectAnimations'));
  picker.append(pickerLabel, list);
  picker.addEventListener('focusout', () => {
    queueMicrotask(() => {
      if (!picker.contains(doc.activeElement)) picker.open = false;
    });
  });
  picker.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      picker.open = false;
      pickerLabel.focus();
      event.preventDefault();
    }
  });
  const message = make('p');
  message.setAttribute('role', 'status');
  const refresh = () => {
    word.hidden = position.value !== 'text';
    hint.textContent = t(
      position.value === 'start'
        ? 'teachingAnimationStartHint'
        : position.value === 'end'
          ? 'teachingAnimationEndHint'
          : 'teachingAnimationTextHint',
    );
    quote.textContent = selectedText
      ? `「${action.text.slice(selectedText.start, selectedText.end)}」`
      : t('teachingSelectWordsFirst');
    marks.textContent = `${t('teachingMarkStart')}: ${markedStart === undefined ? '—' : (markedStart / 1000).toFixed(2)}s · ${t('teachingMarkEnd')}: ${markedEnd === undefined ? '—' : (markedEnd / 1000).toFixed(2)}s`;
    const selected = [...list.querySelectorAll<HTMLInputElement>('input:checked')];
    pickerLabel.textContent = selected.length
      ? t('teachingSelectedAnimationCount', { count: selected.length })
      : t('teachingSelectAnimations');
    message.textContent = selected.map((box) => choices[Number(box.value)]?.owner).filter(Boolean).length
      ? t('teachingAnimationMoveNotice')
      : '';
  };
  word.append(
    quote,
    phase,
    make('p', t('teachingTimingCalibrationHint')),
    button(t('teachingTimingPreview'), () => {
      void input.preview().catch((error) => input.error(String(error)));
    }),
    button(t('teachingMarkStart'), () => {
      markedStart = input.position();
      refresh();
    }),
    button(t('teachingMarkEnd'), () => {
      markedEnd = input.position();
      refresh();
    }),
    marks,
  );
  position.addEventListener('change', refresh);
  phase.addEventListener('change', refresh);
  const open = (kind: 'add' | 'link', cue?: AnimationCue) => {
    mode = kind;
    editing = cue;
    selectedText =
      cue?.trigger?.kind === 'text' ? { start: cue.trigger.start, end: cue.trigger.end } : input.selection();
    position.value = cue?.trigger?.kind ?? (selectedText ? 'text' : 'start');
    phase.value = cue?.trigger?.kind === 'text' ? cue.trigger.phase : 'with';
    markedStart =
      cue?.trigger?.kind === 'text' && cue.trigger.phase === 'with' && !cue.trigger.stale ? cue.atMs : undefined;
    markedEnd =
      cue?.trigger?.kind === 'text' && cue.trigger.phase === 'after' && !cue.trigger.stale ? cue.atMs : undefined;
    if (cue) {
      target.value = cue.effect.elementId;
      effect.value = cue.effect.effect;
      duration.value = String((cue.endMs - cue.atMs) / 1000);
    }
    syncObject();
    choices = kind === 'link' ? input.native() : [];
    list.replaceChildren();
    choices.forEach((choice, index) => {
      const check = make('input');
      check.type = 'checkbox';
      check.value = String(index);
      check.disabled = !choice.supported;
      check.checked = !!cue?.nativeAnimation && cue.nativeAnimation === choice.objectId;
      const object = input.objects.find((object) => object.objectId === choice.objectId);
      const name = `${objectLabel(choice.objectId)} · ${effectLabel(choice.effect)} · ${t('teachingAnimationSeconds', { seconds: choice.durationMs / 1000 })}`;
      const load = appendTeachingObjectChoice(
        doc,
        list,
        check,
        {
          name,
          description: object?.description ?? object?.name ?? '',
          thumbnail: object?.thumbnail ?? (() => undefined),
          locate: object?.locate ?? (() => {}),
        },
        input.error,
      );
      load();
      const row = list.lastElementChild!;
      if (!choice.supported) row.append(make('small', t('teachingAnimationUnsupported')));
      else if (choice.owner) row.append(make('small', choice.owner));
      check.addEventListener('change', refresh);
    });
    addFields.hidden = kind !== 'add';
    picker.hidden = kind !== 'link';
    picker.open = kind === 'link';
    form.hidden = false;
    submit.textContent = kind === 'link' ? t('teachingLinkAnimation') : t('teachingAnimationSave');
    refresh();
  };
  controls.append(
    button(t('teachingAddAnimation'), () => {
      try {
        open('add');
      } catch (e) {
        input.error(String(e));
      }
    }),
    button(t('teachingLinkAnimation'), () => {
      try {
        open('link');
      } catch (e) {
        input.error(String(e));
      }
    }),
  );
  for (const [i, b] of [...controls.querySelectorAll('button')].entries()) {
    decorateTeachingButton(
      b,
      i === 0 ? 'plus' : 'link',
      false,
      t(i === 0 ? 'teachingAddAnimationShort' : 'teachingLinkAnimationShort'),
    );
    b.disabled = !input.editable || !action.audioAssetId;
  }
  const footer = make('div');
  footer.className = 'teaching-effect-actions';
  const submit = make('button', t('teachingAnimationSave'));
  submit.type = 'submit';
  submit.className = 'teaching-primary';
  footer.append(
    button(t('teachingAnimationCancel'), () => {
      form.hidden = true;
    }),
    submit,
  );
  form.append(field(t('teachingTriggerPosition'), position), hint, word, addFields, picker, message, footer);
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    try {
      if (!input.editable || !action.audioAssetId) throw new Error(t('teachingAnimationAudioRequired'));
      let trigger: SpeechTrigger;
      let atMs = position.value === 'end' ? (action.durationMs ?? 0) : 0;
      if (position.value === 'text') {
        if (!selectedText || !action.text.slice(selectedText.start, selectedText.end).trim())
          throw new Error(t('teachingSelectWordsFirst'));
        const value = phase.value === 'after' ? markedEnd : markedStart;
        if (value === undefined || value < 0 || value > (input.audioDuration?.() ?? action.durationMs ?? 0))
          throw new Error(t('teachingTimingMarkRequired'));
        atMs = value;
        trigger = {
          version: 1,
          kind: 'text',
          ...selectedText,
          text: action.text.slice(selectedText.start, selectedText.end),
          phase: phase.value === 'after' ? 'after' : 'with',
          audioAssetId: action.audioAssetId,
        };
      } else trigger = { version: 1, kind: position.value === 'end' ? 'end' : 'start' };
      const selected =
        mode === 'link'
          ? [...list.querySelectorAll<HTMLInputElement>('input:checked')].map((box) => choices[Number(box.value)])
          : [];
      if (mode === 'link' && !selected.length) throw new Error(t('teachingSelectAnimations'));
      const seconds = Number(duration.value);
      if (mode === 'add' && (!target.value || !Number.isFinite(seconds) || seconds < 0.05 || seconds > 5))
        throw new Error(t('teachingAnimationDuration'));
      const linked = selected.map((choice) => choice.objectId);
      const old = (action.visualCues ?? []).filter(
        (cue) => cue.id !== editing?.id && !linked.includes(cue.nativeAnimation),
      );
      const next: AnimationCue[] =
        mode === 'link'
          ? selected.map((choice) => ({
              id: editing && choice.objectId === editing.nativeAnimation ? editing.id : `motion-${crypto.randomUUID()}`,
              trigger,
              atMs,
              endMs: atMs + choice.durationMs,
              nativeAnimation: choice.objectId!,
              effect: { type: 'element_animation', elementId: choice.objectId!, effect: choice.effect! },
            }))
          : [
              {
                id: editing?.id ?? `motion-${crypto.randomUUID()}`,
                trigger,
                atMs,
                endMs: atMs + Math.round(seconds * 1000),
                effect: { type: 'element_animation', elementId: target.value, effect: effect.value },
              },
            ];
      input.save(
        [...old, ...next].sort(
          (a, b) => (a.trigger?.kind === 'end' ? Infinity : a.atMs) - (b.trigger?.kind === 'end' ? Infinity : b.atMs),
        ),
        selected.map((choice) => choice.objectId!),
      );
      form.hidden = true;
    } catch (error) {
      input.error(String(error));
    }
  });
  for (const cue of action.visualCues ?? []) {
    if (cue.effect?.type !== 'element_animation') continue;
    const row = make('div');
    row.className = 'teaching-animation-row';
    let live: NativeChoice | undefined;
    if (cue.nativeAnimation) {
      try {
        live = input.native().find((choice) => choice.objectId === cue.nativeAnimation);
      } catch {
        /* The saved row remains editable after native changes. */
      }
    }
    const when = animationTriggerLabel(cue.trigger, labels);
    const name = input.objects.find((object) => object.objectId === cue.effect.elementId)?.name ?? '';
    const details = `${objectLabel(cue.effect.elementId)} · ${effectLabel(live?.effect ?? cue.effect.effect)} · ${t('teachingAnimationSeconds', { seconds: Number(((live?.durationMs ?? cue.endMs - cue.atMs) / 1000).toFixed(2)) })}`;
    const edit = button('', () => {
      try {
        open(cue.nativeAnimation ? 'link' : 'add', cue);
      } catch (error) {
        input.error(String(error));
      }
    });
    edit.className = 'teaching-animation-edit';
    edit.append(make('strong', when), make('span', details));
    const editLabel = make('small', t('agentQueueEdit'));
    editLabel.className = 'teaching-animation-edit-label';
    edit.append(editLabel);
    edit.setAttribute('aria-label', `${when} · ${details} · ${t('agentQueueEdit')}`);
    edit.title = `${name} · ${t('agentQueueEdit')}`;
    const remove = button(cue.nativeAnimation ? t('teachingUnlinkAnimation') : t('teachingAnimationRemove'), () =>
      input.save(
        (action.visualCues ?? []).filter((c) => c.id !== cue.id),
        [],
      ),
    );
    row.append(edit, remove);
    if (cue.trigger?.kind === 'text' && (cue.trigger.stale || cue.trigger.audioAssetId !== action.audioAssetId))
      row.append(make('small', t('teachingTimingMarkRequired')));
    for (const b of row.querySelectorAll('button')) b.disabled = !input.editable;
    summary.append(row);
  }
  root.append(controls, form, summary);
  return root;
}
