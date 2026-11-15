import { t } from '@ranuts/shared/i18n';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { readSlideShapeText } from './agent-plugin/slide-text-read';
import { attentionTarget, type PreviewObject } from './deckwise-teaching-preview';

type ObjectDescription = PreviewObject & { name?: string; role?: string; aliases?: string[] };
export type TeachingObjectChoice = {
  objectId: string;
  name: string;
  description: string;
  thumbnail(): string | undefined;
  locate(): void;
};

/** Display saved Chinese semantics or actual slide text; never guess a translation of a scene object. */
export function teachingObjectChoices(page: number, objects: ObjectDescription[]): TeachingObjectChoice[] {
  const api = getEditorApi();
  const logic = Reflect.get(Reflect.get(api ?? {}, 'WordControl') ?? {}, 'm_oLogicDocument');
  const slide = logic?.Slides?.[page - 1];
  return objects.map((object, index) => {
    const getShape = () => attentionTarget(slide?.cSld?.spTree ?? [], objects, object.objectId);
    let text = '',
      kind: 'text' | 'image' | 'shape' = 'shape';
    try {
      const shape = getShape();
      text =
        readSlideShapeText({
          getDocContent: () => Reflect.get(shape, 'getDocContent')?.call(shape),
          getText: () => Reflect.get(shape, 'getText')?.call(shape),
        })?.trim() ?? '';
      kind = text ? 'text' : Reflect.get(shape, 'isImage')?.call(shape) ? 'image' : 'shape';
    } catch {
      /* A temporarily unavailable target keeps its saved description. */
    }
    const chinese = [...(object.aliases ?? []), object.role, object.name].find(
      (value) => value && /[\u3400-\u9fff]/u.test(value),
    );
    const name =
      chinese ||
      (text
        ? t('teachingObjectText', { text: text.slice(0, 45) })
        : t(kind === 'image' ? 'teachingObjectImage' : 'teachingObjectShape', { number: index + 1 }));
    return {
      objectId: object.objectId,
      name,
      description: object.role || object.name || object.sourceObjectName || '',
      thumbnail() {
        try {
          const shape = getShape();
          const image = Reflect.get(shape, 'getBase64Img')?.call(shape);
          return typeof image === 'string' && /^(data:image\/(png|jpeg|webp);base64,|blob:)/u.test(image)
            ? image
            : undefined;
        } catch {
          return undefined;
        }
      },
      locate() {
        const shape = getShape(),
          controller = logic?.GetCurrentController?.();
        if (!controller?.selectObject) throw new Error(t('teachingNativeEditorUnavailable'));
        controller.resetSelection();
        const group = shape.group && Reflect.get(shape.group, 'getMainGroup')?.call(shape.group);
        if (group) {
          controller.selectObject(group, page - 1);
          controller.selection.groupSelection = group;
          group.selectObject(shape, page - 1);
        } else controller.selectObject(shape, page - 1);
        logic.Document_UpdateSelectionState?.();
        controller.updateOverlay?.();
        Reflect.get(api ?? {}, 'asc_UpdateInterfaceState')?.call(api);
      },
    };
  });
}

export function appendTeachingObjectChoice(
  doc: Document,
  container: HTMLElement,
  input: HTMLInputElement,
  object: Pick<TeachingObjectChoice, 'name' | 'description' | 'thumbnail' | 'locate'>,
  onError: (message: string) => void,
) {
  const row = doc.createElement('div');
  row.className = 'teaching-object-choice';
  const label = doc.createElement('label');
  label.className = 'teaching-target';
  const picture = doc.createElement('span');
  picture.className = 'teaching-object-picture';
  picture.setAttribute('aria-hidden', 'true');
  const name = doc.createElement('span');
  name.textContent = object.name;
  label.title = object.description;
  label.append(input, picture, name);
  const locate = doc.createElement('button');
  locate.type = 'button';
  locate.textContent = t('teachingLocateObject');
  locate.addEventListener('click', () => {
    try {
      object.locate();
    } catch (error) {
      onError(String(error));
    }
  });
  row.append(label, locate);
  container.append(row);
  let loaded = false;
  return () => {
    if (loaded) return;
    loaded = true;
    const src = object.thumbnail();
    if (!src) {
      picture.textContent = '▧';
      return;
    }
    const image = doc.createElement('img');
    image.src = src;
    image.alt = '';
    picture.append(image);
  };
}
