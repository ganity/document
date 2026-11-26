import { canExtractBackgroundRegion, prepareBackgroundRegion } from './agent-plugin/slide-background-region';
import { prepareAssessmentInsertion, type AssessmentProjection } from './agent-plugin/slide-assessment-insertion';
import { prepareQuizProjection } from './agent-plugin/slide-quiz-projection';
import { prepareAssessmentText } from './agent-plugin/slide-assessment-text';
import { renderNativeSlide } from './deckwise-slide-preview';
import { nativeLayerEvidence } from './agent-plugin/slide-layer-evidence';
import { requireEditorContext } from './agent-plugin/editor-bridge';
import { readPresentationShapeText } from './agent-plugin/slide-text-read';
import { getReadonlyMode } from './onlyoffice/readonly';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import {
  insertNativeTeachingAssessment,
  replaceNativeAssessments,
  assertNativeTeachingAnimations,
  nativeAnimationSignatures,
  readNativeTeaching,
  replaceNativeTeachingPage,
  replaceNativeTeachingRegion,
  nativeTeachingAudio,
  setNativeTeachingLock,
  nativeTeachingObjects,
  adoptNativeTeachingAudio,
  confirmNativeTeachingPage,
} from './deckwise-native-teaching';

/** Return undefined only for a legacy document. Native documents never write through the legacy Host. */
export async function routeNativeTeaching(
  operation: string,
  args: Record<string, unknown>,
  pageNumber: number,
  online = false,
): Promise<{ value: unknown } | undefined> {
  const teaching = readNativeTeaching();
  if (!teaching) return;
  const sourceSha256 = deckwisePageLineage.context()?.sourceSha256 ?? '';
  if (sourceSha256)
    for (const current of teaching.pages)
      for (const action of current.actions) if (action.type === 'native_animation') action.sourceSha256 = sourceSha256;
  const page = teaching.pages[pageNumber - 1];
  if (!page) throw new Error('Native teaching page is unavailable');
  const generation =
    teaching.generation && typeof teaching.generation === 'object'
      ? (teaching.generation as Record<string, unknown>)
      : undefined;
  const requirements = Array.isArray(generation?.pages)
    ? generation.pages
        .filter(
          (p: unknown): p is { pageId: string; teaching: unknown } =>
            !!p &&
            typeof p === 'object' &&
            typeof Reflect.get(p, 'pageId') === 'string' &&
            !!Reflect.get(p, 'teaching'),
        )
        .filter((p) => teaching.pages.some((page) => page.pageId === p.pageId))
        .map(({ pageId, teaching }) => ({ pageId, teaching }))
    : [];
  const isDiscussion = (action: any) =>
    !action.response?.kind || ['discussion', 'practice', 'pbl'].includes(action.response.kind);
  // Page purpose comes from saved planning, never from a question mixed into a lecture.
  const pageRole = (p: typeof page): 'quiz' | 'checkpoint' | undefined => {
    const saved = (teaching.assessments as Record<string, unknown> | undefined)?.[p.pageId];
    const planned = Array.isArray(generation?.pages) ? generation.pages.find((g) => g?.pageId === p.pageId) : undefined;
    if (saved || planned?.quiz || planned?.teaching?.pageType === 'quiz') return 'quiz';
    if (['interactive', 'pbl'].includes(planned?.teaching?.pageType)) return 'checkpoint';
    if (planned) return;
    const unit = teaching.enrichment?.plan?.units.find(
      (u) =>
        u.kind !== 'teaching' &&
        u.activityId &&
        u.pageIds.includes(p.pageId) &&
        p.actions.some((a) => a.id === u.activityId),
    );
    if (unit) return unit.kind === 'quiz' ? 'quiz' : 'checkpoint';
    if (teaching.enrichment?.originalPages.some((original) => original.pageId === p.pageId)) return;
    // Older standalone activities have only their question reading and a checkpoint.
    const checkpoints = p.actions.filter(
      (a) => a.type === 'interaction_checkpoint' && a.checkpointKind === 'chapter_summary',
    );
    if (checkpoints.length !== 1) return;
    const checkpoint = checkpoints[0];
    const response = checkpoint.response as { options?: string[] } | undefined;
    const reading = [
      checkpoint.question,
      ...(response?.options ?? []).map((option, i) => `${String.fromCharCode(65 + i)}. ${option}`),
    ].join('\n');
    if (p.actions.some((a) => a.type !== 'interaction_checkpoint' && (a.type !== 'speech' || a.text !== reading)))
      return;
    return isDiscussion(checkpoint) ? 'checkpoint' : 'quiz';
  };
  const interactionCanvas = (p: typeof page, number: number, checkpoint: any) => {
    const objects = nativeTeachingObjects(number);
    const targets = objects.filter(
      ({ nativeShape }) => readPresentationShapeText(nativeShape).trim() === String(checkpoint?.question ?? '').trim(),
    );
    const text = targets.length === 1 ? targets[0] : undefined;
    const canvas = {
      id: p.pageId,
      viewportSize: 1280,
      viewportRatio: 0.5625,
      theme: { backgroundColor: '#fff', themeColors: [], fontColor: '#222', fontName: 'Arial' },
      elements: text
        ? [
            {
              id: text.objectId,
              type: 'text',
              left: 0,
              top: 0,
              width: 1280,
              height: 200,
              content: String(checkpoint.question)
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;'),
            },
          ]
        : [],
    };
    return { canvas, questionId: text?.objectId };
  };
  const assessment = () => ({
    nativeTeaching: true,
    readOnly: getReadonlyMode(),
    sourceSha256,
    canvas: {
      width: 1000,
      height: (() => {
        const logic = (getEditorApi() as any)?.WordControl?.m_oLogicDocument;
        return ((logic?.GetHeightMM?.() ?? 562.5) / (logic?.GetWidthMM?.() ?? 1000)) * 1000;
      })(),
    },
    projectId: teaching.documentId,
    revision: teaching.revision,
    sourceRevision: teaching.revision,
    classroom: {
      id: teaching.documentId,
      stage: { id: teaching.documentId, name: String(teaching.title ?? 'Teaching'), createdAt: 0, updatedAt: 0 },
      createdAt: new Date(0).toISOString(),
      scenes: teaching.pages.map((p, index) => {
        let checkpoint = pageRole(p)
          ? p.actions.find((a) => a.type === 'interaction_checkpoint' && a.checkpointKind === 'chapter_summary')
          : undefined;
        const details = (teaching.checkpointDetails as Record<string, any> | undefined)?.[p.pageId];
        if (checkpoint && details?.actionId === checkpoint.id)
          checkpoint = {
            ...checkpoint,
            ...(details.constraintPrompt ? { constraintPrompt: details.constraintPrompt } : {}),
            referenceAnswer: {
              ...(checkpoint.referenceAnswer as object),
              ...(details.label ? { label: details.label } : {}),
            },
          };
        const response = checkpoint?.response as any;
        const reference = checkpoint?.referenceAnswer as any;
        const content =
          (teaching.assessments as Record<string, unknown> | undefined)?.[p.pageId] ??
          (Array.isArray(generation?.pages) ? generation.pages.find((g) => g?.pageId === p.pageId)?.quiz : undefined) ??
          (checkpoint
            ? {
                type: 'quiz',
                interactionMode: 'classroom_instant_feedback',
                questions: [
                  {
                    id: checkpoint.id,
                    type: response?.kind === 'discussion' || !response?.kind ? 'short_answer' : response.kind,
                    question: checkpoint.question,
                    options: (response?.options ?? []).map((label: string, i: number) => ({
                      value: String.fromCharCode(65 + i),
                      label,
                    })),
                    answer: (response?.answerIndices ?? []).map((i: number) => String.fromCharCode(65 + i)),
                    hasAnswer: true,
                    points: 1,
                    referenceAnswer: reference?.summary ?? '',
                    keyPoints: reference?.keyPoints ?? [],
                  },
                ],
              }
            : {
                type: 'slide',
                canvas: {
                  id: p.pageId,
                  viewportSize: 1280,
                  viewportRatio: 0.5625,
                  theme: { backgroundColor: '#fff', themeColors: [], fontColor: '#222', fontName: 'Arial' },
                  elements: [],
                },
              });
        const discussion = checkpoint && pageRole(p) === 'checkpoint';
        // Project the MAIC contract explicitly; native playback fields stay in the document.
        const interactionAction =
          discussion && checkpoint
            ? {
                id: checkpoint.id,
                type: checkpoint.type,
                checkpointKind: checkpoint.checkpointKind,
                checkpointId: checkpoint.checkpointId,
                chapterId: checkpoint.chapterId,
                chapterTitle: checkpoint.chapterTitle,
                question: checkpoint.question,
                referenceAnswer: checkpoint.referenceAnswer,
                context: checkpoint.context,
                ...(checkpoint.constraintPrompt ? { constraintPrompt: checkpoint.constraintPrompt } : {}),
              }
            : undefined;
        const interaction = discussion ? interactionCanvas(p, index + 1, checkpoint) : undefined;
        const actualContent = discussion ? { type: 'slide', canvas: interaction!.canvas } : content;
        return {
          id: p.pageId,
          stageId: teaching.documentId,
          title: String(checkpoint?.chapterTitle ?? `Page ${index + 1}`),
          order: index,
          type: (actualContent as any).type,
          content: actualContent,
          actions: [
            ...p.actions
              .filter((a) => a.type === 'speech')
              .map((a) => ({
                id: a.id,
                type: 'speech',
                text: a.text,
                ...(a.audioAssetId ? { audioId: a.audioAssetId } : {}),
              })),
            ...(interactionAction ? [interactionAction] : []),
          ],
          metadata: {
            ...(interaction?.questionId ? { nativeQuestionElementId: interaction.questionId } : {}),
            assessmentSpeechIds: Object.fromEntries(
              (checkpoint ? [discussion ? checkpoint.id : p.pageId] : []).map((id) => [
                id,
                p.actions.find((s) => s.type === 'speech')?.id ?? `${id}-question-tts`,
              ]),
            ),
          },
        };
      }),
    },
  });
  const view = () => {
    return {
      nativeTeaching: true,
      enrichmentCapabilities: { version: 1 },
      backgroundRegionRepair:
        !page.locked &&
        !teaching.enrichment?.originalPages.some((p) => p.pageId === page.pageId) &&
        !(teaching.backgroundRegions as Record<string, unknown> | undefined)?.[page.pageId] &&
        requirements.some(
          (p) =>
            p.pageId === page.pageId &&
            (p.teaching as any)?.pageType === 'slide' &&
            ['animation', 'spotlight'].every((effect) => (p.teaching as any).requiredEffects?.includes(effect)),
        ) &&
        canExtractBackgroundRegion(pageNumber),
      ...(teaching.enrichment ? { enrichment: teaching.enrichment } : {}),
      readOnly: getReadonlyMode(),
      teaching: requirements.find((p) => p.pageId === page.pageId)?.teaching,
      teachingRequirements: requirements,
      viewId: teaching.documentId + '.' + page.pageId,
      pageId: page.pageId,
      pageNumber,
      ...(pageRole(page) === 'checkpoint'
        ? { embeddedAssessment: assessment() }
        : pageRole(page) === 'quiz'
          ? { assessment: assessment() }
          : {}),
      project: {
        schemaVersion: 'deckwise-courseware/1',
        projectId: teaching.documentId,
        taskType: 'plain',
        title: typeof teaching.title === 'string' ? teaching.title : 'Native teaching',
        teachingRevision: teaching.revision,
        ...(generation?.educationContext ? { educationContext: generation.educationContext } : {}),
        source: { revision: teaching.revision, sha256: deckwisePageLineage.context()?.sourceSha256 ?? '' },
        pages: teaching.pages,
      },
      teachingObjects: nativeTeachingObjects(pageNumber).map(({ nativeShape: _nativeShape, ...object }) => {
        const semantics = (teaching.objectSemantics as Record<string, any[]> | undefined)?.[page.pageId]?.find(
          (o) => o.objectId === object.objectId,
        );
        return {
          ...object,
          ...(semantics
            ? { role: semantics.role, aliases: semantics.aliases, effectTargetEligible: semantics.teachingTarget }
            : { aliases: [] }),
        };
      }),
    };
  };
  if (operation === 'open') return { value: view() };
  if (operation === 'assessment-context') return { value: assessment() };
  if (operation === 'assessment-write') {
    if (
      args.projectId !== teaching.documentId ||
      args.expectedRevision !== teaching.revision ||
      args.expectedSourceRevision !== teaching.revision
    )
      throw new Error('Native assessment changed');
    if (getReadonlyMode()) throw new Error('Teaching is read-only');
    const before = assessment().classroom.scenes;
    const changes: Array<{ apply(): void; restore(): void; assertCurrent?(): void }> = [];
    for (const projection of (args.projections ?? []) as any[]) {
      const p = teaching.pages[projection.page - 1];
      if (
        !p ||
        p.locked ||
        !(args.changedPageIds as string[]).includes(p.pageId) ||
        projection.next?.slideId !== p.pageId
      )
        throw new Error('Invalid native Quiz projection');
      changes.push(await prepareQuizProjection(projection));
    }
    for (const scene of (args.classroom as any).scenes) {
      const index = teaching.pages.findIndex((p) => p.pageId === scene.id);
      if (!(args.changedPageIds as string[]).includes(scene.id)) continue;
      const old = before[index];
      if (scene.content.type === 'quiz') {
        if (!online) continue;
        if (old?.content.type !== 'quiz') throw new Error('Online Quiz page changed');
        const [previous] = old.content.questions;
        const [next] = scene.content.questions;
        if (
          !previous ||
          scene.content.questions.length !== 1 ||
          next.type !== previous.type ||
          JSON.stringify(next.image) !== JSON.stringify(previous.image) ||
          JSON.stringify(next.options?.map(({ value, image }: any) => [value, image])) !==
            JSON.stringify(previous.options?.map(({ value, image }: any) => [value, image]))
        )
          throw new Error('题型、选项数量和配图需要在桌面本地编辑后再同步。');
        const text = [
          { suffix: '_title', before: `Quiz · ${old.title}`, after: `Quiz · ${scene.title}` },
          { suffix: '_q0_question', before: previous.question, after: next.question },
          ...(previous.options ?? []).map((option: any, position: number) => ({
            suffix: `_q0_option${position}_label`,
            before: `${option.value}  ${option.label}`,
            after: `${option.value}  ${next.options[position].label}`,
          })),
        ].filter((entry) => entry.before !== entry.after);
        if (!text.length) continue;
        const objects = nativeTeachingObjects(index + 1);
        const anchors = objects.filter(
          ({ name, nativeShape }) =>
            name.endsWith('_q0_question') && readPresentationShapeText(nativeShape).trim() === previous.question.trim(),
        );
        if (anchors.length !== 1) throw new Error('题目文本框已修改，请核对 PPT 后重试。');
        const prefix = anchors[0].name.slice(0, -'_q0_question'.length);
        for (const { suffix, before, after } of text) {
          const matches = objects.filter(({ name }) => name === prefix + suffix);
          if (matches.length !== 1) throw new Error('题目文本框已修改，请核对 PPT 后重试。');
          changes.push(prepareAssessmentText(index + 1, matches[0].nativeShape, before, after));
        }
        continue;
      }
      for (const action of scene.actions ?? [])
        if (action.type === 'interaction_checkpoint') {
          const previous = old?.actions.find((a) => a?.id === action.id) as Record<string, any> | undefined;
          if (!previous || previous.question === action.question) continue;
          const id = (old.metadata as any).nativeQuestionElementId;
          if (!id || scene.metadata?.nativeQuestionElementId !== id) throw new Error('Question text binding changed');
          const object = nativeTeachingObjects(index + 1).find((o) => o.objectId === id);
          if (!object) throw new Error('Question text binding changed');
          changes.push(
            prepareAssessmentText(index + 1, object.nativeShape, String(previous.question), action.question),
          );
        }
    }
    if (readNativeTeaching()?.revision !== teaching.revision || getReadonlyMode())
      throw new Error('Native assessment changed');
    changes.forEach((change) => change.assertCurrent?.());
    replaceNativeAssessments({
      documentId: teaching.documentId,
      expectedRevision: teaching.revision,
      scenes: (args.classroom as any).scenes,
      changedPageIds: args.changedPageIds as string[],
      ...(changes.length
        ? {
            change: {
              apply() {
                changes.forEach((c) => c.apply());
              },
              restore() {
                [...changes].reverse().forEach((c) => c.restore());
              },
            },
          }
        : {}),
    });
    return routeNativeTeaching('assessment-context', {}, pageNumber);
  }
  if (operation.startsWith('classroom:')) return { value: { nativeService: true } };
  if (operation === 'notes-context') {
    const api = getEditorApi() as any;
    const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
    return {
      value: {
        view: view(),
        text: (slide?.notes?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n').trim(),
      },
    };
  }
  if (operation === 'narration-visual') {
    const logic = (getEditorApi() as any)?.WordControl?.m_oLogicDocument;
    const slide = logic?.Slides?.[pageNumber - 1];
    if (!slide) throw new Error('Native page unavailable');
    const layers = (
      await Promise.all(
        nativeTeachingObjects(pageNumber).map(async ({ objectId, nativeShape }) => {
          const evidence = await nativeLayerEvidence(nativeShape);
          return evidence ? { objectId, dataUrl: evidence.dataUrl } : undefined;
        }),
      )
    ).filter(Boolean);
    return { value: { dataUrl: renderNativeSlide(slide, logic, requireEditorContext().AscCommon as any), layers } };
  }
  if (operation === 'narration-context') {
    const api = getEditorApi() as any;
    const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
    const text = (slide?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n');
    const plannedPage = Array.isArray(generation?.pages)
      ? generation.pages.find((p) => p?.pageId === page.pageId)
      : undefined;
    const notes = (slide?.notes?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n').trim();
    return {
      value: {
        view: view(),
        text,
        evidence: {
          width: api?.WordControl?.m_oLogicDocument?.GetWidthMM?.() ?? 0,
          height: api?.WordControl?.m_oLogicDocument?.GetHeightMM?.() ?? 0,
          notes: notes || (typeof plannedPage?.narration === 'string' ? plannedPage.narration : ''),
          visualSummary: text,
          objects: await Promise.all(
            nativeTeachingObjects(pageNumber).map(async ({ nativeShape, objectId, name }) => {
              const shape = nativeShape as unknown as {
                x?: number;
                y?: number;
                extX?: number;
                extY?: number;
                spTree?: unknown[];
                isImage?(): boolean;
              };
              const text = readPresentationShapeText(nativeShape);
              const layer = await nativeLayerEvidence(nativeShape);
              return {
                objectId,
                name,
                aliases: [name],
                ...(layer ? { layerMetadata: layer.layerMetadata } : {}),
                text,
                kind: shape.spTree?.length ? 'group' : text.trim() ? 'text' : shape.isImage?.() ? 'image' : 'shape',
                bbox: [shape.x ?? 0, shape.y ?? 0, shape.extX ?? 0, shape.extY ?? 0],
              };
            }),
          ),
        },
      },
    };
  }
  if (
    ['history', 'draft:get', 'draft:set'].includes(operation) &&
    args.viewId !== teaching.documentId + '.' + page.pageId
  )
    throw new Error('Native teaching view changed');
  if (operation === 'history') return { value: { entries: [], skippedEntryCount: 0 } };
  if (
    operation !== 'open' &&
    (args.viewId !== teaching.documentId + '.' + page.pageId || args.expectedTeachingRevision !== teaching.revision) &&
    !['draft:get', 'draft:set', 'undo', 'redo'].includes(operation)
  )
    throw new Error('Native teaching view changed');
  if (operation === 'validate-animations') {
    assertNativeTeachingAnimations(pageNumber);
    return { value: { valid: true } };
  }
  if (operation === 'adopt-audio') {
    await adoptNativeTeachingAudio({
      ...(args.command as Parameters<typeof adoptNativeTeachingAudio>[0]),
      documentId: teaching.documentId,
      expectedRevision: teaching.revision,
      pageId: page.pageId,
    });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'confirm') {
    confirmNativeTeachingPage({
      documentId: teaching.documentId,
      expectedRevision: teaching.revision,
      pageId: page.pageId,
    });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'draft:set' && args.draft != null)
    throw new Error('Native teaching edits require write with the current teaching revision');
  if (operation === 'draft:get' || operation === 'draft:set') return { value: { draft: null } };
  if (operation === 'lock') {
    setNativeTeachingLock({
      documentId: teaching.documentId,
      expectedRevision: String(args.expectedTeachingRevision),
      pageId: page.pageId,
      locked: args.locked as boolean,
    });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'undo' || operation === 'redo') {
    nativeTeachingHistory(operation === 'undo');
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'assessment-insert') {
    if (args.viewId !== teaching.documentId + '.' + page.pageId) throw new Error('Native teaching view changed');
    const projection = args.projection as AssessmentProjection;
    if (projection?.slideId !== args.pageId) throw new Error('Assessment projection identity changed');
    const change = prepareAssessmentInsertion(pageNumber - 1, projection);
    insertNativeTeachingAssessment({
      documentId: teaching.documentId,
      expectedRevision: String(args.expectedTeachingRevision),
      afterPageId: page.pageId,
      pageId: String(args.pageId),
      actions: args.actions as Parameters<typeof insertNativeTeachingAssessment>[0]['actions'],
      enrichment: args.enrichment as Parameters<typeof insertNativeTeachingAssessment>[0]['enrichment'],
      quiz: args.quiz,
      change,
    });
    return routeNativeTeaching('open', {}, pageNumber + 1);
  }
  if (operation === 'repair-background-region') {
    if (
      !view().backgroundRegionRepair ||
      typeof args.actionId !== 'string' ||
      typeof args.description !== 'string' ||
      !Array.isArray(args.bounds) ||
      args.bounds.length !== 4 ||
      !args.bounds.every((n) => typeof n === 'number' && Number.isFinite(n))
    )
      throw new Error('Background region repair is unavailable or invalid');
    if (typeof args.expectedVisualSha256 !== 'string' || !/^[a-f0-9]{64}$/.test(args.expectedVisualSha256))
      throw new Error('Region repair requires reviewed page pixels');
    const logic = (getEditorApi() as any)?.WordControl?.m_oLogicDocument;
    const png = renderNativeSlide(logic.Slides[pageNumber - 1], logic, requireEditorContext().AscCommon as any);
    const bytes = Uint8Array.from(atob(png.split(',')[1]), (c) => c.charCodeAt(0));
    const digest = [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))]
      .map((v) => v.toString(16).padStart(2, '0'))
      .join('');
    if (
      digest !== args.expectedVisualSha256 ||
      readNativeTeaching()?.revision !== teaching.revision ||
      renderNativeSlide(logic.Slides[pageNumber - 1], logic, requireEditorContext().AscCommon as any) !== png
    )
      throw new Error('Reviewed region pixels changed');
    const change = await prepareBackgroundRegion(pageNumber, args.bounds as number[], `region-${crypto.randomUUID()}`);
    if (readNativeTeaching()?.revision !== teaching.revision || getReadonlyMode())
      throw new Error('Teaching changed during region preparation');
    replaceNativeTeachingRegion({
      documentId: teaching.documentId,
      expectedRevision: teaching.revision,
      pageId: page.pageId,
      actionId: args.actionId,
      description: args.description,
      change,
    });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'write') {
    if (args.viewId !== teaching.documentId + '.' + page.pageId) throw new Error('Native teaching view changed');
    const actions = structuredClone(args.actions) as Parameters<typeof replaceNativeTeachingPage>[0]['actions'];
    let verifiedNativeAnimationBindings: Record<string, string> | undefined;
    if (online && Array.isArray(actions)) {
      const steps = actions.filter((action) => action?.type === 'native_animation');
      const saved = page.actions.filter((action) => action.type === 'native_animation');
      if (
        steps.length &&
        !steps.every(
          (step, index) => step.id === saved[index]?.id && step.stepIndex === index && steps.length === saved.length,
        )
      ) {
        const expected = steps.map((step) => step.expectedAnimation);
        if (steps.length > 200 || expected.some((signature) => typeof signature !== 'string'))
          throw new Error('请重新选择本页 PPT 动画后再保存');
        // The carrier uses a common anchor. Fixed exports re-pin it to the saved package hash.
        const hash = await crypto.subtle.digest(
          'SHA-256',
          new TextEncoder().encode(JSON.stringify([teaching.documentId, page.pageId, expected])),
        );
        const anchor = [...new Uint8Array(hash)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
        const current = nativeAnimationSignatures(pageNumber, true);
        if (
          current.length !== steps.length ||
          steps.some(
            (step, index) => step.version !== 1 || step.stepIndex !== index || expected[index] !== current[index],
          )
        )
          throw new Error('本页 PPT 动画在绑定后发生变化，请移除草稿中的动画绑定并重新关联');
        for (const step of steps) {
          step.sourceSha256 = anchor;
          delete step.expectedAnimation;
        }
        verifiedNativeAnimationBindings = Object.fromEntries(steps.map((step, index) => [step.id, current[index]]));
      } else {
        for (const [index, step] of steps.entries()) {
          step.sourceSha256 = saved[index].sourceSha256;
          delete step.expectedAnimation;
        }
      }
    }
    const next = replaceNativeTeachingPage(
      {
        documentId: teaching.documentId,
        expectedRevision: String(args.expectedTeachingRevision),
        pageId: page.pageId,
        enrichment: args.enrichment as Parameters<typeof replaceNativeTeachingPage>[0]['enrichment'],
        semantics: args.semantics as Parameters<typeof replaceNativeTeachingPage>[0]['semantics'],
        actions,
        verifiedNativeAnimationBindings,
      },
      true,
    );
    return routeNativeTeaching('open', {}, next.pages.findIndex((p) => p.pageId === page.pageId) + 1);
  }
  if (operation === 'tts:preview' || operation === 'media:preview') {
    const action = page.actions.find((a) => a.id === args.actionId);
    const id = action?.audioAssetId ?? action?.videoAssetId;
    if (typeof id !== 'string') throw new Error('Native teaching action has no adopted media');
    const blob = await nativeTeachingAudio(id);
    return { value: { bytes: await blob.arrayBuffer(), mimeType: blob.type } };
  }
  if (operation === 'media:adopt') {
    const asset = teaching.media.find((m) => m.id === args.intentId);
    if (!asset) return { value: { nativeService: true } };
    if (
      asset.mimeType.startsWith('video/') !== (args.kind === 'video') ||
      typeof args.actionId !== 'string' ||
      !['music', 'video'].includes(String(args.kind))
    )
      throw new Error('Embedded media unavailable');
    replaceNativeTeachingPage({
      documentId: teaching.documentId,
      expectedRevision: teaching.revision,
      pageId: page.pageId,
      actions: [
        ...page.actions,
        {
          id: args.actionId,
          type: args.kind === 'video' ? 'video_clip' : 'audio_clip',
          [args.kind === 'video' ? 'videoAssetId' : 'audioAssetId']: asset.id,
          durationMs: asset.durationMs,
        },
      ],
    });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'media:list')
    return {
      value: {
        nativeService: true,
        embedded: teaching.media
          .filter((m) => !m.id.startsWith('tts-'))
          .map((m) => ({
            kind: m.mimeType.startsWith('video/') ? 'video' : 'music',
            intentId: m.id,
            title: m.id,
            model: 'embedded',
            inPage: page.actions.some((a) => a.audioAssetId === m.id || a.videoAssetId === m.id),
          })),
      },
    };
  if (operation.startsWith('tts:')) return { value: { nativeService: true } };
  throw new Error('Native teaching operation requires the native media service: ' + operation);
}

export function nativeTeachingHistory(undo: boolean): void {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const api = getEditorApi() as unknown as { Undo(): void; Redo(): void } | null;
  if (!api) throw new Error('Native teaching history unavailable');
  if (undo) api.Undo();
  else api.Redo();
}
