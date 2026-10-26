import { afterEach, expect, it, vi } from 'vitest';

const { waitForControl } = vi.hoisted(() => ({ waitForControl: vi.fn() }));
vi.mock('ranui/message', () => ({}));
vi.mock('../../lib/offline-worker-control', () => ({ waitForAppWorkerControl: waitForControl }));
import { loadEditorApi } from '../../lib/onlyoffice-editor';

afterEach(() => {
  document.querySelectorAll('script[src*="documents/api.js"]').forEach((script) => script.remove());
  Reflect.deleteProperty(window, 'DocsAPI');
  history.replaceState({}, '', '/');
  waitForControl.mockReset();
});

it('loads an embedded editor without waiting for optional offline cache control', async () => {
  history.replaceState({}, '', '/editor?embed=1');
  waitForControl.mockImplementation(() => new Promise(() => {}));
  const loading = loadEditorApi();
  await Promise.resolve();
  expect(waitForControl).not.toHaveBeenCalled();
  const script = document.querySelector<HTMLScriptElement>('script[src*="documents/api.js"]');
  expect(script).not.toBeNull();
  script!.dispatchEvent(new Event('load'));
  await loading;
});

it('preserves the standalone offline cache handoff before loading the editor', async () => {
  history.replaceState({}, '', '/editor');
  let controlled!: () => void;
  waitForControl.mockImplementation(
    () =>
      new Promise<void>((resolve) => {
        controlled = resolve;
      }),
  );
  const loading = loadEditorApi();
  expect(waitForControl).toHaveBeenCalledOnce();
  expect(document.querySelector('script[src*="documents/api.js"]')).toBeNull();
  controlled();
  await Promise.resolve();
  document.querySelector('script[src*="documents/api.js"]')!.dispatchEvent(new Event('load'));
  await loading;
});
