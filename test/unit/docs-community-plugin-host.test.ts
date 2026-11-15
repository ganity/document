import { afterEach, describe, expect, it, vi } from 'vitest';
import { connectDocsCommunityReads } from '../../lib/docs-community-plugin-host';
import { installDocsCommunityPluginEntry } from '../../lib/docs-community-plugin-entry';
import { DOCS_COMMUNITY_PROTOCOL, type CommunityPluginApi } from '../../lib/docs-community-plugin';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
  document.body.replaceChildren();
});
const ready = { protocol: DOCS_COMMUNITY_PROTOCOL, type: 'ready' };
function frame() {
  const element = document.createElement('iframe');
  document.body.append(element);
  return element.contentWindow!;
}
function message(source: Window, origin = location.origin, data = ready, ports: MessagePort[] = []) {
  window.dispatchEvent(new MessageEvent('message', { source, origin, data, ports }));
}
class FakePort {
  onmessage = null;
  close = vi.fn();
  start = vi.fn();
  postMessage = vi.fn();
}
class FakeChannel {
  port1 = new FakePort();
  port2 = new FakePort();
}

describe('owned plugin discovery', () => {
  it('ignores wrong origin, protocol and sibling frame, then connects only its own descendant', async () => {
    vi.stubGlobal('MessageChannel', FakeChannel);
    const editor = frame();
    const sibling = frame();
    const post = vi.spyOn(editor, 'postMessage').mockImplementation(() => {});
    const controller = new AbortController();
    const connected = connectDocsCommunityReads(editor, 's', controller.signal);
    message(editor, 'https://other.example');
    message(sibling);
    message(editor, location.origin, { ...ready, protocol: 'wrong' });
    expect(post).not.toHaveBeenCalled();
    message(editor);
    const client = await connected;
    expect(post).toHaveBeenCalledWith({ ...ready, type: 'connect', sessionId: 's' }, location.origin, [
      expect.any(FakePort),
    ]);
    client.close();
    controller.abort();
  });

  it('accepts a nested plugin under the owned editor only', async () => {
    vi.stubGlobal('MessageChannel', FakeChannel);
    const editor = frame();
    const child = editor.document.createElement('iframe');
    editor.document.body.append(child);
    const post = vi.spyOn(child.contentWindow!, 'postMessage').mockImplementation(() => {});
    const connected = connectDocsCommunityReads(editor, 's', new AbortController().signal);
    message(child.contentWindow!);
    (await connected).close();
    expect(post).toHaveBeenCalledOnce();
  });

  it.each(['abort', 'timeout'])('cleans discovery after %s', async (why) => {
    vi.useFakeTimers();
    const editor = frame();
    const controller = new AbortController();
    const rejected = expect(connectDocsCommunityReads(editor, 's', controller.signal)).rejects.toThrow();
    if (why === 'abort') controller.abort();
    else await vi.advanceTimersByTimeAsync(30000);
    await rejected;
    expect(vi.getTimerCount()).toBe(0);
  });

  it('rejects an invalid owned window or session before listening', () => {
    expect(() => connectDocsCommunityReads(window, 's', new AbortController().signal)).toThrow();
    expect(() => connectDocsCommunityReads(frame(), '../s', new AbortController().signal)).toThrow();
    const aborted = new AbortController();
    aborted.abort();
    expect(() => connectDocsCommunityReads(frame(), 's', aborted.signal)).toThrow();
  });

  it('closes both ports and the client when transfer fails', async () => {
    const channel = new FakeChannel();
    vi.stubGlobal(
      'MessageChannel',
      vi.fn(function () {
        return channel;
      }),
    );
    const editor = frame();
    const post = vi.spyOn(editor, 'postMessage').mockImplementation(() => {
      throw new Error('Transfer failed');
    });
    const rejected = expect(connectDocsCommunityReads(editor, 's', new AbortController().signal)).rejects.toThrow(
      'Transfer failed',
    );
    message(editor);
    await rejected;
    expect(channel!.port1.close).toHaveBeenCalled();
    expect(channel!.port2.close).toHaveBeenCalled();
    post.mockRestore();
  });
});

describe('plugin entry connection boundary', () => {
  it('connects to a same-origin launcher below the outer desktop window', () => {
    const launcher = frame();
    const host = launcher.document.createElement('iframe');
    launcher.document.body.append(host);
    const child = host.contentDocument!.createElement('iframe');
    host.contentDocument!.body.append(child);
    const pluginWindow = child.contentWindow!;
    const sibling = frame();
    const post = vi.spyOn(launcher, 'postMessage').mockImplementation(() => {});
    vi.stubGlobal('window', pluginWindow);
    const plugin: CommunityPluginApi & { init?: () => void } = {
      info: { editorType: 'slide' },
      executeMethod: vi.fn(),
      callCommand: vi.fn(),
    };
    const controller = new AbortController();
    installDocsCommunityPluginEntry(plugin, controller.signal);
    plugin.init!();
    expect(post).toHaveBeenCalledWith(ready, location.origin);
    const port = new FakePort();
    const connect = { ...ready, type: 'connect', sessionId: 'nested' };
    message(sibling, location.origin, connect, [port as unknown as MessagePort]);
    message(launcher, 'https://other.example', connect, [port as unknown as MessagePort]);
    expect(port.start).not.toHaveBeenCalled();
    message(launcher, location.origin, connect, [port as unknown as MessagePort]);
    expect(port.start).toHaveBeenCalledOnce();
    controller.abort();
    expect(port.close).toHaveBeenCalledOnce();
    post.mockRestore();
  });

  it('announces readiness after native init and refuses unowned ports', () => {
    const plugin: CommunityPluginApi & { init?: () => void } = {
      info: { editorType: 'slide' },
      executeMethod: vi.fn(),
      callCommand: vi.fn(),
      init: vi.fn(),
    };
    const previous = plugin.init;
    const post = vi.spyOn(window, 'postMessage').mockImplementation(() => {});
    const controller = new AbortController();
    installDocsCommunityPluginEntry(plugin, controller.signal);
    expect(post).not.toHaveBeenCalled();
    plugin.init!();
    expect(previous).toHaveBeenCalledOnce();
    expect(post).toHaveBeenCalledWith(ready, location.origin);
    const port = new FakePort();
    const connect = { ...ready, type: 'connect', sessionId: 's' };
    message(frame(), location.origin, connect, [port as unknown as MessagePort]);
    message(window, 'https://other.example', connect, [port as unknown as MessagePort]);
    message(window, location.origin, connect, []);
    expect(port.start).not.toHaveBeenCalled();
    message(window, location.origin, connect, [port as unknown as MessagePort]);
    expect(port.start).toHaveBeenCalledOnce();
    controller.abort();
    expect(port.close).toHaveBeenCalledOnce();
    expect(plugin.init).toBe(previous);
    post.mockRestore();
  });
});
