import { installCoursewareOnlineEntry } from './courseware-online-entry';

const params = new URLSearchParams(location.search);
const root = document.getElementById('courseware-online');
try {
  if (!root || window.parent === window) throw new Error('Missing owned online editor parent');
  const dispose = installCoursewareOnlineEntry(root, {
    parentWindow: window.parent,
    parentOrigin: params.get('parentOrigin') ?? '',
    channelId: params.get('channelId') ?? '',
    apiUrl: params.get('apiUrl') ?? '',
    ...(params.has('communityPluginConfigUrl')
      ? { communityPluginConfigUrl: params.get('communityPluginConfigUrl')! }
      : {}),
  });
  window.addEventListener('pagehide', dispose, { once: true });
} catch {
  if (root) root.textContent = 'ONLINE_ENTRY_CONFIGURATION_INVALID';
}
