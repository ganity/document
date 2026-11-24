# iframe 嵌入 API

本项目支持通过 iframe 嵌入到任何 Web 应用中。推荐架构是：**父系统负责鉴权、下载文件和上传保存结果；iframe 只负责文档编辑**。Token、Cookie、业务接口都留在父系统内，编辑器不需要知道授权细节。

项目内置了完整示例页面 `/embed-demo.html`，包含保存时的 sha256 哈希日志，方便调试。

---

## 嵌入编辑器

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1"
  style="width: 100%; height: 720px; border: 0"
></iframe>
```

### DeckWise 受限模式

增加 `deckwise=1` 后，编辑器保留内部保存按钮和 Ctrl+S，同时隐藏并阻止下载、另存为、
其他格式导出和打印：

```html
<iframe src="https://your-deployment/editor?embed=1&deckwise=1&embedOrigin=https://your-system.example.com"></iframe>
```

该模式接受所有已支持格式的 `document:open-buffer`：Word（`docx`、`doc`、`odt`、`rtf`、`txt`）、
Excel（`xlsx`、`xls`、`ods`、`csv`）、PowerPoint（`pptx`、`ppt`、`odp`）及 PDF（`pdf`）。
设置 `readonly: true` 即可预览。不支持的格式返回 `document:error`，由宿主提供系统默认应用打开入口。
格式识别不保证所有文件都能转换；部分较大的 PPT、ODP 文件仍可能被底层转换引擎拒绝。
URL 输入仍被禁止。保存使用文档自身格式，旧版 `doc`、`xls`、`ppt` 分别保存为 `docx`、`xlsx`、`pptx`，
禁止导出成其他格式。用户点击
内部保存按钮时，编辑器会向宿主发送 `document:save-requested`；宿主随后发送正常的
`document:save` 命令，编辑后的文件最终通过 `document:saved` 返回。

> 编辑器位于 `/editor`，首页 `/` 是纯静态落地页。旧链接 `/?embed=1`、`/?src=`、`/?file=`、`/?new=` 仍然有效——`/` 会带同样的参数跳转到 `/editor`。

控制命令只接受直接父窗口发送的消息。未设置 `embedOrigin` 时，首次接受的父窗口命令会
固定本次编辑器页面的来源，之后来自其他来源的命令会被忽略。
如需从启动时就限制只接受指定来源的消息，增加 `embedOrigin`：

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1&embedOrigin=https://your-system.example.com"
></iframe>
```

---

## 发送命令

建议每条命令带上 `id`，便于匹配响应：

```js
const iframe = document.getElementById('documentEditor');
const editorOrigin = 'https://your-deployment';

function sendEditorCommand(type, payload = {}) {
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  iframe.contentWindow.postMessage({ id, type, payload }, editorOrigin);
  return id;
}

window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { id, type, payload } = event.data || {};
  if (!type?.startsWith('document:')) return;

  switch (type) {
    case 'document:ready':
      console.log('编辑器已就绪');
      break;
    case 'document:opened':
      console.log('文档已打开', id, payload);
      break;
    case 'document:saved':
      console.log('保存完成', payload.fileName, payload.file);
      break;
    case 'document:error':
      console.error('操作失败', payload.message);
      break;
  }
});
```

---

## 打开文档

### 通过 URL

> **这个 URL 必须允许 CORS。** 取文件的是浏览器而不是服务器，所以响应里要有覆盖编辑器所在
> 页面的 `Access-Control-Allow-Origin` 头。没有它，请求在编辑器拿到第一个字节之前就被拦掉了；
> URL 上的 `?src=`、`?file=` 同理。这是接入时最容易踩的一脚：如果一个文件手动下载能打开、
> 走编辑器打不开，先看响应头。跨域跳转也要一路带着这个头。
> 加不了的场景（第三方主机、签名 URL、需要鉴权的接口），改由父页面自己 fetch，再用
> `document:open-buffer` 把字节传进来。

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/files/demo.xlsx',
  fileName: 'demo.xlsx',
  readonly: false,
});
```

如果 URL 需要授权头，可以传 `fetchOptions`。但对于需要鉴权的文件，更推荐由父系统自己 `fetch` 后传入二进制数据：

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/api/files/1',
  fileName: 'demo.xlsx',
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

### 通过本地文件选择

```js
const input = document.createElement('input');
input.type = 'file';
input.accept = '.xlsx,.xls,.csv,.docx,.doc,.pptx,.ppt';
input.onchange = () => {
  sendEditorCommand('document:open-file', { file: input.files[0], readonly: false });
};
input.click();
```

### 通过父系统授权请求（推荐用于受保护文件）

```js
const response = await fetch('/api/files/1', {
  headers: { Authorization: `Bearer ${token}` },
});
const buffer = await response.arrayBuffer();
sendEditorCommand('document:open-buffer', { fileName: 'demo.xlsx', buffer, readonly: false });
```

---

## 只读模式

在打开文档时通过 `readonly` 字段设置，或随时切换：

```js
sendEditorCommand('document:set-readonly', { readonly: true });
```

只读模式下编辑权限关闭，但导出仍然可用：`document:save` 会通过 `document:saved`
返回当前字节（只读禁止的是编辑，不是下载）。embed iframe 内 File 页签的下载
控件在父页没有发起 `document:save` 时，也会触发本地下载。

---

## 保存并上传

保存命令触发编辑器导出当前文档，通过 `document:saved` 返回 `File` 对象。默认保存为当前文档自身的格式（`.docx` 存为 DOCX、`.csv` 存为 CSV……），通过 `targetExt` 指定其他格式。

```js
sendEditorCommand('document:save', { targetExt: 'XLSX' }); // XLSX、DOCX、PPTX、CSV
```

默认情况下，命令会等待编辑器返回**编辑后**的文件。超时则返回 `document:error`，避免误上传原始文件。如需超时时回传原文件，显式开启：

```js
sendEditorCommand('document:save', { targetExt: 'XLSX', returnOriginalOnTimeout: true });
```

DeckWise 模式下，SDK 正在导出时会拒绝编辑工具的上下文查询/命令、打开另一文档和第二次保存，即使未设置 `requireHostCommit`。普通导出成功后解除限制；设置 `requireHostCommit: true` 的保存仍需同 ID 的 `document:commit-save` 确认持久化。导出期间的手工修改仍保留为未保存。

父页面拿到文件后上传：

```js
window.addEventListener('message', async (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type !== 'document:saved') return;

  await fetch('/api/files/1', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}` },
    body: payload.file,
  });
});
```

> **注意：** 不要只用 `file.size` 判断文件是否变化。`.xlsx` 是 zip 压缩包格式，轻微编辑后文件大小可能完全相同。内置的 `/embed-demo.html` 每次保存都会在日志中打印 `sha256` 哈希值，方便调试。

---

## 查询当前状态

```js
sendEditorCommand('document:get-state');
// 响应：{ type: 'document:state', payload: { readonly: false, hasDocument: true, dirty: false } }
```

---

## 跟踪未保存的修改

编辑器会告诉父页面：当前文档是否有父页面尚未拿到的修改。父页面据此决定何时保存（切换文档、空闲定时器、关闭自己的页面），而不必导出没人动过的文档。

```js
window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type === 'document:dirty-changed') {
    console.log(payload.dirty ? '有未保存的修改' : '已全部保存');
  }
});
```

- `document:dirty-changed` 只在状态翻转时推送，而不是每次按键：第一次编辑后是 `{ dirty: true }`，一次 `document:save` 覆盖了全部修改（或撤销回到保存时的状态）后是 `{ dirty: false }`。
- `document:save` 成功即清除标志：父页面已经拿到字节。导出进行中做的修改不在这份字节里，文档仍保持 dirty——请检查 `document:saved` 的 `payload.dirty`，为 `true` 时再保存一次。
- 父页面一侧上传失败时编辑器无从得知：在上传成功之前请自己保留"需要保存"的状态。
- 带 `returnOriginalOnTimeout: true` 的 `document:save` 永远不会清除标志，因为超时返回的可能是原始文件。
- 打开另一篇文档会把标志重置为 `false`。

---

## 消息类型参考

| 方向            | 类型                        | 说明                                       |
| --------------- | --------------------------- | ------------------------------------------ |
| 父页面 → iframe | `document:open-url`         | 通过 URL 打开文档                          |
| 父页面 → iframe | `document:open-file`        | 通过 `File` / `Blob` 打开文档              |
| 父页面 → iframe | `document:open-buffer`      | 通过 `ArrayBuffer` / `Uint8Array` 打开文档 |
| 父页面 → iframe | `document:set-readonly`     | 设置只读或可编辑                           |
| 父页面 → iframe | `document:save`             | 保存并返回 `File`                          |
| 父页面 → iframe | `document:get-state`        | 查询当前状态                               |
| iframe → 父页面 | `document:ready`            | 编辑器初始化完成                           |
| iframe → 父页面 | `document:opened`           | 文档打开完成                               |
| iframe → 父页面 | `document:readonly-changed` | 只读状态已切换                             |
| iframe → 父页面 | `document:saved`            | 保存完成，返回文件（含 `dirty`）           |
| iframe → 父页面 | `document:dirty-changed`    | 未保存标志翻转（`{ dirty }`）              |
| iframe → 父页面 | `document:state`            | 当前状态响应                               |
| iframe → 父页面 | `document:error`            | 操作失败                                   |

## DeckWise 页面保存凭据（可选）

绑定 PPTX 的 `document:open-buffer` 可传 `trackPageLineage: true`。
本次打开的内容就绪回调为真实页面对象建立会话 token，计算打开文件 hash。
`document:context.pageLineage` 返回 `sourceSha256`、打开时 `openedPageTokens`
和当前顺序 `pageTokens`；重排保留 token，复制/新增得到新 token，重开全部失效。
它们不是逻辑 pageId；源 hash 始终指会话打开字节，不随后续保存前移。

Host 确认式保存的 `document:saved.pageLineage` 另含协议
`deckwise-office-page-lineage/1`、绑定、`saveRequestId` 和实际导出文件
`outputSha256`。导出或 hash 期间历史、页面或会话变化时不发凭据。
无凭据仍保护保存文件，禁止按页码迁移教学。

仅可信 Host 可从绑定 Office view 直接消费凭据；必须独立核对会话、保存请求、
源/输出 hash 和绑定打开源的完整 mapping，并与保存 revision 一起持久化。
renderer/模型 JSON 不是可信凭据。新增/删除/复制页需显式身份审核，当前 helper
拒绝自动关联；页级凭据不能继承对象 ID。改稿后的教学/音频仍走 review/CAS。
页级凭据本身不扩展工具能力，也不代表 A 的持久化已接入。

另增有限原生能力 `set_slide_font_size`：仅在当前页一个普通文本对象有非空
文字选区、可写且 SDK 设置/回读接口可用时通告。输入只接受
`{ fontSizePt }`，范围 1–96 pt、步长 0.5；仍须携带最新绑定和编辑代次。
返回 `{ fontSizePt, page, verified: true }`，并核验选区/文字未变和准确字号。
母版、放映、忙碌、只读、缺接口或多对象选区不开放。结果不确定时检查原稿并
使用原生撤销，不自动重试。该命令不保存；只有原三工具的旧 Office 仍可使用。

`set_slide_text_color` is an optional native selection capability with input
`{ "colorHex": "336699" }` (six uppercase RGB digits, no prefix). It uses the
same current-slide plain-text and writable-state guards as font size. Native
solid sRGB fill, exact text/selection and page are verified before returning
`{ "colorHex": "336699", "page": 2, "verified": true }`. Legacy `get_Color()`
does not reflect PPT Unifill; verification reads the calculated native text fill.
No automatic save/retry, alpha, theme color, gradient or shape fill is supported.

`set_slide_font_family` 沿用单个普通文本对象的非空文字选区限制。输入
`{ fontFamily }`，只能选择最新 capability 的
`inputSchema.properties.fontFamily.enum` 中准确的字体名（最多 100 个，每个
1–128 字符）。仅通告常规、斜体、粗体、粗斜体四种字形均已加载、且原生字体
加载器空闲时的字体，防止 SDK 异步下载结束后改到另一个选区。执行后核对准确
字体名、文字、选区和页码，返回 `{ fontFamily, page, verified: true }`。
字体别名可能使用已安装替代字形；此工具不下载字体、不自动保存或重试。

## 原生 PNG 换图完整性

现有已审核素材换图路径在原生注册前核对每个 PNG chunk 的 CRC，并要求完整且有界的 chunk、唯一 IHDR、图像数据及末尾 IEND。截断、尾随数据或校验错误会直接拒绝，不改动选中图片；不能只凭浏览器解码成功认定有效。assetId、hash 与明确保存边界保持不变。

## 已审核图片传输

绑定的 DeckWise 父宿主可发 `document:review-image`，载荷为 `{binding, asset}`。
`asset` 仅含已审核 PNG 的 `assetId/sha256/mime/byteLength/width/height`；传 `null`
撤销。响应为当前 `document:context`。这只是本会话图片选择，不授予服务端权限；
父宿主 Main 仍须在每次执行前核对当前用户、项目成员权限、源文件和会话。

`document:execute-image` 接受 `{command, bytes}`；消息 ID 必须等于既有有限
`replace_slide_image_from_asset` 命令的 requestId，`bytes` 为当次授权取得的
Uint8Array。普通 `document:execute` 不能复用这些字节；成功、失败和 reset 都会清理。
最新编辑代次、支持的单图选区、只读/保存/会话、PNG 完整性和原生回读守卫仍然生效。
响应 `document:command-result` 不代表已保存；载荷不接受凭据、URL 或本地路径。

## 当前页预览

当前页由原生绘制器直接绘到长边 1600 像素、保持原始比例的画布，不使用缓存低清
缩略图。不支持的页面尺寸或缺失绘制接口会拒绝；绘制失败也恢复渲染标志、清理原生
共享字体缓存，返回前释放临时画布。

`document:render-page-preview` 接受 `{ binding, editGeneration, omitNativeText?, page? }`（可选标志必须为 boolean）。可选 `page` 为从 1 开始、最多 5000 且在源文档范围内的安全整数；直接渲染指定页，不切换当前显示页。省略时仍渲染当前页，返回 token 始终对应实际渲染页。
`document:page-preview` 返回源 SHA-256、打开会话内的页令牌、从 1
开始的页码、`image/png` MIME 和有界 PNG `Uint8Array`。请求必须使用最新代次及会话
绑定的源/页面 lineage；过时代次或页面、保存进行中、未知字段、原生渲染器不可用、
非 PNG 或超预算都会拒绝。此操作不编辑也不保存源文件。Host 可将图像作为独立单页
候选的已审阅视觉输入，并须把源和页面身份绑定到候选。

预览仅允许新打开的只读 PPTX 副本：原生撤销历史必须为空且无未保存更改。即便全部撤销也会拒绝，防止将改后的画面绑定到旧源 hash。编辑或保存后须以独立只读会话重新打开最新 clean PPTX，再获取源绑定预览。

设置 `omitNativeText:true` 后，仅临时绘制器跳过原生文本，返回额外的
`nativeTextOmitted:true`。仅支持指定页、版式和母版中的普通不透明形状及图片；
组合、图表、表格、变形文字、签名、透明填充及外阴影在绘制前拒绝。普通预览行为不变。
此操作不能去掉位图内嵌文字，**不自动证明清底图合格**；Host 采用为候选背景前
仍须审阅，不支持的页面须使用另外审阅的清底图。所有源/页面/代次/只读/保存
守卫及 1600 像素/8 MiB 限制保持不变。

### 讲解与动画关联

教学流程保留“添加动画”（淡入/淡出），新增卡片内“关联动画”多选下拉。
两者均支持本句开始、本句结束、所选文字“同时”和“之后”。TTS 尚不提供字词
时间戳，选词通过“试听定位”在真实音频进度标记开始或结束；修改讲解词后保留
绑定并提示重新定位。取消关联不删除原生动画。当前支持单个对象只有一个有限
原生效果的关联；多效果、循环、分组或不支持的预设在列表说明并禁用。
临时教学放映执行完动画后才进入讲后停顿，退出恢复原 PPT 动画与历史。

教学侧栏的状态行集中显示整课教学、当前页幻灯片放映、画布预览和本页教学说明
图标，均带名称与悬停提示；说明可点击展开。讲解输入框随内容高度变化，最大
240 像素后内部滚动。动画表单以“什么时候播放动画”说明讲解与动画先后关系，
操作按钮采用图标和短文字，保留完整可访问名称及悬停说明；窄面板下按按钮换行，避免中文逐字竖排。
摘要分行显示播放时机、可读对象名、中文效果与秒数。对象选择共用原生图层缩略图、已有中文别名和画布定位；无中文别名时显示实际文字或中文类型编号，英文语义描述仅保留在提示中，不猜测翻译。

### 教学放映控制

原生教学放映底部使用独立的深色控制条，按钮保持清晰的文字对比度。
提供播放/暂停切换、上一页、下一页、重播本页、教师继续、退出和当前媒体进度。
进度条可拖动，快进按钮前进 5 秒并限制在当前媒体范围内。
有同步原生动画的讲解仅允许向前调整进度，回看使用重播本页，以保持动画状态一致。
翻页会释放旧页媒体与动画再读取目标页；首尾页分别禁用上一页和下一页。
