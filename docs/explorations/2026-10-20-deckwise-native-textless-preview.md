# Native text omission for reviewed page candidates

The existing source-page renderer now accepts one optional boolean, `omitNativeText`. The pinned SDK already uses a graphics-instance `isSupportTextDraw` override for native layout thumbnails. A real Office probe confirms ordinary text disappears while picture pixels remain, without native history or source writes. The override belongs to the temporary graphics instance; later normal rendering and editing retain text.

The SDK also uses that flag when compositing translucent fills, and warped text follows another path. The textless path therefore requires ordinary shape/image types across page, layout and master, opaque fills, no outer shadows, and no warped/Word text or signature representation. Unsupported content rejects before canvas allocation. Normal previews retain their previous broader native-render behavior. All existing source/page/generation/readonly/empty-history/save fences and cleanup apply.

This produces a reviewable native-textless visual, not a semantic raster text-removal service. Pictures, including any burned-in text, remain. The Host must review the output before using it as a clean plate; C retains media-generation ownership. Main/coordinator/UI mounting belongs to A. No new Chat tool or assistant is introduced.
