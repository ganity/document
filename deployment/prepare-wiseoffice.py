#!/usr/bin/env python3
"""Copy a document production build for the existing /wiseoffice/ mount.

From the document project root:
python3 deployment/prepare-wiseoffice.py dist /tmp/wiseoffice-release
Vendor files stay byte-for-byte unchanged. Run after document's full build.
"""

import re
import shutil
import sys
from pathlib import Path

source, destination = map(Path, sys.argv[1:])
prefix = "/wiseoffice"
shutil.copytree(source, destination)
roots = {path.name for path in source.iterdir()} | {"editor"}
vendor = {"sdkjs", "web-apps", "fonts", "libs", "plugins"}


def scope_path(match):
    quote, path = match.groups()
    if path.split("/", 2)[1].split("?", 1)[0] in roots:
        return quote + prefix + path + quote
    return match.group(0)


for path in destination.rglob("*"):
    relative = path.relative_to(destination)
    if relative.parts[0] in vendor or not path.is_file():
        continue
    if path.suffix not in {".html", ".js"}:
        continue
    text = path.read_text()
    if path.name == "sw.js":
        # Replacement suffixes are not URLs (e.g. index.html -> app.js).
        text = re.sub(
            r"(?<!\.replace\()(['\"`])(/[^'\"`\s]*?)\1", scope_path, text
        )
        text = text.replace("^\\/", "^\\/wiseoffice\\/")
        text = text.replace("url.pathname === '/'", "url.pathname === '/wiseoffice/'")
    else:
        text = re.sub(r"(['\"`])(/[^'\"`\s]*?)\1", scope_path, text)
        # The shared converter base also resolves SheetJS and font downloads.
        text = re.sub(
            r"return (\w+)\.startsWith\(`/document/`\)",
            rf"return \1.startsWith(`{prefix}/`)?`{prefix}/`:\1.startsWith(`/document/`)",
            text,
        )
    path.write_text(text)

# A source HTML stylesheet and compiled chunk must both resolve inside the mount.
html = (destination / "editor.html").read_text()
assert "./assets/editor-" in html
chunks = "".join(path.read_text() for path in (destination / "assets").glob("*.js"))
assert "new URL(`/wiseoffice/sw.js`" in chunks
assert "startsWith(`/wiseoffice/`)?`/wiseoffice/`:" in chunks
worker = (destination / "sw.js").read_text()
assert "'/wiseoffice/web-apps/apps/api/documents/api.js'" in worker
assert "path.replace('/index.html', '/wiseoffice/app.js')" not in worker
assert "^\\/wiseoffice\\/editor" in worker
print(f"Prepared {destination} for {prefix}/")
