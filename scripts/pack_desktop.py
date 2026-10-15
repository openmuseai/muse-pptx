"""Package the standalone GenOffice MCP server as a standard MCP Bundle."""

from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NAME = "genoffice-pptx"
VERSION = "0.2.0"


def build(output: Path) -> Path:
    output.mkdir(parents=True, exist_ok=True)
    worker = (ROOT / "dist/server.mjs").read_bytes()
    manifest = {
        "manifest_version": "0.3",
        "name": NAME,
        "display_name": "GenOffice PPTX",
        "version": VERSION,
        "description": "Inspect a user-selected PowerPoint deck with GenOffice.",
        "author": {"name": "OpenMuse contributors"},
        "license": "Apache-2.0",
        "repository": {"type": "git", "url": "https://github.com/genspark-ai/genoffice"},
        "tools": [{"name": "inspect_pptx", "description": "Inspect a PPTX deck"}],
        "compatibility": {
            "platforms": ["darwin", "win32", "linux"],
            "runtimes": {"node": ">=22.0.0"},
        },
        "server": {
            "type": "node",
            "entry_point": "server/index.mjs",
            "mcp_config": {
                "command": "node",
                "args": ["${__dirname}/server/index.mjs"],
            },
        },
    }
    package = output / f"{NAME}-{VERSION}.mcpb"
    entries = {
        "manifest.json": (json.dumps(manifest, separators=(",", ":")) + "\n").encode(),
        "server/index.mjs": worker,
        "LICENSE": (ROOT / "UPSTREAM-LICENSE").read_bytes(),
        "NOTICE": (ROOT / "UPSTREAM-NOTICE").read_bytes(),
    }
    with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in sorted(entries.items()):
            info = zipfile.ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, data)
    digest = hashlib.sha256(package.read_bytes()).hexdigest()
    (output / "mcpb.sha256").write_text(f"{digest}  {package.name}\n")
    return package


if __name__ == "__main__":
    print(build(Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT / "dist"))
