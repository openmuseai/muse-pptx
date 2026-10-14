"""Produce the Android A2UI/QuickJS control bundle for Muse-Client.

The native Rust PPTX viewer is provided by the Host and is not downloadable.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "mobile/pptx.js").read_text()
IDENTITY = "com.openmuse.pptx"
VERSION = "0.1.0"
TARGET = {"os": "android", "arch": "aarch64", "libc": "bionic"}
MANIFEST = {
    "manifest_version": 2,
    "id": IDENTITY,
    "name": "PPTX 阅读器",
    "version": VERSION,
    "protocol": {"major": 1, "minor": 0},
    "ui_runtime": {"kind": "a2ui", "entrypoint": "main"},
    "execution_connector": {"kind": "quickjs", "protocol": "openmuse.quickjs.bridge/v1"},
    "compatibility": {"targets": [{"target": TARGET, "status": "supported"}]},
    "artifacts": [{
        "id": "js-android",
        "kind": "js-bundle",
        "target": TARGET,
        "digest": {"algorithm": "sha256", "value": hashlib.sha256(SOURCE.encode()).hexdigest()},
        "license": "Apache-2.0",
        "abi": "quickjs-es2020",
    }],
    "activation_events": ["onPanel:mobile.plugin"],
    "requested_permissions": [],
    "presentation": {"surfaces": ["editor"], "remote_capable": False},
    "contributes": {"panels": [{"id": f"mobile.plugin.{IDENTITY}", "region": "editor"}]},
}


def build(output: Path) -> tuple[Path, dict]:
    output.mkdir(parents=True, exist_ok=True)
    bundle = {
        "format": "openmuse-mobile-js/v1",
        "uiProtocol": "a2ui/v0.9",
        "engine": "quickjs",
        "id": IDENTITY,
        "version": VERSION,
        "js": SOURCE,
        "manifest": MANIFEST,
    }
    data = (json.dumps(bundle, ensure_ascii=False, separators=(",", ":")) + "\n").encode()
    path = output / "pptx-mobile.json"
    path.write_bytes(data)
    entry = {
        "id": IDENTITY,
        "name": "PPTX 阅读器",
        "version": VERSION,
        "description": "通过 OpenMuse Host 的受限 PPTX 引擎，在 Android Cloud Workspace 查看幻灯片文字。",
        "author": "OpenMuse · GenOffice Core",
        "platforms": ["android"],
        "source": "https://github.com/genspark-ai/genoffice",
        "categories": ["办公"],
        "media": [],
        "bundle": path.name,
        "sha256": hashlib.sha256(data).hexdigest(),
    }
    return path, entry


if __name__ == "__main__":
    path, entry = build(ROOT / "dist")
    print(path)
    print(json.dumps(entry, ensure_ascii=False))
