"""Build a local Desktop Host installation fixture from the pinned PPTX worker.

This fixture is intentionally local and unsigned. Public distribution requires
the store signing and Host verification path described in the architecture doc.
"""

from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PLUGIN_ID = "com.openmuse.pptx"
VERSION = "0.1.0"


def digest_payload(files: dict[str, bytes]) -> str:
    digest = hashlib.sha256()
    for name in sorted(files):
        digest.update(name.encode())
        digest.update(b"\x00")
        digest.update(files[name])
        digest.update(b"\x00")
    return digest.hexdigest()


def build(output: Path) -> None:
    worker = (ROOT / "dist/worker.mjs").read_bytes()
    files = {"worker.mjs": worker}
    target = {"os": "macos", "arch": "aarch64", "libc": "darwin"}
    manifest = {
        "manifest_version": 2,
        "id": PLUGIN_ID,
        "name": "PPTX · GenOffice Core",
        "version": VERSION,
        "protocol": {"major": 1, "minor": 0},
        "ui_runtime": {"kind": "none"},
        "execution_connector": {
            "kind": "host-process",
            "protocol": "openmuse.pptx.worker/v1",
        },
        "compatibility": {"targets": [{"target": target, "status": "supported"}]},
        "artifacts": [
            {
                "id": "pptx-worker",
                "kind": "host-bundle",
                "target": target,
                "digest": {"algorithm": "sha256", "value": digest_payload(files)},
                "license": "Apache-2.0",
                "abi": "node-esm-22",
            }
        ],
        "activation_events": ["onCommand:pptx.inspect"],
        "requested_permissions": ["workspace.read"],
        "presentation": {"surfaces": [], "remote_capable": False},
        "contributes": {
            "commands": [
                {
                    "id": "pptx.inspect",
                    "title": "检查 PPTX 文件",
                    "required_permissions": ["workspace.read"],
                }
            ]
        },
        "install": {
            "steps": [
                {"type": "workspace.choose", "id": "pptx-workspace", "title": "PPTX 工作区"},
                {"type": "workspace.layout", "directories": ["PPTX"]},
                {
                    "type": "workspace.materialize",
                    "artifact": "pptx-worker",
                    "into": "PPTX",
                },
            ]
        },
    }
    output.mkdir(parents=True, exist_ok=True)
    package = output / f"{PLUGIN_ID}-{VERSION}-macos-aarch64.omplugin"
    entries = {
        "openmuse.plugin.json": json.dumps(manifest, ensure_ascii=False, separators=(",", ":")).encode(),
        "payload/pptx-worker/worker.mjs": worker,
    }
    with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in sorted(entries.items()):
            metadata = zipfile.ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
            metadata.compress_type = zipfile.ZIP_DEFLATED
            metadata.external_attr = 0o644 << 16
            archive.writestr(metadata, data)
    package_bytes = package.read_bytes()
    catalog = {
        "schema": "openmuse.local-catalog/v1",
        "packages": [
            {
                "id": PLUGIN_ID,
                "version": VERSION,
                "url": package.name,
                "sha256": hashlib.sha256(package_bytes).hexdigest(),
                "size": len(package_bytes),
            }
        ],
    }
    (output / "catalog.json").write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2) + "\n"
    )
    print(package)


if __name__ == "__main__":
    build(Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT / "dist")
