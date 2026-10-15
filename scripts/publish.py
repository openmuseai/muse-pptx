"""Upload an MCPB and publish standard server.json metadata to a registry."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path
from urllib.request import Request, urlopen


def request(url: str, token: str, method: str, data: bytes) -> dict:
    with urlopen(Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json" if method == "POST" else "application/octet-stream",
    }), timeout=30) as response:
        return json.load(response)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("bundle", type=Path)
    parser.add_argument("--registry", required=True)
    parser.add_argument("--token", required=True)
    parser.add_argument("--client-version", default="latest",
                        help="MCP Host compatibility test target; defaults to latest")
    parser.add_argument("--client", choices=["desktop", "mobile"], default="desktop",
                        help="Public client compatibility target; defaults to desktop")
    parser.add_argument("--remote", help="Optional MCP Streamable HTTP endpoint for Mobile/Web")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[3]
    matrix = [item for item in json.loads((root / "compatibility/clients.json").read_text())["clients"]
              if item["client"] == args.client]
    selected = (max(matrix, key=lambda item: tuple(map(int, item["version"].split("."))))
                if args.client_version == "latest" else
                next((item for item in matrix if item["version"] == args.client_version), None))
    if selected is None:
        parser.error("Unknown client version; update the public compatibility matrix")
    subprocess.run(["npm", "test"], cwd=Path(__file__).resolve().parents[1], check=True)
    data = args.bundle.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    origin = args.registry.rstrip("/")
    package_url = f"{origin}/packages/{digest}.mcpb"
    request(package_url, args.token, "PUT", data)
    document = {
        "$schema": "https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json",
        "name": "io.openmuse/genoffice-pptx",
        "title": "GenOffice PPTX",
        "description": "Inspect PowerPoint slide structure using the GenOffice engine.",
        "version": "0.2.0",
        "repository": {"url": "https://github.com/genspark-ai/genoffice", "source": "github"},
        "packages": [{
            "registryType": "mcpb",
            "identifier": package_url,
            "fileSha256": digest,
            "transport": {"type": "stdio"},
        }],
    }
    if args.remote:
        document["remotes"] = [{"type": "streamable-http", "url": args.remote}]
    result = request(f"{origin}/v0.1/publish", args.token, "POST", json.dumps(document).encode())
    print(json.dumps({
        "submissionId": result["_meta"]["submissionId"],
        "state": result["_meta"]["status"],
        "clientVersionTested": selected["version"],
        "clientTested": selected["client"],
        "mcpProtocolTested": selected["mcpProtocol"],
        "server": document["name"],
        "sha256": digest,
    }, indent=2))


if __name__ == "__main__":
    main()
