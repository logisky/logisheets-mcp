"""Who is in the comparison, and how to start each one.

Scoped to servers that operate on a local .xlsx. Google Sheets connectors are
a different product and are not here.
"""
import os

WORK = os.environ.get("BENCH_WORK", "/tmp/bench-work")
MCP_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CONTESTANTS = {
    # This project.
    "logisheets": {
        "command": ["node", f"{MCP_REPO}/dist/cli.js"],
        "env": {"LOGISHEETS_MCP_TOOLS": "core"},
        "note": "local Rust/WASM engine, 20-tool core surface",
    },
    # The nearest peer: its own native Rust recalc engine, formula tracing,
    # named ranges, event-sourced edits.
    # The nearest peer. Shipped as an amd64 OCI image; on Apple silicon it
    # runs under emulation, so its wall-clock times here mean nothing and are
    # not reported. The npm package of the same name is a CLI, not this server.
    "spreadsheet-kit": {
        "command": [
            "docker", "run", "-i", "--rm", "--platform", "linux/amd64",
            "-v", f"{WORK}:/data",
            "ghcr.io/psu3d0/spreadsheet-mcp:0.11.1-full",
            "--workspace-root", "/data", "--transport", "stdio",
        ],
        "env": {},
        "note": "native Rust recalc (Formualizer), optional LibreOffice; amd64 image",
    },
    # The most-installed Excel MCP server. openpyxl underneath.
    "excel-mcp-server": {
        "command": [f"{WORK}/.venv/bin/excel-mcp-server", "stdio"],
        "env": {"EXCEL_FILES_PATH": WORK},
        "note": "openpyxl; no formula evaluation by construction",
    },
}
