"""A minimal MCP stdio client, so every contestant is driven the same way.

Deliberately not the official SDK: the point is to speak the wire protocol to
an arbitrary command, including servers written in other languages, and to see
exactly what each one returns rather than a normalized view of it.
"""
from __future__ import annotations

import json
import os
import subprocess
import time


class Server:
    """One MCP server, spawned as a child process and spoken to over stdio."""

    def __init__(self, name: str, command: list[str], env: dict | None = None,
                 cwd: str | None = None, timeout: float = 120.0) -> None:
        self.name = name
        self.command = command
        self.timeout = timeout
        self.calls = 0            # tool calls, the context-cost metric
        self.stderr_path = f"/tmp/bench-{name.replace('/', '_')}.stderr"
        self._id = 0
        self._stderr = open(self.stderr_path, "w")
        self.proc = subprocess.Popen(
            command,
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=self._stderr,
            text=True, bufsize=1, env={**os.environ, **(env or {})}, cwd=cwd,
        )

    # -- protocol ---------------------------------------------------------
    def _send(self, method: str, params=None, notify: bool = False):
        msg = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params
        if not notify:
            self._id += 1
            msg["id"] = self._id
        self.proc.stdin.write(json.dumps(msg) + "\n")
        self.proc.stdin.flush()
        if notify:
            return None
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            line = self.proc.stdout.readline()
            if line == "":
                raise RuntimeError(f"{self.name}: server closed stdout. "
                                   f"stderr in {self.stderr_path}")
            try:
                r = json.loads(line)
            except json.JSONDecodeError:
                continue          # some servers print banners on stdout
            if r.get("id") == self._id:
                return r
        raise TimeoutError(f"{self.name}: no reply to {method}")

    def initialize(self) -> dict:
        r = self._send("initialize", {
            "protocolVersion": "2025-06-18",
            "capabilities": {},
            "clientInfo": {"name": "logisheets-bench", "version": "0"},
        })
        self._send("notifications/initialized", {}, notify=True)
        return r.get("result", {})

    def tools(self) -> list[dict]:
        out, cursor = [], None
        while True:
            params = {"cursor": cursor} if cursor else {}
            r = self._send("tools/list", params).get("result", {})
            out.extend(r.get("tools", []))
            cursor = r.get("nextCursor")
            if not cursor:
                return out

    def call(self, name: str, args: dict | None = None):
        """Returns (payload, raw_text). payload is parsed JSON when the server
        returns any; otherwise the text. `None` payload means the server
        reported an error — the text carries its reason."""
        self.calls += 1
        r = self._send("tools/call", {"name": name, "arguments": args or {}})
        if "error" in r:
            return None, json.dumps(r["error"])
        res = r.get("result", {})
        text = "\n".join(c.get("text", "") for c in res.get("content", [])
                         if c.get("type") == "text")
        if res.get("isError"):
            return None, text
        if res.get("structuredContent") is not None:
            return res["structuredContent"], text
        # Servers embed JSON in text in every imaginable way; take the last
        # line that parses, else hand back the text.
        for line in reversed(text.split("\n")):
            for i, ch in enumerate(line):
                if ch in "[{":
                    try:
                        return json.loads(line[i:]), text
                    except json.JSONDecodeError:
                        pass
        return text, text

    def close(self) -> None:
        try:
            self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()
        finally:
            self._stderr.close()

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, *exc):
        self.close()
