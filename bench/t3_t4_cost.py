"""T3 and T4 — the same answers, counted in round trips.

Both tasks are questions about a model rather than edits to it: a 4x4
sensitivity grid, and one inverse solve. Correctness is table stakes here — the
interesting number is how many tool calls each server needs, because context is
what an agent actually runs out of.

Every expected value comes from the Python model. Each server is given the best
route its own surface offers, which took some finding: spreadsheet-kit's
`execute_manifest` runs a whole scenario in one call, so using its
edit-recalculate-read triple sixteen times instead would have tripled its count
for no reason other than my ignorance of its API.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mcp_client import Server                      # noqa: E402
from contestants import CONTESTANTS, WORK          # noqa: E402
from t2_dcf import (A, a1_grid, blank_seed,        # noqa: E402
                    run_logisheets as build_logisheets)

WACCS = (0.09, 0.10, 0.11, 0.12)
TGS = (0.015, 0.020, 0.025, 0.030)
GOAL = 30.0


def per_share(wacc: float, tg: float) -> float:
    pvs, last = [], None
    for n in (1, 2, 3, 4, 5):
        rev = A["base"] * (1 + A["growth"]) ** n
        fcf = rev * A["margin"] * (1 - A["tax"])
        df = 1 / (1 + wacc) ** n
        pvs.append(fcf * df)
        last = (fcf, df)
    tv = last[0] * (1 + tg) / (wacc - tg) * last[1]
    return (sum(pvs) + tv - A["net_debt"]) / A["shares"]


def wacc_for(target: float) -> float:
    lo, hi = 0.05, 0.20
    for _ in range(200):
        mid = (lo + hi) / 2
        if per_share(mid, A["tg"]) > target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


GRID = {(w, t): per_share(w, t) for w in WACCS for t in TGS}
SOLUTION = wacc_for(GOAL)

MANIFEST = """spec: fio
spec_version: 0.3.0
capabilities:
  profile: core-v0
manifest:
  id: dcf
  name: dcf
  workbook:
    uri: file://fork
ports:
- id: wacc
  dir: in
  shape: scalar
  required: true
  location:
    a1: "'Sheet'!B5"
  schema:
    type: number
- id: tg
  dir: in
  shape: scalar
  required: true
  location:
    a1: "'Sheet'!B6"
  schema:
    type: number
- id: per_share
  dir: out
  shape: scalar
  required: true
  location:
    a1: "'Sheet'!B21"
  schema:
    type: number
"""


def sk_build(s: Server, path: str):
    base = os.path.basename(path)
    wbs, _ = s.call("list_workbooks", {"include_paths": True})
    wid = next((w["workbook_id"] for w in (wbs or {}).get("workbooks", [])
                if w.get("path") == base), None)
    if wid is None:
        raise RuntimeError(f"no workbook_id for {base}")
    fork, _ = s.call("create_fork", {"workbook_or_fork_id": wid})
    fid = fork["fork_id"]
    edits = [({"address": a, "formula": v}
              if isinstance(v, str) and v.startswith("=")
              else {"address": a, "value": str(v)}) for a, v in a1_grid()]
    s.call("edit_batch", {"fork_id": fid, "sheet_name": "Sheet",
                          "edits": edits})
    s.call("recalculate", {"fork_id": fid})
    return fid


def sk_run(s: Server, fid: str, wacc: float, tg: float):
    got, raw = s.call("execute_manifest", {
        "workbook_or_fork_id": fid, "manifest_yaml": MANIFEST,
        "inputs": {"wacc": wacc, "tg": tg}})
    outs = (got or {}).get("outputs", {}) if isinstance(got, dict) else {}
    return outs.get("per_share"), raw


# --- T3 ------------------------------------------------------------------

def t3_logisheets(s: Server, path: str):
    build_logisheets(s, path)
    base, base_bytes = s.calls, s.bytes
    # One call. Each scenario is applied to the temp branch independently and
    # only the watched cell comes back, so sixteen answers cost one round trip
    # and none of them touch the workbook.
    scenarios = [{
        "label": f"{w}/{t}",
        "changes": [{"block": "assum", "row_key": "wacc", "field": "v", "value": w},
                    {"block": "assum", "row_key": "tg", "field": "v", "value": t}],
    } for w in WACCS for t in TGS]
    got, raw = s.call("preview_changes", {
        "scenarios": scenarios,
        "watch": [{"block": "val", "row_key": "base",
                   "field": "per_share"}]})
    return got, (s.calls - base, s.bytes - base_bytes), raw


def t3_spreadsheet_kit(s: Server, path: str):
    fid = sk_build(s, path)
    base, base_bytes = s.calls, s.bytes
    out = {}
    for w in WACCS:
        for t in TGS:
            v, _ = sk_run(s, fid, w, t)
            out[(w, t)] = v
    return out, (s.calls - base, s.bytes - base_bytes), ""


def t3_excel_mcp(s: Server, path: str):
    # Nothing to run: this surface returns formula text, so no scenario can be
    # read back. Recorded rather than scored.
    return None, (0, 0), "no evaluation: a scenario cannot be read back"


# --- T4 ------------------------------------------------------------------

def t4_logisheets(s: Server, path: str):
    build_logisheets(s, path)
    base, base_bytes = s.calls, s.bytes
    got, raw = s.call("goal_seek", {
        "set": {"block": "assum", "row_key": "wacc", "field": "v"},
        "target": {"block": "val", "row_key": "base", "field": "per_share"},
        "to": GOAL, "between": [0.05, 0.20]})
    return got, (s.calls - base, s.bytes - base_bytes), raw


def t4_spreadsheet_kit(s: Server, path: str):
    fid = sk_build(s, path)
    base, base_bytes = s.calls, s.bytes
    lo, hi = 0.05, 0.20
    # Bisect by hand, which is what an agent without a solver has to do.
    while hi - lo > 1e-6:
        mid = (lo + hi) / 2
        v, _ = sk_run(s, fid, mid, A["tg"])
        if v is None:
            break
        if v > GOAL:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2, (s.calls - base, s.bytes - base_bytes), ""


if __name__ == "__main__":
    print("T3 — value per share across WACC x terminal growth, 4x4 = 16 answers")
    print(f"     expected corners: {GRID[(0.09, 0.015)]:.6f} .. "
          f"{GRID[(0.12, 0.030)]:.6f}\n")
    for name, runner in (("logisheets", t3_logisheets),
                         ("spreadsheet-kit", t3_spreadsheet_kit),
                         ("excel-mcp-server", t3_excel_mcp)):
        c = CONTESTANTS[name]
        path = os.path.join(WORK, f"t3-{name}.xlsx")
        if os.path.exists(path):
            os.remove(path)
        if name == "spreadsheet-kit":
            blank_seed(path)
        with Server(name, c["command"], c["env"]) as s:
            got, calls, raw = runner(s, path)
            if name == "spreadsheet-kit":
                bad = [k for k, v in got.items()
                       if not (isinstance(v, (int, float))
                               and abs(v - GRID[k]) < 1e-6)]
                print(f"--- {name} ---\n    16 answers: "
                      f"{'all correct' if not bad else f'{len(bad)} wrong'}")
            elif name == "logisheets":
                # Check all sixteen against the Python model. One call is only
                # interesting if it actually answered.
                rows = (got or {}).get("scenarios", []) if isinstance(got, dict) else []
                by_label = {}
                for r in rows:
                    vals = r.get("values") or []
                    by_label[r.get("label")] = vals[0] if vals else None
                bad = []
                for w in WACCS:
                    for t in TGS:
                        got_v = by_label.get(f"{w}/{t}")
                        if isinstance(got_v, dict):
                            got_v = got_v.get("value")
                        if not (isinstance(got_v, (int, float))
                                and abs(got_v - GRID[(w, t)]) < 1e-6):
                            bad.append((w, t, got_v))
                print(f"--- {name} ---\n    16 answers: "
                      f"{'all correct' if not bad else f'{len(bad)} wrong'}")
                if bad:
                    print(f"      first: {bad[0]}  expected "
                          f"{GRID[(bad[0][0], bad[0][1])]:.6f}")
                    print(f"      shape: {str(rows[:1])[:300]}")
            else:
                print(f"--- {name} ---\n    {raw}")
            n, nbytes = calls
            print(f"    cost: {n} tool call(s), {nbytes} bytes of output\n")

    print(f"T4 — which WACC makes value per share {GOAL}? "
          f"expected {SOLUTION:.6f}\n")
    for name, runner in (("logisheets", t4_logisheets),
                         ("spreadsheet-kit", t4_spreadsheet_kit)):
        c = CONTESTANTS[name]
        path = os.path.join(WORK, f"t4-{name}.xlsx")
        if os.path.exists(path):
            os.remove(path)
        if name == "spreadsheet-kit":
            blank_seed(path)
        with Server(name, c["command"], c["env"]) as s:
            got, calls, raw = runner(s, path)
            print(f"--- {name} ---")
            n, nbytes = calls
            print(f"    answer: {str(got)[:200]}")
            print(f"    cost: {n} tool call(s), {nbytes} bytes of output\n")
    print("--- excel-mcp-server ---\n    no evaluation: cannot search for an "
          "input it cannot measure\n")
