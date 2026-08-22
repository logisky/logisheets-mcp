"""T11 — read a model you did not build, and explain it.

ADDED AFTER the pre-registered list in TASKS.md, and labelled as such: T1-T10
were fixed before any contestant was examined, this one was not. It exists
because the first five tasks all measured building and computing, and none of
them touched comprehension — which is where a schema is supposed to earn its
keep.

Each server reopens, in a fresh session, the file it saved in T2. That is the
durable-memory claim under test: not "can you build a model" but "can you come
back to one".

Three questions with checkable answers:
  Q1  how many projected years does the model have?          5
  Q2  how is free cash flow computed?      revenue x margin x (1 - tax)
  Q3  which inputs does value per share depend on?  the eight assumptions

And one measurement that matters more than the answers: does the model come
back described in ITS OWN terms, or as coordinates the agent has to resolve
against labels before they mean anything? A rule that reads
`#FIELD("rev") * BLOCKREF("assum","margin","v")` is an answer. One that reads
`B11*$B$3*(1-$B$4)` is a lookup problem — correct, but the agent now has to
fetch A3 and A4 to learn what it just read, and do that again next session.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mcp_client import Server                      # noqa: E402
from contestants import CONTESTANTS, WORK          # noqa: E402

ASSUMPTIONS = {"base", "growth", "margin", "tax", "wacc", "tg",
               "net_debt", "shares"}
A1_REF = re.compile(r"\$?[A-Z]{1,2}\$?\d{1,4}")


def names_or_coordinates(rule: str) -> str:
    """How the returned rule refers to things."""
    if rule is None:
        return "nothing returned"
    has_coords = bool(A1_REF.search(rule))
    has_names = any(w in rule for w in ("#FIELD", "BLOCKREF")) or \
        any(a in rule for a in ASSUMPTIONS)
    if has_names and not has_coords:
        return "names"
    if has_coords and not has_names:
        return "coordinates"
    return "mixed"


def run_logisheets(path: str):
    c = CONTESTANTS["logisheets"]
    out = {}
    with Server("logisheets", c["command"], c["env"]) as s:
        s.call("open_workbook", {"path": path})
        blocks, _ = s.call("list_blocks")
        # Q1: the projection's row count, stated as such.
        groups = blocks if isinstance(blocks, list) else []
        proj = None
        for g in groups:
            for b in g.get("blocks", []):
                if b.get("name") == "proj":
                    proj = b
        out["q1_years"] = proj and proj.get("key_count")
        # Q2: the field rule, verbatim, as the schema stores it.
        d, _ = s.call("describe_block", {"name": "proj"})
        fcf = next((f for f in (d or {}).get("fields", [])
                    if f["name"] == "fcf"), {})
        out["q2_rule"] = fcf.get("value_formula")
        # Q3: what the answer reads, from the engine's own graph.
        tr, _ = s.call("trace", {
            "target": {"block": "val", "row_key": "base", "field": "per_share"},
            "direction": "precedents"})
        out["q3_raw"] = tr
        out["calls"], out["bytes"] = s.calls, s.bytes
    return out


def run_spreadsheet_kit(path: str):
    c = CONTESTANTS["spreadsheet-kit"]
    base = os.path.basename(path)
    out = {}
    with Server("spreadsheet-kit", c["command"], c["env"]) as s:
        wbs, _ = s.call("list_workbooks", {"include_paths": True})
        wid = next((w["workbook_id"] for w in (wbs or {}).get("workbooks", [])
                    if w.get("path") == base), None)
        if wid is None:
            return {"error": f"no workbook_id for {base}"}
        ov, rawo = s.call("sheet_overview", {"workbook_or_fork_id": wid,
                                             "sheet_name": "Sheet"})
        # Region detection is genuinely useful on a file with no schema: it
        # picks A1:B8 out as `likely_parameters` at 0.70 confidence. But it
        # merges the projection and the valuation waterfall into one 11-row
        # `likely_outputs` block, so the number of projected years is not in
        # here — it has to be inferred from the cells.
        regions = (ov or {}).get("detected_regions", []) if isinstance(ov, dict) else []
        out["q1_regions"] = [(r.get("bounds"), r.get("region_kind"),
                              r.get("row_count")) for r in regions]
        out["q1_years"] = None
        fm, rawf = s.call("sheet_formula_map", {"workbook_or_fork_id": wid,
                                                "sheet_name": "Sheet",
                                                "include_addresses": True})
        out["formula_map"] = rawf[:300]
        # Q2: read the fcf cell itself.
        ins, _ = s.call("inspect_cells", {"workbook_or_fork_id": wid,
                                          "sheet_name": "Sheet",
                                          "targets": ["C11"]})
        cells = (ins or {}).get("cells", []) if isinstance(ins, dict) else []
        out["q2_rule"] = cells[0].get("formula") if cells else None
        # Q3: the engine's own trace.
        tr, rawt = s.call("formula_trace", {"workbook_or_fork_id": wid,
                                            "sheet_name": "Sheet",
                                            "cell_address": "B21",
                                            "direction": "precedents",
                                            "depth": 9})
        out["q3_raw"] = rawt[:400]
        out["calls_before_resolving"] = s.calls
        out["bytes_before_resolving"] = s.bytes
        # The rule came back as `B11*$B$3*(1-$B$4)`, which is correct and does
        # not yet mean anything. Turning it into domain terms costs another
        # read of the label column — and costs it again next session, because
        # nothing in the file records the answer.
        lab, _ = s.call("range_values", {"workbook_or_fork_id": wid,
                                         "sheet_name": "Sheet",
                                         "ranges": ["A1:A8", "A11:A11"]})
        out["labels"] = str(lab)[:200]
        out["calls"], out["bytes"] = s.calls, s.bytes
    return out


def run_excel_mcp(path: str):
    c = CONTESTANTS["excel-mcp-server"]
    out = {}
    with Server("excel-mcp-server", c["command"], c["env"]) as s:
        md, rawm = s.call("get_workbook_metadata", {"filepath": path})
        out["metadata"] = rawm[:200]
        got, raw = s.call("read_data_from_excel", {
            "filepath": path, "sheet_name": "Sheet",
            "start_cell": "A1", "end_cell": "F25"})
        cells = []
        if isinstance(got, dict) and isinstance(got.get("result"), str):
            import json as _json
            try:
                cells = _json.loads(got["result"]).get("cells", [])
            except Exception:
                pass
        by = {x["address"]: x.get("value") for x in cells}
        out["q2_rule"] = by.get("C11")
        out["q1_years"] = sum(1 for r in range(11, 20)
                              if isinstance(by.get(f"A{r}"), (int, float)))
        out["calls"], out["bytes"] = s.calls, s.bytes
    return out


if __name__ == "__main__":
    print("T11 — reopen the model you saved, and explain it. "
          "(added after the pre-registered list)\n")
    files = {
        "logisheets": os.path.join(WORK, "t2-logisheets.xlsx"),
        "spreadsheet-kit": os.path.join(WORK, "t2-spreadsheet-kit-saved.xlsx"),
        "excel-mcp-server": os.path.join(WORK, "t2-excel-mcp-server.xlsx"),
    }
    runners = {"logisheets": run_logisheets,
               "spreadsheet-kit": run_spreadsheet_kit,
               "excel-mcp-server": run_excel_mcp}
    for name, runner in runners.items():
        p = files[name]
        print(f"--- {name} ---")
        if not os.path.exists(p):
            print(f"    no saved file at {p}; run t2_dcf.py first\n")
            continue
        r = runner(p)
        if "error" in r:
            print(f"    {r['error']}\n")
            continue
        if r.get("q1_regions") is not None:
            print(f"    Q1 projected years : not reported — regions detected: "
                  f"{r['q1_regions']}")
        else:
            print(f"    Q1 projected years : {r.get('q1_years')}")
        rule = r.get("q2_rule")
        print(f"    Q2 fcf rule        : {rule}")
        print(f"                         -> refers by {names_or_coordinates(rule)}")
        q3 = r.get("q3_raw")
        print(f"    Q3 trace           : {str(q3)[:280]}")
        if r.get("calls_before_resolving") is not None:
            print(f"    cost               : {r['calls_before_resolving']} calls, "
                  f"{r['bytes_before_resolving']} bytes to read the model, then "
                  f"{r['calls'] - r['calls_before_resolving']} more call "
                  f"({r['bytes'] - r['bytes_before_resolving']} bytes) to learn "
                  f"what its coordinates mean")
        else:
            print(f"    cost               : {r.get('calls')} calls, "
                  f"{r.get('bytes')} bytes")
        print()
