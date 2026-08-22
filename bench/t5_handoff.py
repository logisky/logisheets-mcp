"""T5 — the model changes shape, and the agent has to answer anyway.

The realistic version of this is a person editing the file in Excel, which
reflows every reference. There is no Excel or LibreOffice on this machine, so
the edits are made through each server's own structural tools instead. That is
a weaker proxy, and it is also more informative: it measures whether each
server reflows references the way a spreadsheet application would. Where a
server documents that it may not (spreadsheet-kit says so outright), the result
is reported, not held against it as a surprise.

The edits, in this order:
  1. delete the year-3 row from the projection
  2. insert two rows at the top   (a title and a spacer)
  3. insert one column at the left (somewhere to put notes)

Then: what is the value per share? Expected 19.383943020285496 — a Python model
over the four remaining years, terminal value from year 5.

The answer must be found by LABEL, never by a remembered address. That is the
whole point: after step 2 and 3 every address in the file has moved.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mcp_client import Server                      # noqa: E402
from contestants import CONTESTANTS, WORK          # noqa: E402
from t2_dcf import (A, a1_grid, blank_seed, col_row,  # noqa: E402
                    run_logisheets as build_logisheets)

EXPECTED = 19.383943020285496
YEAR3_ROW_1BASED = 13          # years occupy rows 11..15 before any edit


def expected_now() -> float:
    pvs, last = [], None
    for n in (1, 2, 4, 5):
        rev = A["base"] * (1 + A["growth"]) ** n
        fcf = rev * A["margin"] * (1 - A["tax"])
        df = 1 / (1 + A["wacc"]) ** n
        pvs.append(fcf * df)
        if n == 5:
            last = (fcf, df)
    tv = last[0] * (1 + A["tg"]) / (A["wacc"] - A["tg"]) * last[1]
    return (sum(pvs) + tv - A["net_debt"]) / A["shares"]


# --- logisheets ----------------------------------------------------------

def t5_logisheets(s: Server, path: str):
    build_logisheets(s, path)
    calls_after_build = s.calls
    # The row is named, so this needs no arithmetic and no knowledge of where
    # the projection currently sits.
    s.call("delete_block_rows", {"block": "proj", "keys": ["Y3"]})
    s.call("insert_rows", {"sheetIdx": 0, "start": 0, "count": 2})
    s.call("insert_cols", {"sheetIdx": 0, "start": 0, "count": 1})
    edit_calls = s.calls - calls_after_build
    # Answer by name. Nothing here depends on the shape of the sheet.
    got, raw = s.call("describe_block", {"name": "val", "include_rows": True})
    rows = (got or {}).get("rows", [])
    val = rows[0]["values"].get("per_share") if rows else None
    return val, s.calls - calls_after_build - edit_calls, edit_calls, raw


# --- excel-mcp-server ----------------------------------------------------

def t5_excel_mcp(s: Server, path: str):
    s.call("create_workbook", {"filepath": path})
    grid = {}
    for addr, content in a1_grid():
        c, r = col_row(addr)
        grid.setdefault(r, {})[c] = content
    rows = [[grid.get(r, {}).get(c) for c in range(5)] for r in range(21)]
    s.call("write_data_to_excel", {"filepath": path, "sheet_name": "Sheet",
                                   "data": rows, "start_cell": "A1"})
    calls_after_build = s.calls
    s.call("delete_sheet_rows", {"filepath": path, "sheet_name": "Sheet",
                                 "start_row": YEAR3_ROW_1BASED, "count": 1})
    s.call("insert_rows", {"filepath": path, "sheet_name": "Sheet",
                           "start_row": 1, "count": 2})
    s.call("insert_columns", {"filepath": path, "sheet_name": "Sheet",
                              "start_col": 1, "count": 1})
    edit_calls = s.calls - calls_after_build
    # No search tool in this surface, so the label has to be found by reading a
    # window and scanning it — which is the honest cost of no search.
    got, raw = s.call("read_data_from_excel", {
        "filepath": path, "sheet_name": "Sheet",
        "start_cell": "A1", "end_cell": "F30"})
    # No search tool in this surface, so the window has to be scanned. That is
    # the honest cost of not having one, and it is one call either way here.
    cells = []
    if isinstance(got, dict) and isinstance(got.get("result"), str):
        import json as _json
        try:
            cells = _json.loads(got["result"]).get("cells", [])
        except Exception:
            cells = []
    by_addr = {c["address"]: c.get("value") for c in cells}
    label = next((c["address"] for c in cells
                  if c.get("value") == "per_share"), None)
    val = None
    if label is not None:
        col = "".join(ch for ch in label if ch.isalpha())
        row = "".join(ch for ch in label if ch.isdigit())
        val = by_addr.get(f"{chr(ord(col[-1]) + 1)}{row}")
    return val, s.calls - calls_after_build - edit_calls, edit_calls, \
        f"label at {label}, beside it: {val!r}"


# --- spreadsheet-kit -----------------------------------------------------

def t5_spreadsheet_kit(s: Server, path: str):
    base = os.path.basename(path)
    wbs, rawl = s.call("list_workbooks", {"include_paths": True})
    wid = None
    for w in (wbs or {}).get("workbooks", []) if isinstance(wbs, dict) else []:
        if w.get("path") == base:
            wid = w.get("workbook_id")
    if wid is None:
        return None, 0, 0, f"no workbook_id for {base}: {rawl}"
    fork, rawf = s.call("create_fork", {"workbook_or_fork_id": wid})
    fid = (fork or {}).get("fork_id") if isinstance(fork, dict) else None
    if fid is None:
        return None, 0, 0, f"create_fork: {rawf}"
    edits = []
    for addr, content in a1_grid():
        if isinstance(content, str) and content.startswith("="):
            edits.append({"address": addr, "formula": content})
        else:
            edits.append({"address": addr, "value": str(content)})
    s.call("edit_batch", {"fork_id": fid, "sheet_name": "Sheet",
                          "edits": edits})
    s.call("recalculate", {"fork_id": fid})
    calls_after_build = s.calls
    s.call("structure_batch", {"fork_id": fid, "ops": [
        {"kind": "delete_rows", "sheet_name": "Sheet",
         "start_row": YEAR3_ROW_1BASED, "count": 1},
        {"kind": "insert_rows", "sheet_name": "Sheet", "at_row": 1, "count": 2},
        # A column letter, not an index — the server says so plainly if you
        # pass an integer, which is how this got found.
        {"kind": "insert_cols", "sheet_name": "Sheet", "at_col": "A", "count": 1},
    ]})
    s.call("recalculate", {"fork_id": fid})
    edit_calls = s.calls - calls_after_build
    # Find the label, then read the cell beside it. Two calls, and no
    # assumption about where anything ended up.
    found, rawv = s.call("find_value", {"workbook_or_fork_id": fid,
                                        "query": "per_share"})
    matches = (found or {}).get("matches", []) if isinstance(found, dict) else []
    if not matches:
        return None, s.calls - calls_after_build - edit_calls, edit_calls, \
            f"label not found: {rawv[:200]}"
    addr = matches[0]["address"]
    col = "".join(ch for ch in addr if ch.isalpha())
    row = "".join(ch for ch in addr if ch.isdigit())
    beside = f"{chr(ord(col[-1]) + 1)}{row}" if len(col) == 1 else None
    got, rawc = s.call("inspect_cells", {"workbook_or_fork_id": fid,
                                         "sheet_name": "Sheet",
                                         "targets": [beside]})
    cells = (got or {}).get("cells", []) if isinstance(got, dict) else []
    val = None
    if cells:
        v = cells[0].get("value")
        val = v.get("value") if isinstance(v, dict) else v
    return val, s.calls - calls_after_build - edit_calls, edit_calls, \
        f"label at {addr}, read {beside}: {rawc[:200]}"


RUNNERS = {
    "logisheets": (t5_logisheets, {"LOGISHEETS_MCP_TOOLS": "full"}),
    "spreadsheet-kit": (t5_spreadsheet_kit, {}),
    "excel-mcp-server": (t5_excel_mcp, {}),
}

if __name__ == "__main__":
    want = expected_now()
    assert abs(want - EXPECTED) < 1e-12
    print(f"T5 — delete year 3, insert 2 rows at the top and a column at the "
          f"left, then answer.\n     Expected {want!r}\n")
    for name, (runner, env) in RUNNERS.items():
        c = CONTESTANTS[name]
        path = os.path.join(WORK, f"t5-{name}.xlsx")
        if os.path.exists(path):
            os.remove(path)
        if name == "spreadsheet-kit":
            blank_seed(path)
        try:
            with Server(name, c["command"], {**c["env"], **env}) as s:
                val, read_calls, edit_calls, raw = runner(s, path)
                ok = isinstance(val, (int, float)) and abs(val - want) < 1e-6
                print(f"--- {name} ---")
                print(f"    answer     : {val if not isinstance(val, float) else round(val, 6)}"
                      f"  {'✓' if ok else '✗'}")
                print(f"    calls      : {edit_calls} to edit, {read_calls} to "
                      f"find the answer again")
                if not ok:
                    print(f"    raw        : {' '.join(str(raw).split())[:400]}")
        except Exception as e:
            print(f"--- {name} ---\n    failed: {type(e).__name__}: {e}")
        print()

    # What the edited files actually contain. A server that cannot evaluate
    # cannot tell you it broke the model, so the only way to know is to compute
    # the file somewhere that can.
    print("Reopening each edited file in an engine — what does the model say now?")
    c = CONTESTANTS["logisheets"]
    for name in ("excel-mcp-server",):
        path = os.path.join(WORK, f"t5-{name}.xlsx")
        if not os.path.exists(path):
            print(f"  {name}: no file")
            continue
        with Server("logisheets", c["command"], c["env"]) as s:
            s.call("open_workbook", {"path": path})
            got, _ = s.call("get_cells", {"sheetIdx": 1, "startRow": 0,
                                          "startCol": 0, "endRow": 25,
                                          "endCol": 5})
            cells = [x for x in (got or {}).get("cells", [])
                     if x.get("value") not in (None, "")]
            label = next((x for x in cells if x.get("value") == "per_share"), None)
            print(f"  {name}: label at {label and label['ref']}")
            for x in cells[-4:]:
                print(f"    {x['ref']:>4}  value={x.get('value')!r:<14} "
                      f"formula={x.get('formula')!r}")
