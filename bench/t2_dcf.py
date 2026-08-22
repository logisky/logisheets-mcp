"""T2 — build a five-year DCF and read the value per share.

The expected answer is computed here, in Python, from the same inputs. No
server's output feeds it.

Fairness rules, applied uniformly:
  * each server is driven the way its own tool descriptions steer you — the
    A1-based ones get the identical cell grid, built once below, so neither is
    handicapped by a layout the other did not get;
  * every server's batch facility is used where it has one, since tool calls
    are the scarce resource and it would be easy to inflate a rival's count by
    refusing to batch;
  * a server that cannot evaluate still gets credit for the file it wrote: we
    reopen it in an engine and check the model is right, so "cannot answer in
    session" is not confused with "cannot build the model".
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mcp_client import Server                      # noqa: E402
from contestants import CONTESTANTS, WORK          # noqa: E402

A = dict(base=1000.0, growth=0.08, margin=0.20, tax=0.25,
         wacc=0.10, tg=0.025, net_debt=500.0, shares=100.0)
YEARS = [1, 2, 3, 4, 5]


def expected(**over) -> float:
    a = {**A, **over}
    pvs, last = [], None
    for n in YEARS:
        rev = a["base"] * (1 + a["growth"]) ** n
        fcf = rev * a["margin"] * (1 - a["tax"])
        df = 1 / (1 + a["wacc"]) ** n
        pvs.append(fcf * df)
        last = (fcf, df)
    tv = last[0] * (1 + a["tg"]) / (a["wacc"] - a["tg"]) * last[1]
    ev = sum(pvs) + tv
    return (ev - a["net_debt"]) / a["shares"]


ORDER = ["base", "growth", "margin", "tax", "wacc", "tg", "net_debt", "shares"]

def a1_grid() -> list[tuple[str, object]]:
    """The same model, as (address, content) pairs, for every A1-based server."""
    cells: list[tuple[str, object]] = []
    for i, k in enumerate(ORDER, start=1):
        cells.append((f"A{i}", k))
        cells.append((f"B{i}", A[k]))
    for row, n in zip(range(11, 16), YEARS):
        cells += [
            (f"A{row}", n),
            (f"B{row}", f"=$B$1*POWER(1+$B$2,A{row})"),
            (f"C{row}", f"=B{row}*$B$3*(1-$B$4)"),
            (f"D{row}", f"=1/POWER(1+$B$5,A{row})"),
            (f"E{row}", f"=C{row}*D{row}"),
        ]
    cells += [
        ("A17", "pv_explicit"), ("B17", "=SUM(E11:E15)"),
        ("A18", "pv_terminal"), ("B18", "=C15*(1+$B$6)/($B$5-$B$6)*D15"),
        ("A19", "enterprise_value"), ("B19", "=B17+B18"),
        ("A20", "equity"), ("B20", "=B19-$B$7"),
        ("A21", "per_share"), ("B21", "=B20/$B$8"),
    ]
    return cells


def col_row(addr: str) -> tuple[int, int]:
    letters = "".join(c for c in addr if c.isalpha())
    digits = int("".join(c for c in addr if c.isdigit()))
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return col - 1, digits - 1


# --- contestants ---------------------------------------------------------

def run_logisheets(s: Server, path: str):
    """Blocks and field rules — the model this server is built around. The
    rules are stated once and apply to every row, which is why this costs more
    calls up front than writing cells; T5 is where that is meant to pay off."""
    s.call("open_workbook")
    s.call("create_block", {
        "sheet": "M", "name": "assum", "position": {"row": 0, "col": 0},
        "fields": [{"name": "k"}, {"name": "v", "field_type": "number"}],
        "initial_rows": [{"key": k, "values": {"v": A[k]}} for k in ORDER]})
    s.call("create_block", {
        "sheet": "M", "name": "proj", "position": {"row": 11, "col": 0},
        "fields": [{"name": "yr"}, {"name": "n", "field_type": "number"},
                   {"name": "rev", "field_type": "number"},
                   {"name": "fcf", "field_type": "number"},
                   {"name": "df", "field_type": "number"},
                   {"name": "pv", "field_type": "number"}],
        "initial_rows": [{"key": f"Y{n}", "values": {"n": n}} for n in YEARS]})
    ref = lambda k: f'BLOCKREF("assum","{k}","v")'
    for f, rule in [
        ("rev", f'={ref("base")}*POWER(1+{ref("growth")},#FIELD("n"))'),
        ("fcf", f'=#FIELD("rev")*{ref("margin")}*(1-{ref("tax")})'),
        ("df", f'=1/POWER(1+{ref("wacc")},#FIELD("n"))'),
        ("pv", '=#FIELD("fcf")*#FIELD("df")'),
    ]:
        s.call("set_field_rule", {"block": "proj", "field": f,
                                  "value_formula": rule})
    s.call("create_block", {
        "sheet": "M", "name": "val", "position": {"row": 20, "col": 0},
        "fields": [{"name": "case"}, {"name": "pv_explicit", "field_type": "number"},
                   {"name": "pv_terminal", "field_type": "number"},
                   {"name": "ev", "field_type": "number"},
                   {"name": "equity", "field_type": "number"},
                   {"name": "per_share", "field_type": "number"}],
        "initial_rows": [{"key": "base"}]})
    p = lambda f: f'BLOCKREF("proj","Y5","{f}")'
    for f, rule in [
        ("pv_explicit", '=SUM(BLOCKREFS("proj","*","pv"))'),
        ("pv_terminal", f'={p("fcf")}*(1+{ref("tg")})/({ref("wacc")}-{ref("tg")})*{p("df")}'),
        ("ev", '=#FIELD("pv_explicit")+#FIELD("pv_terminal")'),
        ("equity", f'=#FIELD("ev")-{ref("net_debt")}'),
        ("per_share", f'=#FIELD("equity")/{ref("shares")}'),
    ]:
        s.call("set_field_rule", {"block": "val", "field": f,
                                  "value_formula": rule})
    got, raw = s.call("describe_block", {"name": "val", "include_rows": True})
    rows = (got or {}).get("rows", [])
    s.call("save_workbook", {"path": path})
    return (rows[0]["values"].get("per_share") if rows else None), raw


def run_excel_mcp(s: Server, path: str):
    """`write_data_to_excel` takes formulas too, so the whole model goes in one
    call rather than 30 `apply_formula`s."""
    s.call("create_workbook", {"filepath": path})
    grid = {}
    for addr, content in a1_grid():
        c, r = col_row(addr)
        grid.setdefault(r, {})[c] = content
    rows = []
    for r in range(0, 21):
        row = grid.get(r, {})
        rows.append([row.get(c, None) for c in range(0, 5)])
    s.call("write_data_to_excel", {"filepath": path, "sheet_name": "Sheet",
                                   "data": rows, "start_cell": "A1"})
    got, raw = s.call("read_data_from_excel", {
        "filepath": path, "sheet_name": "Sheet",
        "start_cell": "B21", "end_cell": "B21"})
    return got, raw


def run_spreadsheet_kit(s: Server, path: str):
    base = os.path.basename(path)
    wbs, rawl = s.call("list_workbooks", {"include_paths": True})
    wid = None
    for w in (wbs or {}).get("workbooks", []) if isinstance(wbs, dict) else []:
        if w.get("path") == base:
            wid = w.get("workbook_id")
    if wid is None:
        return None, f"no workbook_id for {base}: {rawl}"
    fork, rawf = s.call("create_fork", {"workbook_or_fork_id": wid})
    fid = (fork or {}).get("fork_id") if isinstance(fork, dict) else None
    if fid is None:
        return None, f"create_fork: {rawf}"
    edits = []
    for addr, content in a1_grid():
        if isinstance(content, str) and content.startswith("="):
            edits.append({"address": addr, "formula": content})
        else:
            edits.append({"address": addr, "value": str(content)})
    _, rawe = s.call("edit_batch", {"fork_id": fid, "sheet_name": "Sheet",
                                    "edits": edits})
    _, rawr = s.call("recalculate", {"fork_id": fid})
    got, raw = s.call("inspect_cells", {"workbook_or_fork_id": fid,
                                        "sheet_name": "Sheet",
                                        "targets": ["B21"]})
    cells = (got or {}).get("cells", []) if isinstance(got, dict) else []
    val = None
    if cells:
        v = cells[0].get("value")
        val = v.get("value") if isinstance(v, dict) else v
    # A fork is scratch space; persist it so the file can be checked like the
    # others'.
    # Two things to get right here. Inside the container the workspace is
    # mounted at /data, so that is the path the server writes to. And it
    # refuses to overwrite the workbook it forked from unless told otherwise —
    # a safety default worth noting, since ours will write wherever it is
    # pointed — so the model goes to a new name.
    saved = os.path.basename(path).replace(".xlsx", "-saved.xlsx")
    _, raws = s.call("save_fork", {"fork_id": fid,
                                   "target_path": f"/data/{saved}"})
    return val, f"{rawe[:120]}\n---\n{rawr}\n---\n{raw[:200]}\n--- save: {raws[:200]}"


def blank_seed(path: str) -> None:
    from openpyxl import Workbook
    Workbook().save(path)


RUNNERS = {
    "logisheets": run_logisheets,
    "spreadsheet-kit": run_spreadsheet_kit,
    "excel-mcp-server": run_excel_mcp,
}

if __name__ == "__main__":
    want = expected()
    print(f"T2 — five-year DCF, value per share. Expected {want:.6f} "
          f"(Python, from the same inputs)\n")
    results = {}
    for name, runner in RUNNERS.items():
        c = CONTESTANTS[name]
        path = os.path.join(WORK, f"t2-{name}.xlsx")
        if os.path.exists(path):
            os.remove(path)
        if name == "spreadsheet-kit":
            blank_seed(path)            # it edits existing workbooks only
        try:
            with Server(name, c["command"], c["env"]) as s:
                val, raw = runner(s, path)
                ok = isinstance(val, (int, float)) and abs(val - want) < 1e-6
                results[name] = (val, s.calls, ok, path)
                print(f"--- {name} ---")
                print(f"    per_share : {val if not isinstance(val, float) else round(val, 6)}"
                      f"   {'✓' if ok else '✗'}")
                print(f"    tool calls: {s.calls}")
                if not ok:
                    print(f"    raw       : {' '.join(str(raw).split())[:260]}")
        except Exception as e:
            print(f"--- {name} ---\n    failed: {type(e).__name__}: {e}")
        print()

    # A server that cannot evaluate may still have written a correct model, and
    # that distinction is worth more than the in-session answer for anyone who
    # only wants a file. Reopen each saved file in an engine and look.
    #
    # `write_data_to_excel` with a sheet_name the workbook does not have adds a
    # sheet, so excel-mcp-server's model lands on sheet 1 next to the empty
    # default at sheet 0 — reading sheet 0 would have reported an empty file
    # and blamed the server for it.
    print("Reopening each saved file in the logisheets engine, to separate "
          "'cannot answer' from 'cannot build':")
    where = {"excel-mcp-server": 1, "spreadsheet-kit": 0}
    saved_as = {"spreadsheet-kit": lambda p: p.replace(".xlsx", "-saved.xlsx")}
    c = CONTESTANTS["logisheets"]
    for name, (_, _, _, path) in results.items():
        if name == "logisheets":
            continue                    # already answered in session
        path = saved_as.get(name, lambda p: p)(path)
        if not os.path.exists(path):
            print(f"  {name:<18} no file written")
            continue
        with Server("logisheets", c["command"], c["env"]) as s:
            _, raw = s.call("open_workbook", {"path": path})
            got, _ = s.call("get_cells", {
                "sheetIdx": where[name], "startRow": 20, "startCol": 1,
                "endRow": 20, "endCol": 1})
            cells = (got or {}).get("cells", [])
            v = cells[0].get("value") if cells else None
            ok = isinstance(v, (int, float)) and abs(v - want) < 1e-6
            print(f"  {name:<18} B21 = {v!r}  {'✓ the model is right' if ok else '—'}")
            if not ok:
                print(f"                     {' '.join(str(raw).split())[:150]}")
