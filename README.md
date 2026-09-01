# logisheets-mcp

[![logsky/logisheets-mcp](https://glama.ai/mcp/servers/logisky/logisheets-mcp/badges/score.svg)](https://glama.ai/mcp/servers/logisky/logisheets-mcp)

**A real spreadsheet engine your agent can think in.** Excel-compatible formulas
it doesn't have to do in its head, a table it addresses by name instead of by
coordinate, and a genuine `.xlsx` at the end that a person can open, audit and
keep using.

An [MCP](https://modelcontextprotocol.io) server over
[LogiSheets](https://github.com/logisky/LogiSheets), a spreadsheet engine written
in Rust. MIT, runs on your machine, opens no sockets.

## The trouble with a grid

Ask a model for a five-year projection and it writes twenty formulas, each with
the row number adjusted by hand. That is where the silent mistake lives: one of
them reads `B7` where it meant `B8`, the total looks plausible, and nothing
raises an error.

Then the sheet moves. Someone inserts a row at the top, deletes a year, adds a
column. Every coordinate the model was holding is now off by one and it has no
way to notice, so it spends the next turns re-reading cells to work out where
things went instead of on the question you asked.

And every "what if" costs a round trip — write the input, recalculate, read the
output, put it back. Sixteen scenarios is sixteen of those, and a scan that dies
half way leaves a scenario behind in your model.

## Blocks

A **block** is a named table on the sheet. Rows have keys, columns have names,
and everything is addressed by those rather than by position.

- **A field's formula is stated once, for the whole column** — not per cell. Add
  a row and it computes. There is no twentieth formula to get wrong.
- **A reference names what it means**: *the `pv` field of the row keyed `Y3`*.
  Insert a row above it and the reference still says the same thing, because it
  never said "row 8".
- **The engine owns computed values.** A formula field cannot be overwritten with
  a number the model worked out itself.

```
create_block     proj    fields: year, fcf, df, pv
set_field_rule   proj.pv = fcf × df          ← once, for the column
add_block_rows   Y1 … Y5
describe_block   proj
  →  Y1 147.2727   Y2 144.5950   Y3 141.9660   Y4 139.3848   Y5 136.8506

… the sheet is then reshaped: two rows inserted at the top, a column at the left …

describe_block   proj
  →  Y3 141.9660               ← same answer, same address, nothing re-derived
```

Blocks are created by the agent as it works, so nothing needs preparing. Point it
at a blank workbook or at a spreadsheet someone emailed you — `convert_to_block`
adopts a table that is already in ordinary cells, reading the field names off the
header row and working out which column is the key.

## The second session

The conversation that builds a spreadsheet is almost never the conversation that
has to answer a question about it. A week later there is a new session, with
none of the context, holding only the file — and what the file records is what
that session can know.

A grid records coordinates. `=B11*$B$3*(1-$B$4)` is correct and means nothing
until the agent fetches the label column and *infers* that `A3` describes `B3`.
The schema is where the meaning goes instead, and it is written into the
`.xlsx`: field names, the key column, which fields the engine computes, and the
rule behind each one. One `list_blocks` call and the workbook introduces itself;
one `describe_block` and the rules come back as
`#FIELD("revenue")*BLOCKREF("assum","margin","v")` — an explanation rather than
a second lookup problem.

[`src/cold-read.test.ts`](src/cold-read.test.ts) pins that down rather than
asserting it. It builds a model in one session, saves it, and reopens the file
in a second session sharing nothing with the first — own server, own workbook,
no memory. Then: `list_blocks` recovers every block's fields, key field,
computed fields and row count in one call; every returned rule is checked to
contain `#FIELD` or `BLOCKREF` and **no A1 coordinate at all**; the fresh
session writes a `BLOCKREF` formula from orientation alone and the engine agrees
with arithmetic done independently in the test; and `trace` names what reads an
assumption before anyone edits it. Cost is metered on the wire, over the same
text a host shows the model: **540 B for a five-row model, 545 B for a
hundred-and-five-row one**, one call each. Reading a schema is `O(columns)`;
reading a grid to understand it is `O(cells)`. Asking for the data still costs
what the data costs — 11 kB for those 105 rows — and the point is that the
second session gets to choose.

Longer version, with the reasoning: [`docs/the-second-session.md`](docs/the-second-session.md).

## Benchmarks

Measured, not asserted. Against the two other MCP servers that work on a local
`.xlsx` — [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) 0.11.1,
which has its own Rust recalc engine, and
[excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) 0.1.8, the
most-installed one, on openpyxl:

| | this | spreadsheet-kit | excel-mcp-server |
| --- | --- | --- | --- |
| Write a formula, read its value | **30** | **30** | `"=SUM(A1:A2)"` |
| Five-year DCF, value per share | **20.803603** · 15 calls | **20.803603** · 6 calls | formula text |
| 4×4 sensitivity, 16 answers | **1 call**, 950 B | 16 calls, 1245 B | can't |
| Solve backwards for an input | **1 call**, 202 B | 18 calls, 1399 B | can't |
| Reopen it later and explain it | 4 calls, **2.4 kB** | 5 calls, 21 kB | 2 calls, 24 kB |
| Answer again after the shape changed | **19.383943** | `#VALUE!` | formula text |
| Keep a handed-over file's features | **8 of 8** | **8 of 8** | **8 of 8** |

Reproduce it — one file per task, and each one runs all three servers:

```bash
npm run build                     # ours is driven as dist/cli.js
python3 bench/t1_compute.py       # bench/t*.py
```

The other two contestants have to be reachable first: spreadsheet-kit as an
amd64 Docker image, excel-mcp-server in a virtualenv at `$BENCH_WORK/.venv`
(default `/tmp/bench-work`). See [`bench/contestants.py`](bench/contestants.py)
for exactly how each is started.

The tasks were committed *before* any other server's tool list was read
([`bench/TASKS.md`](bench/TASKS.md)), every expected value is derived
independently in Python rather than read off a server's output, and tasks we
expected to lose are in the list on purpose.

Three caveats, so the table is not read for more than it says. `"=SUM(A1:A2)"` is
not a bug: openpyxl stores formulas without evaluating them, so that server
writes correct models but cannot answer a question about one. spreadsheet-kit is
a genuine peer, correct on everything it can attempt, and builds the model in
fewer calls than we do — our extra calls declare a schema, which is the trade
that pays off in the rows below. And on the reading row each server was reading
back a file *it* wrote, so only half of that margin transfers to a spreadsheet
that came from a person. The last row started at 0 of 8; writing the task is what
found that saves were dropping everything the engine had no opinion about.

## Install

Requires Node 20+.

```bash
npm install -g logisheets-mcp
```

For Claude Desktop, add to `claude_desktop_config.json` (macOS:
`~/Library/Application Support/Claude/claude_desktop_config.json`; Windows:
`%APPDATA%\Claude\claude_desktop_config.json`), then restart:

```json
{
    "mcpServers": {
        "logisheets": {
            "command": "npx",
            "args": ["-y", "logisheets-mcp"]
        }
    }
}
```

Any MCP host that spawns a stdio server works the same way — Cursor reads the
same block from `~/.cursor/mcp.json`.

## Try it

> Build me a three-year revenue model: 100 units at $9.50 growing 40% a year,
> with a 30% cost of goods. Then save it to ~/model.xlsx.

The numbers come back from the engine rather than from the model's guesses, and
the `.xlsx` has live formulas in it — change an assumption in Excel and watch it
recompute. To see the same thing with no LLM involved,
`npm run build && npm run demo` drives the real server over stdio and checks
every claim as it goes.

## Tools

Twenty by default. Tool-selection accuracy falls as the list grows and every
description costs context on every turn.

| Tool | What it does |
| --- | --- |
| `open_workbook` | Start a fresh workbook, or load an existing `.xlsx`. Optional — one appears on first use. |
| `save_workbook` | Write a real `.xlsx`. This is how work gets handed back. |
| `export_xlsx` | The file as base64, for hosts with no shared filesystem. |
| `list_blocks` | Every sheet and block, plus where the next block should go. |
| `describe_block` | A block's schema, keys, field rules, and optionally its values. |
| `eval_formula` | Evaluate a formula and return the value. Nothing is stored. |
| `create_block` | Create a named table. First field is the row key. |
| `convert_to_block` | Adopt a table that is already in ordinary cells, in place. |
| `add_block_rows` | Add records — at the end, or `after_key` / `before_key`. |
| `delete_block_rows` | Remove records. |
| `move_block_row` | Reorder rows by key. Presentation only: no value changes. |
| `set_block_cells` | Write cells by `(block, row_key, field)`. Batched, atomic. |
| `set_field_rule` | Give a field a formula, a validation rule, or an editability rule. |
| `list_violations` | Which cells break their field's validation rule, and why. |
| `preview_changes` | What edits *would* do, without doing them — one hypothetical, or a whole grid of scenarios in a single call. |
| `trace` | What a cell reads, and what reads it, from the dependency graph. |
| `goal_seek` | What input makes an output hit a target. Searches inside the engine. |
| `create_sheet` | Add a sheet. |
| `get_cells` / `set_cells` | Raw-cell escape hatch for data with no structure. |

Formulas are Excel-compatible plus `BLOCKREF(block, key, field)` for reading a
block cell by name. Inside a field rule, `#FIELD("name")` is the same row's
sibling and `#FIELD("name", "key")` is another row of the same block — the row
carrying that key, never a positional offset.

`preview_changes` and `goal_seek` are the two that change how a model gets
explored: each scenario runs on its own temp branch and is discarded, so a 4×4
sensitivity grid is one call returning sixteen numbers with nothing written to
the workbook, and an inverse solve is one call rather than one per bisection
step. `trace` answers the question formula text cannot — not what a cell reads,
but what reads *it*, which is what you want before touching an assumption.

Set `LOGISHEETS_MCP_TOOLS=full` for 50: undo/redo, formatting, merges, comments,
checkpoints, block move/resize, cross-block links, raw row/column structure.
Mutating tools carry MCP's `readOnlyHint` / `destructiveHint` annotations so a
host can gate them behind approval.

## The file you get back

`save_workbook` writes a real `.xlsx` and returns an MCP **resource link** — a
uri, media type and size — rather than the bytes, which would cost ~280 KB of
context for a 200 KB workbook and teach the model nothing. Hosts that want the
file read it from `workbook://current.xlsx`; `export_xlsx` returns base64 for
hosts implementing no resources at all.

Formulas can be written out as `BLOCKREF("proj","Y3","pv")` for a person to read,
or resolved to plain coordinates for Excel to chew on.

One MCP session holds one active workbook, alive across tool calls — that
persistence is what makes it memory rather than a calculator.

Reads and writes go wherever the server process can reach, which is normal for a
local stdio server and the same posture as the official filesystem server. Run it
as a user with only the access you intend it to have.

## No network

No sockets, no ports, no telemetry. Your host spawns this as a child process and
they exchange newline-delimited JSON-RPC over stdin and stdout; the engine is
WASM in that same process, so a formula is a function call rather than a request.
An air-gapped machine is a supported way to run this. Checked rather than
asserted: after a full session — create a block, attach a field rule, evaluate a
formula, save an `.xlsx` — the process holds six pipes and no sockets, on no
listening port.

## Development

A thin shell over three LogiSheets packages:
[`logisheets-runtime`](https://www.npmjs.com/package/logisheets-runtime) (the
headless engine), `logisheets-logician` (the tool definitions), and the
Rust/WASM core.

```bash
npm install && npm test
```

To work on the engine at the same time, check out
[LogiSheets](https://github.com/logisky/LogiSheets) as a sibling directory, build
its packages, and run `npm run link:local` — that symlinks the three into
`node_modules` so local engine changes take effect without reinstalling. Re-run
it after any `npm install`.

To use it as a library, `createServer` returns the MCP `Server`, the
`WorkbookSession` and the tool map, so you can host it over any transport:

```ts
import {createServer} from 'logisheets-mcp'
const {server, session, tools} = createServer({mode: 'full'})
```

## License

MIT. Part of the [LogiSheets](https://github.com/logisky/LogiSheets) project.
