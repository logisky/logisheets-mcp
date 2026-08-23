# logisheets-mcp

**A real spreadsheet engine your agent can think in.** Excel-compatible formulas
it doesn't have to do in its head, structured memory it addresses by name, and a
genuine `.xlsx` at the end that a person can open, audit and keep using.

An [MCP](https://modelcontextprotocol.io) server over
[LogiSheets](https://github.com/logisky/LogiSheets), a spreadsheet engine written
in Rust. MIT, runs on your machine, opens no sockets.

---

Three things an agent is bad at on a grid, and does not have to do here.

### It doesn't do the arithmetic

```
set_field_rule   proj.pv = #FIELD("fcf") * #FIELD("df")
describe_block   proj
  →  Y1 147.2727   Y2 144.5950   Y3 141.9660   Y4 139.3848   Y5 136.8506
```

The rule is stated once, **for the field** — not per cell. Five years of it
materialise, and a sixth computes the moment a row is added, with no formula
written for it. Writing the same formula N times with the row number adjusted is
precisely where a model makes a silent mistake.

### It doesn't keep track of where anything is

```
… the sheet is reshaped: a projected year deleted,
  two rows inserted at the top, a column at the left …

describe_block   val
  →  per_share 19.383943        ← still right, nothing re-derived
```

Same question, same address, after the shape changed underneath it.
`SUM(BLOCKREFS("proj","*","pv"))` never referred to a position, so the edit had
nothing to break. A model that spent its attention on bookkeeping — *did my rows
shift? is my range still right?* — spends none of it here.

### It doesn't burn context on round trips

```
preview_changes  scenarios: WACC × terminal growth, 4 × 4
  →  16 answers, one call, 950 bytes, nothing written to the workbook

goal_seek        per_share = 30, by varying WACC
  →  0.080699, one call
```

A whole sensitivity table and an inverse solve are single questions, asked on a
temp branch that is discarded. No loop of write-recalculate-read, and no risk of
leaving a scenario behind in the model.

---

**And the file at the end is a real spreadsheet.** Live formulas, not baked
numbers — open it in Excel, change an input, watch it recompute. Formulas can be
written out as `BLOCKREF("proj","Y3","pv")` for a person to read, or resolved to
plain coordinates for Excel to chew on.
## Install

Requires Node 20+.

```bash
npm install -g logisheets-mcp
```

### Claude Desktop

Add to `claude_desktop_config.json`:

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

On macOS that file lives at
`~/Library/Application Support/Claude/claude_desktop_config.json`; on Windows,
`%APPDATA%\Claude\claude_desktop_config.json`. Restart Claude Desktop
afterwards.

### Cursor / Cline / other hosts

Any MCP host that can spawn a stdio server works — point it at the
`logisheets-mcp` command. For Cursor, add the same block to
`~/.cursor/mcp.json`.

## Try it

> Build me a three-year revenue model: 100 units at $9.50 growing 40% a year,
> with a 30% cost of goods. Then save it to ~/model.xlsx.

The agent creates a block, fills it, writes the formulas, and hands back a file.
The numbers are the engine's, not the model's guesses — and the `.xlsx` has real
formulas in it, so you can change an assumption in Excel and watch it recompute.

## See it work

```bash
npm run build && npm run demo
```

Builds a small revenue model over real MCP-on-stdio against `dist/cli.js` — the
same code path Claude Desktop drives — and checks every claim as it goes: totals
the engine computed, a rule that reaches rows added later, blocks that keep
resolving after the model grows underneath them, and a real `.xlsx` whose
formulas are verified by reading the file's own bytes. No LLM is involved; the
engine is the subject, and hard-coding the calls is what makes the guarantees
checkable rather than a story about a chat session.

## The agent loop

```
list_blocks                     orient: what do I have?
create_block                    open a structured workspace
add_block_rows / set_block_cells    fill it, addressed by (block, key, field)
eval_formula / a stored formula      the engine does the math
describe_block                  read structured results back
save_workbook                   hand the human a real .xlsx
```

## Tools

The default surface is deliberately small — 20 tools. Tool-selection accuracy
falls as the list grows, and every description costs context on every turn.

| Tool | What it does |
| --- | --- |
| `open_workbook` | Start a fresh workbook, or load an existing `.xlsx` from disk. Optional — one appears on first use. |
| `save_workbook` | Write to a real `.xlsx` file. This is how work gets handed back. |
| `export_xlsx` | The file as base64, for hosts with no shared filesystem. |
| `list_blocks` | Every sheet and block, plus where the next block should go. |
| `describe_block` | A block's schema, keys, and (optionally) its current values. |
| `eval_formula` | Evaluate an Excel formula and return the value. Nothing is stored. |
| `create_block` | Create a named, structured table. First field is the row key. |
| `convert_to_block` | Adopt a table that is already in ordinary cells, in place. Reads the field names off the header row and works out which column is the row key. |
| `add_block_rows` | Add records — at the end, or `after_key` / `before_key` to place them. |
| `delete_block_rows` | Remove records. |
| `move_block_row` | Reorder rows, by key. Presentation only: no computed value changes. |
| `set_block_cells` | Write cells by `(block, row_key, field)`. Batched, atomic. |
| `set_field_rule` | Give a field a formula, a validation rule, or an editability rule. |
| `list_violations` | Which cells break their field's validation rule, and why. |
| `preview_changes` | What edits *would* do, without doing them. One hypothetical, or a whole grid of scenarios in a single call. |
| `trace` | What a cell reads, and what reads it — from the engine's dependency graph. |
| `goal_seek` | What input makes a chosen output equal a target. Searches inside the engine; changes nothing. |
| `create_sheet` | Add a sheet. |
| `get_cells` / `set_cells` | Raw-cell escape hatch for data with no structure. |

Formulas are Excel-compatible, plus `BLOCKREF(block, key, field)` for reading a
block cell semantically. Inside a field rule, `#FIELD("name")` is the same row's
sibling and `#FIELD("name", "key")` is another row of the same block — the one
carrying that key, never a positional offset, so reordering rows cannot change
what a formula means.

### Analysing a model, not just building one

`preview_changes` takes a list of `scenarios` and an optional `watch`, which is
what turns exploration from dozens of round trips into one:

```jsonc
{
  "scenarios": [
    {"label": "wacc 9%",  "changes": [{"block":"assum","row_key":"wacc","field":"v","value":0.09}]},
    {"label": "wacc 12%", "changes": [{"block":"assum","row_key":"wacc","field":"v","value":0.12}]}
  ],
  "watch": [{"block":"valuation","row_key":"per_share","field":"v"}]
}
```

Each scenario runs on its own temp branch and is discarded, so the live model is
never touched — no mutate-and-revert, and nothing left behind if a scan fails
half way. A 4×4 sensitivity grid is one call returning sixteen numbers.

`goal_seek` runs the same trick backwards — "what discount rate gives a value per
share of 30" — with the search inside the engine rather than as a conversation,
so it is one call instead of one per bisection step. It says when a target is
simply not reachable in the bracket instead of returning the nearest number it
happened to stop on.

`trace` answers the two audit questions from the engine's dependency graph:
what a cell reads, and what reads it. The second one is why it exists — formula
text can be read forwards but not backwards, and "what breaks if I change this"
is the question you want before touching an assumption.

Reading a model is semantic too: `describe_block` returns each field's rule, so
an agent learns the model's logic without visiting a cell, and formulas come back
naming what they read (`B24 / BLOCKREF("assum","shares","v")`) rather than as
coordinate chains you have to chase.

### The full surface

Set `LOGISHEETS_MCP_TOOLS=full` for 50 tools: undo/redo, cell formatting,
merges, comments, checkpoints, block move/resize, cross-block links, and raw
row/column structure.

```json
{
    "mcpServers": {
        "logisheets": {
            "command": "npx",
            "args": ["-y", "logisheets-mcp"],
            "env": {"LOGISHEETS_MCP_TOOLS": "full"}
        }
    }
}
```

Mutating tools are marked with MCP's `readOnlyHint` / `destructiveHint`
annotations, so a host can gate them behind user approval.

## Blocks, briefly

A **block** is a named, structured region of a sheet — a table with a schema.

- The first field is the **row key**: the stable name of each record.
- Fields can carry a **value formula** (engine-computed, so the agent can't
  write a stale number into it), a **validation** rule, or an **editability**
  rule.
- Everything is addressed by name. Row and column indices never enter the
  agent's reasoning.

Because blocks are created *by the agent as it works*, this needs no
pre-prepared file — you can point it at a blank workbook or at a spreadsheet
someone sent you.

## Getting the file back

`save_workbook` writes a real `.xlsx` and its result carries an MCP
**resource link** — a uri, media type and size — not the file. The workbook is
also listed as a resource (`workbook://current.xlsx`), so a host that wants the
bytes reads them with `resources/read` and hands the human a download.

That split is the point: a tool result goes into the model's context, where a
200 KB workbook would cost roughly 280 KB of text and teach the model nothing.
The link costs a line. `export_xlsx` still returns base64 for hosts that
implement no resources at all, but it is the fallback, not the mechanism.

Reads go through the same serialization lane as tool calls, so a host fetching
the file can never catch a half-applied transaction.

`open_workbook` and `save_workbook` read and write wherever the server process
can — normal for a local stdio server, and the same posture as the official
filesystem server. Both are marked as mutating so a host can prompt before they
run; if you need tighter limits, run the server as a user with only the access
you intend it to have.

## State model

One MCP session holds one active workbook, alive across tool calls — that
persistence is what makes it memory rather than a calculator. `open_workbook`
replaces it. Multiple named workbooks per session may come later.

## No network

The server opens no sockets and listens on no ports. "stdio transport" is
literal: your MCP host spawns this as a child process and they exchange
newline-delimited JSON-RPC over its stdin and stdout — the same pipes any
command-line program gets. The engine is WASM running in that same process,
so a formula is a function call, not a request.

Checked rather than asserted. After a full session — create a block, attach a
field rule, evaluate a formula, save an `.xlsx` — the process holds:

```
fd types: {CHR: 2, DIR: 4, KQUEUE: 3, PIPE: 6, REG: 13}
network files (lsof -a -i):     0
unix sockets  (lsof -a -U):     0
listening ports:                0
```

Six pipes, no sockets. Nothing is uploaded, no telemetry is collected, and an
air-gapped machine is a supported way to run this. The only things it touches
outside its own memory are the files you name — see the filesystem note under
[Getting the file back](#getting-the-file-back).

That is the `logisheets-mcp` binary, which is what an MCP host runs. Using it
[as a library](#use-as-a-library) you can attach any transport you like,
including an HTTP one — but then the socket is yours, opened deliberately.

## Benchmarks

The claims above are measured, not asserted. The harness is in
[`bench/`](bench/): tasks written down and committed *before* any other server
was looked at ([`bench/TASKS.md`](bench/TASKS.md)), every expected value derived
independently in Python, so you can re-run it and disagree.

Against the two other MCP servers that work on a local `.xlsx` —
[spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) 0.11.1, which has
its own Rust recalc engine, and
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
| Keep a handed-over file's features | 6 of 8 | **8 of 8** | **8 of 8** |

The last row is ours to fix: an open-write-save keeps conditional formatting,
page setup, frozen panes, data validation, merged cells and column widths, and
still drops an Excel table (ListObject) and a defined name.

`"=SUM(A1:A2)"` in the third column is not a bug — openpyxl stores formulas
without evaluating them, so no scenario can be read back and no inverse solve is
possible. It writes a correct model; it just cannot answer a question about one.

spreadsheet-kit is a genuine peer, correct on everything it can attempt, and
needed fewer calls than we did to build the model — our extra calls declare a
schema rather than write cells, which is the trade that pays off in the rows
below it. It also has forks with undo, branching and checkpoints, and
LibreOffice-backed screenshots, none of which are here.

One caveat on the reading row: each server was reading back a file *it* wrote,
so ours had blocks in it because we put them there. Given a plain spreadsheet
from a person, `convert_to_block` adopts the table first — reading the field
names off the header row and working out the key column — and then `BLOCKREFS`
by name works. The individual cell formulas stay in coordinates, so only half of
that advantage transfers.

Building this turned up defect after defect in our own engine before it said
anything about anyone else's: a workbook openpyxl wrote failing to load, every
inline-string label dropped on the way in, adoption freezing the model it
adopted into static numbers. Every one presented as a wrong answer reported as a
success — never as a crash — and fuzzing had found none of them.

## Use as a library

```ts
import {createServer} from 'logisheets-mcp'
import {StreamableHTTPServerTransport} from '@modelcontextprotocol/sdk/server/streamableHttp.js'

const {server, session} = createServer({mode: 'full'})
await server.connect(new StreamableHTTPServerTransport(/* … */))
```

`createServer` returns the MCP `Server`, the `WorkbookSession`, and the tool map,
so you can host it over any transport or embed it in an agent framework.

## Development

The server is a thin shell over three LogiSheets packages:
[`logisheets-runtime`](https://www.npmjs.com/package/logisheets-runtime) (the
headless engine), `logisheets-logician` (the agent tool definitions), and the
Rust/WASM core. Working on the server alone needs nothing special:

```bash
git clone https://github.com/logisky/logisheets-mcp.git
cd logisheets-mcp
npm install
npm test
```

Working on the engine at the same time is the other mode. Check out
[LogiSheets](https://github.com/logisky/LogiSheets) as a sibling directory,
build its packages, then:

```bash
npm run link:local     # re-run after any npm install
```

That symlinks the three packages into `node_modules` so local engine changes
take effect without reinstalling. `scripts/release-deps.mjs` puts the registry
ranges back before publishing.

### Releasing

A tag does it. `.github/workflows/publish.yaml` runs the tests, publishes to
npm with provenance, and registers the new version with the MCP Registry:

```bash
npm version 0.2.0        # bumps both files, commits, tags v0.2.0
npm run check-release    # optional; CI runs it too
git push --follow-tags
```

The workflow can also be run by hand from the Actions tab, which takes the
version from `package.json` instead of a tag. The npm step skips a version that
is already published, so a run that failed at the registry step can just be
re-run — the two publishes are not a transaction.

`npm version` also rewrites `server.json`, via the `version` lifecycle script.
The registry keeps the version in two places — the server's own `version` and
the version of the npm package it points at — and hand-editing them is the step
most likely to be missed.

`check-release` is the gate. Four things have to agree: the tag, `package.json`,
and both `server.json` version fields. `mcpName` also has to equal
`server.json`'s `name`, because the registry proves ownership by reading
`mcpName` out of the *published* npm package. `npm publish` cannot be undone —
a version number is spent the moment it lands — so the workflow runs this check
before publishing, not after.

Registry auth needs no secret: the workflow authenticates with GitHub OIDC,
which is what grants the `io.github.logisky/` namespace. The one secret is
`NPM_TOKEN`.

## License

MIT. Part of the [LogiSheets](https://github.com/logisky/LogiSheets) project.
