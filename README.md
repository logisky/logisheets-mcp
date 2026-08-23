# logisheets-mcp

**A real spreadsheet engine your AI agent can think in.**

An [MCP](https://modelcontextprotocol.io) server that gives any LLM agent a
real, Excel-compatible calculation engine — with structured memory it can
address semantically, and a genuine `.xlsx` at the end that a human can open,
audit, and keep using.

Built on [LogiSheets](https://github.com/logisky/LogiSheets), a spreadsheet
engine written in Rust. MIT licensed, self-hostable, no cloud dependency.

## Why

Agents are doing real work that is spreadsheet-shaped — financial models, data
reconciliation, analysis — and they are bad at exactly the parts a spreadsheet
engine is good at.

**Arithmetic.** Agents mis-sum and mis-multiply. Here they don't have to: they
write a formula and a deterministic engine evaluates it.

**Memory.** Across a thirty-step task, intermediate state has to live
*somewhere* structured. A context window is lossy and expensive; a code
sandbox's variables vanish. This server gives the agent an external structured
disk it reads and writes across the whole task.

**Addressing.** Agents are bad at spatial reasoning, so a raw grid is a fragile
surface — they lose track of where things are, and their own edits break their
references. So the agent doesn't address `C7`. It addresses
**`(block, row_key, field)`**:

> set the `price` field of the `2025` record in the `revenue` block

Insert a row, move the block, add a column — that address still resolves. This
is the whole point: **memory that survives the agent's own edits.**

### vs. a Python sandbox

A code interpreter can compute, but you get a throwaway script result. Here you
get a real `.xlsx` with **live formulas still in it** — open it in Excel, change
an input, and the model recalculates. It round-trips the human's existing files,
and it runs on your machine, which matters when the data can't leave.

## Measured against the alternatives

The tasks were written down and committed *before* any other server was looked
at ([`bench/TASKS.md`](bench/TASKS.md)), every expected value is derived
independently in Python, and the harness is in [`bench/`](bench/) so you can
re-run this and disagree with it. Two other servers that work on a local
`.xlsx`: [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) 0.11.1,
which has its own Rust recalc engine, and
[excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) 0.1.8, the
most-installed one, built on openpyxl. Google Sheets connectors solve a
different problem and are not here.

| | logisheets-mcp | spreadsheet-kit | excel-mcp-server |
| --- | --- | --- | --- |
| Write a formula, read its value | **30** | **30** | `"=SUM(A1:A2)"` |
| Five-year DCF, value per share | **20.803603** · 15 calls | **20.803603** · 6 calls | formula text |
| 4×4 sensitivity, 16 answers | **all 16** · 1 call, 950 B | **all 16** · 16 calls, 1245 B | can't |
| Solve backwards for an input | **0.080699** · 1 call, 202 B | **0.080699** · 18 calls, 1399 B | can't |
| Reopen it later and explain it | 4 calls, **2.4 kB** | 5 calls, 21 kB | 2 calls, 24 kB |
| Answer again after the shape changed | **19.383943** | `#VALUE!` | formula text |

Three things in that table are worth spelling out.

**openpyxl cannot evaluate.** `apply_formula` then `read_data_from_excel`
returns the string `"=SUM(A1:A2)"`, not `30`. The file it writes is a correct
model — reopen it in an engine and the DCF computes to 20.803603 — but nothing
in the session can read an answer out of it, so no sensitivity table and no
inverse solve are possible at all. This is structural, not a bug.

**Building costs us more; asking costs us less.** The DCF took 15 calls against
6, because those calls declare a schema rather than write cells. The bet is that
you ask more often than you build, and asking is where it comes back: one call
for a sixteen-cell sensitivity grid, one call to solve backwards. The other
engine is entirely correct on both — it just has to be driven a scenario at a
time.

**Reading a model back is where the schema pays.** Same question — how is free
cash flow computed? — to a file each server saved itself:

```
logisheets-mcp    #FIELD("rev") * BLOCKREF("assum","margin","v")
                                * (1 - BLOCKREF("assum","tax","v"))
spreadsheet-kit   B11*$B$3*(1-$B$4)
excel-mcp-server  =B11*$B$3*(1-$B$4)
```

All three are correct. Only the first is an *answer*: the others send the agent
to read A3 to find out what `$B$3` is, and again next session, because the file
records no meaning. That is the 2.4 kB against 21 kB in the table — a tenth of
the context for a better answer.

Then the shape changes. Delete a projected year, insert two rows at the top and
a column at the left — through each server's own row and column tools — and ask
for the value per share again. `SUM(E11:E15)` and the terminal-value references
to `C15` no longer point where they used to, and two of the three models stop
computing. `SUM(BLOCKREFS("proj","*","pv"))` never referred to a position, so
there is nothing to fix; the answer is read back by name.

### Where they are better

Not a clean sweep, and the parts that are not deserve saying.

spreadsheet-kit is a genuine peer, correct on every task it can attempt, and has
things this server does not: forks with undo, branching and checkpoints,
LibreOffice-backed screenshots, and region detection that picked the assumptions
block out of an unlabelled sheet at 0.70 confidence unprompted. Its 53 tools
cover ground our 20 do not. excel-mcp-server has charts, pivot tables and cell
formatting, none of which are here.

And a caveat on the reading task: each server was reading back a file *it* had
written, so ours had blocks in it because we put them there. On a plain
spreadsheet from a person, `convert_to_block` adopts the table first — it reads
the field names off the header row and works out the key column — after which
`BLOCKREFS` by name works. But the individual cell formulas stay in
coordinates, so that half of the advantage does not transfer, and the honest
version of the comparison says so.

### The bench found our bugs first

Its first run had us returning `0` instead of `30`: the seed workbook was
written by openpyxl, our reader panicked on an absolute relationship target and
then on a `docProps` without the children it assumed, and the session carried on
against an empty workbook. Building the rest of it turned up more — every
inline-string label in a file dropped on the way in, so a header row read as
empty; adopting a table freezing the model it adopted into static numbers while
reporting success.

The failure mode was the same every time: a wrong answer, reported as a success,
never a crash. Fuzzing had not found any of them. That is the argument for
testing this way rather than with a feature matrix.

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

## License

MIT. Part of the [LogiSheets](https://github.com/logisky/LogiSheets) project.
