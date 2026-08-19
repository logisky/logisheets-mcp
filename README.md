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

The default surface is deliberately small — 14 tools. Tool-selection accuracy
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
| `add_block_rows` | Append records. |
| `delete_block_rows` | Remove records. |
| `set_block_cells` | Write cells by `(block, row_key, field)`. Batched, atomic. |
| `set_field_rule` | Give a field a formula, a validation rule, or an editability rule. |
| `create_sheet` | Add a sheet. |
| `get_cells` / `set_cells` | Raw-cell escape hatch for data with no structure. |

Formulas are Excel-compatible, plus `BLOCKREF(block, key, field)` for reading a
block cell semantically.

### The full surface

Set `LOGISHEETS_MCP_TOOLS=full` for ~48 tools: undo/redo, cell formatting,
merges, comments, checkpoints, block move/resize, cross-block links,
validation-violation inspection, and raw row/column structure.

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
Rust/WASM core. Working on it usually means working on those too, so
dependencies point at a sibling checkout:

```bash
git clone https://github.com/logisky/LogiSheets.git
git clone https://github.com/logisky/logisheets-mcp.git
cd logisheets-mcp
npm install
npm run link:local     # re-run after any npm install
npm test
```

`npm run link:local` symlinks the three packages into `node_modules`, so local
engine changes take effect without reinstalling. `scripts/release-deps.mjs`
flips them back to registry ranges before publishing.

## State model

One MCP session holds one active workbook, alive across tool calls — that
persistence is what makes it memory rather than a calculator. `open_workbook`
replaces it. Multiple named workbooks per session may come later.

## License

MIT. Part of the [LogiSheets](https://github.com/logisky/LogiSheets) project.
