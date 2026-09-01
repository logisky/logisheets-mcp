# The second session

Most of what is written about agents and spreadsheets is about the first
session: can the model build the model. That question is close to settled — a
capable model with any file-writing tool produces a plausible five-year
projection.

The interesting question is the second session. The conversation that built the
spreadsheet is almost never the conversation that has to answer a question about
it. A week later, someone opens a new chat and says *why is Q3 down* or *rerun
this at 12% instead*. The agent that answers has none of the context that built
the file. All it has is the file.

What the file records is what the next session can know. That is the whole
argument for blocks.

## What a grid hands over

Here is a cell from a discounted cash flow model, as a fresh agent finds it:

```
C11  =B11*$B$3*(1-$B$4)
```

Correct. Also meaningless. To learn what it computes, the agent must fetch the
label column, discover that `B3` sits next to the word "margin" and `B4` next to
"tax", and infer that the labels describe the cells to their right — a
convention, not a fact, and one that plenty of real spreadsheets break.

Three costs follow, and the third is the one that hurts:

1. **Extra round trips.** Every formula worth understanding needs a second read
   to resolve, and often a third to check the resolution.
2. **A guess in the middle.** Nothing in an `.xlsx` says that `A3` labels `B3`.
   The agent infers it, silently, and is occasionally wrong.
3. **It is not saved.** The agent works out what the model means and then the
   session ends. The next session pays the same bill, makes the same guess, and
   has no way to know whether the previous one got it right.

The file grew data. It did not grow understanding.

## A block is a table that knows its own name

A **block** is a named table on the sheet. Rows have keys, columns have names,
and everything is addressed as `(block, row_key, field)` rather than by
position.

```
create_block     proj    fields: year, units, price, revenue, profit
set_field_rule   proj.revenue = units × price            ← once, for the column
set_field_rule   proj.profit  = revenue × assum.margin × (1 − assum.tax)
add_block_rows   y1 … y5
```

Two properties matter for the first session — a field's formula is stated once
for the whole column, so there is no twentieth formula to get the row number
wrong in; and a reference survives reshaping, because *the `profit` field of the
row keyed `y3`* does not stop being true when someone inserts two rows at the
top.

But the property that matters for the second session is quieter: **all of that
is schema, and schema is written into the file.** The names, the key column,
which fields the engine computes and which a human types, and the rule for each
computed field — those are not conventions recovered by inference. They are
stored, and they come back.

## What the second session sees

Same model, reopened by an agent that has never seen it. One call:

```
list_blocks
  →  Model
       assum   fields: k, v                                  key: k    2 rows
       proj    fields: year, units, price, revenue, profit    key: year 5 rows
               computed: revenue, profit
       next_block_start: row 13
```

The workbook has introduced itself. Which tables exist, what every column is
called, which column is the row key, which columns the engine owns — so writing
to them will be refused, and knowing that in advance saves a failed call — and
where there is room to write something new.

A second call asks what the computed columns actually do:

```
describe_block proj
  →  revenue = #FIELD("units") * #FIELD("price")
     profit  = #FIELD("revenue") * BLOCKREF("assum","margin","v")
                                 * (1 - BLOCKREF("assum","tax","v"))
```

Compare that with `=B11*$B$3*(1-$B$4)`. It is the same arithmetic. The
difference is that this version is an *explanation* — it can be read aloud to a
person — and it needs no second lookup, because there are no coordinates in it
to resolve.

Two calls in, the agent knows the model. It can now write

```
=BLOCKREF("proj","y3","profit")
```

on the first try, from orientation alone, and get a number back from a real
engine rather than a guess.

## Orienting costs the same on a big model as on a small one

The cheap part is not an accident of the example being small. Reading a schema
is `O(columns)`; reading a grid to understand it is `O(cells)`.

Measured on the wire — the actual text a host puts in front of the model:

| | `list_blocks` |
| --- | --- |
| 5-row projection | 1 call, **540 bytes** |
| 105-row projection | 1 call, **545 bytes** |

Twenty-one times the data, five more bytes — the extra digits in the row count.
This is the shape of the claim: what an agent needs in order to *find its way
around* a workbook does not grow with what is in it. Asking for the data still
costs what the data costs; `describe_block include_rows: true` on the 105-row
model is about twenty times the size of the schema read, exactly as it should
be. The point is that the second session gets to decide, instead of having to
read everything before it knows what any of it is.

## The question formula text cannot answer

Before changing an assumption in a model you did not write, the thing you want
to know is *what reads this*. A formula tells you its inputs; nothing in a file
tells you its consumers. The dependency graph does:

```
trace  assum.margin  direction: dependents
  →  proj.profit          (field granularity — approximate: true)
```

Block dependencies are tracked per field rather than per row, so this is an
over-approximation and says so. An honest "these three fields, possibly fewer
rows than that" is a usable answer. A confident "nothing depends on this" would
be the most dangerous possible wrong one.

## Files that arrive without a schema

Most spreadsheets in the world are grids, and they are not going to stop being
grids. `convert_to_block` adopts a table that is already in ordinary cells — in
place, reading the field names off the header row and working out which column
holds the keys. Nothing moves; the same cells acquire a schema.

That is a one-time cost paid by whichever session touches the file first. Every
session after it gets the two-call path above, including the ones that happen
months later in a different tool.

## Tested, not asserted

The claim in this article is a test in the repository, not a paragraph.
[`src/cold-read.test.ts`](../src/cold-read.test.ts) builds a model in one
session, saves it, and opens the file in a second session that shares nothing
with the first — its own server, its own workbook, no memory of what the first
one did. It then checks:

- one `list_blocks` call recovers block names, field names, the key field, the
  computed fields and the row counts;
- every field rule that comes back contains `#FIELD` or `BLOCKREF` and **no A1
  coordinate at all** — asserted with a regex, so a regression that starts
  writing coordinates into rules fails the suite;
- the fresh session writes a `BLOCKREF` formula from orientation alone, and the
  engine's answer matches arithmetic done independently in TypeScript;
- `trace` finds the dependent field, and admits its own precision;
- orientation costs one call and does not grow with the row count.

The comparison against other MCP servers that work on a local `.xlsx` is in
[`bench/`](../bench), including the reopen-and-explain task: 4 calls and 2.4 kB
here, against 21 kB and 24 kB for the two contestants — with the caveat, stated
there too, that each server was reading back a file it had written itself, so
only part of that margin transfers to a spreadsheet that came from a person.

## Try it

```bash
npm install -g logisheets-mcp
```

Then point any MCP host at it — `npx -y logisheets-mcp` over stdio — and ask for
a model. It runs on your machine, opens no sockets, and hands back a real
`.xlsx` a person can open in Excel.

The part worth trying is the second half: close the chat, start a new one, and
ask the file to explain itself.
