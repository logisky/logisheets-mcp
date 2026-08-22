# Comparison tasks

Written before any contestant's tool list was read, and not edited afterwards —
see the git history of this file. A benchmark authored by one of the entrants is
worth only as much as its ability to be checked, so:

* every expected value is derived independently (closed form or Python), never
  from any server's own output;
* tool-call counts are recorded, because context is the scarce resource;
* tasks we expect to lose are in the list on purpose, and their results are
  reported the same way as the rest;
* the harness and the raw transcripts are committed, so anyone can re-run this
  and disagree with the conclusion.

The field is scoped to servers that work on a local `.xlsx`. Google Sheets
connectors solve a different problem and are not included.

## T1 — Does a formula written in this session return a value?

Write `=SUM(...)` over cells written in the same session, then read the cell.
Splits engines that evaluate from libraries that only store formula text.

Expected: the sum. Fail = the formula string, a stale cached value, null, or an
error.

## T2 — A DCF, end to end

Assumptions (base revenue 1000, growth 8%, margin 20%, tax 25%, WACC 10%,
terminal growth 2.5%, net debt 500, shares 100), five projected years, terminal
value by Gordon growth, discount, subtract net debt, divide by shares.

Expected: per-share value, computed in Python from the same inputs.
Record: tool calls used.

## T3 — Sensitivity grid

Value per share across WACC × terminal growth, 4×4 = 16 cells.

Expected: 16 values from the same Python model.
Record: tool calls used — this is where a per-cell API gets expensive.

## T4 — Solve backwards

Which WACC makes value per share equal 30?

Expected: bisection in Python to 1e-4. A server with no search facility may
answer by iterating itself; that is a pass, and the tool-call count is the
interesting number.

## T5 — Hand-off after a human edit

Build the T2 model, then apply edits a person would make in Excel: insert two
rows at the top, insert a column at the left, rename a column header, delete one
projected year, move the assumptions block. Then answer a question that requires
reading the model back.

Expected: an answer consistent with the edited model, recomputed in Python.
This is the task blocks exist for; a server addressing cells by A1 has to
re-derive every address.

## T6 — Round-trip fidelity

Open a real Excel-produced `.xlsx` carrying features no agent tool models
(conditional formatting, a table, page setup), write one cell, save, and compare
every part that was not touched.

Expected: untouched parts byte-identical, or at least semantically preserved.

## T7 — Formula correctness corpus

A fixed list of formulas with known answers: INDEX/MATCH, COUNTIFS, date
arithmetic, `TEXT()` with a number format, operator precedence.

Expected: the documented Excel result for each.

## T8 — Is the saved file really a spreadsheet?

Reopen the saved file in an independent engine and check the formulas are live,
not just the numbers.

## T9 — Network behaviour

Observe the server process's sockets during a full session.

Expected: recorded, not scored. A server that uploads the workbook to a hosted
API is a different privacy proposition, and that is a fact a reader should have.

## T10 — Tasks we expect to lose

Create a chart. Apply cell formatting. Read a 200k-row sheet. Undo a mistake.
Screenshot a sheet. Reported in the same table as everything else.
