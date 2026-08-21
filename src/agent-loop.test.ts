/**
 * The claim this server makes, tested: an agent can build structured memory,
 * offload the arithmetic to a real engine, read the results back, and hand a
 * human a real .xlsx.
 *
 * These tests drive the tools the way an agent does — by MCP name, with JSON
 * arguments — so they exercise the same path a host does, minus the transport.
 */

import {mkdtemp, rm, readFile, copyFile} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {inflateRawSync} from 'node:zlib'
import {fileURLToPath} from 'node:url'
import {join, resolve, dirname} from 'node:path'
import {afterEach, beforeEach, describe, expect, it} from 'vitest'
import type {Tool, ToolContext} from 'logisheets-logician'
import {createServer} from './server.js'
import {WorkbookSession} from './session.js'

/** Call a tool by its MCP name, as the CallTool handler would. */
function makeCaller(session: WorkbookSession, tools: Map<string, Tool>) {
    const ctx: ToolContext = {
        workbook: session.client,
        signal: new AbortController().signal,
        confirm: async () => true,
        log: () => {},
    }
    return async function call<T = unknown>(
        name: string,
        args: Record<string, unknown> = {}
    ): Promise<T> {
        const tool = tools.get(name)
        if (tool === undefined) throw new Error(`no such tool: ${name}`)
        // ctx.workbook is read per call: `open_workbook` swaps the active
        // workbook, so a context captured once must not pin the old client.
        const result = await tool.handler(args, {
            ...ctx,
            workbook: session.client,
        })
        return result.data as T
    }
}

/** Every worksheet part of an .xlsx, concatenated. An .xlsx is a zip; reading
 *  the bytes is the only way to check what another spreadsheet would see. */
function worksheetXml(buf: Buffer): string {
    let i = 0
    let out = ''
    while (i + 30 <= buf.length && buf.readUInt32LE(i) === 0x04034b50) {
        const method = buf.readUInt16LE(i + 8)
        const size = buf.readUInt32LE(i + 18)
        const nameLen = buf.readUInt16LE(i + 26)
        const extraLen = buf.readUInt16LE(i + 28)
        const name = buf.subarray(i + 30, i + 30 + nameLen).toString('utf8')
        const start = i + 30 + nameLen + extraLen
        const data = buf.subarray(start, start + size)
        if (name.startsWith('xl/worksheets/sheet')) {
            out += (method === 0 ? data : inflateRawSync(data)).toString('utf8')
        }
        i = start + size
    }
    return out
}

describe('logisheets-mcp agent loop', () => {
    let session: WorkbookSession
    let call: ReturnType<typeof makeCaller>
    let dir: string

    beforeEach(async () => {
        const created = createServer({mode: 'core'})
        session = created.session
        call = makeCaller(session, created.tools)
        dir = await mkdtemp(join(tmpdir(), 'logisheets-mcp-'))
    })

    afterEach(async () => {
        session.close()
        await rm(dir, {recursive: true, force: true})
    })

    it('exposes a small core surface with clean, unique names', () => {
        const {tools} = createServer({mode: 'core'})
        const names = [...tools.keys()]
        expect(names).toHaveLength(17)
        expect(new Set(names).size).toBe(names.length)
        // No namespace prefixes leaked into the model-facing names.
        expect(names.filter((n) => n.includes('__'))).toEqual([])
        // The loop from the design doc is all present.
        for (const n of [
            'open_workbook',
            'list_blocks',
            'create_block',
            'add_block_rows',
            'set_block_cells',
            'eval_formula',
            'describe_block',
            'save_workbook',
        ]) {
            expect(names).toContain(n)
        }
    })

    it('opts into the full surface without losing the core', () => {
        const core = createServer({mode: 'core'}).tools
        const full = createServer({mode: 'full'}).tools
        expect(full.size).toBeGreaterThan(core.size)
        for (const name of core.keys()) expect(full.has(name)).toBe(true)
        // Browser-only tools stay out of `full` too.
        expect(full.has('register_radio_group')).toBe(false)
        expect(full.has('get_active_selection')).toBe(false)
        // …but the genuinely useful extras are in.
        expect(full.has('undo')).toBe(true)
        expect(full.has('list_violations')).toBe(true)
    })

    it('does not hallucinate arithmetic: the engine computes it', async () => {
        // 8 * 7.5 + 100 — the kind of thing an agent gets wrong on its own.
        const v = await call<{type: string; value: unknown}>('eval_formula', {
            expr: '=8*7.5+100',
        })
        expect(v).toEqual({type: 'number', value: 160})

        // Precedence, which is where hand-arithmetic really breaks down.
        expect(await call('eval_formula', {expr: '10/2*5'})).toEqual({
            type: 'number',
            value: 25,
        })
    })

    it('builds structured memory and addresses it by (block, key, field)', async () => {
        await call('create_block', {
            sheet: 'Model',
            name: 'revenue',
            position: {row: 0, col: 0},
            fields: [
                {name: 'year'},
                {name: 'units', field_type: 'number'},
                {name: 'price', field_type: 'number'},
            ],
            initial_rows: [{key: '2024', values: {units: 100, price: 9.5}}],
        })

        await call('add_block_rows', {
            block: 'revenue',
            rows: [
                {key: '2025', values: {units: 140, price: 10}},
                {key: '2026', values: {units: 180, price: 11}},
            ],
        })

        // A formula written by semantic address — no A1 anywhere.
        await call('set_block_cells', {
            changes: [
                {
                    block: 'revenue',
                    row_key: '2025',
                    field: 'price',
                    value: '=10.5',
                },
            ],
        })

        const described = await call<{
            keys: string[]
            fields: Array<{name: string}>
            rows?: Array<{key: string; values: Record<string, unknown>}>
        }>('describe_block', {name: 'revenue', include_rows: true})

        expect(described.keys).toEqual(['2024', '2025', '2026'])
        expect(described.fields.map((f) => f.name)).toEqual([
            'year',
            'units',
            'price',
        ])
        const byKey = new Map(described.rows?.map((r) => [r.key, r.values]))
        expect(byKey.get('2024')?.units).toBe(100)
        expect(byKey.get('2025')?.price).toBe(10.5)
        expect(byKey.get('2026')?.units).toBe(180)
    })

    it('keeps semantic addresses stable across an edit that shifts rows', async () => {
        // The differentiator: an agent's own edits must not break its memory.
        await call('create_block', {
            sheet: 'Model',
            name: 'costs',
            position: {row: 0, col: 0},
            fields: [{name: 'item'}, {name: 'amount', field_type: 'number'}],
            initial_rows: [{key: 'rent', values: {amount: 1000}}],
        })
        await call('add_block_rows', {
            block: 'costs',
            rows: [{key: 'salary', values: {amount: 5000}}],
        })

        // Insert raw sheet rows above the block, moving every cell down.
        // create_block auto-created 'Model' as the second sheet -> sheetIdx 1.
        const {tools} = createServer({mode: 'full', session})
        const fullCall = makeCaller(session, tools)
        await fullCall('insert_rows', {sheetIdx: 1, start: 0, count: 3})

        // The same (block, key, field) address still resolves, unchanged.
        const after = await call<{
            rows?: Array<{key: string; values: Record<string, unknown>}>
        }>('describe_block', {name: 'costs', include_rows: true})
        const byKey = new Map(after.rows?.map((r) => [r.key, r.values]))
        expect(byKey.get('rent')?.amount).toBe(1000)
        expect(byKey.get('salary')?.amount).toBe(5000)
    })

    it('orients on an empty session without an explicit open', async () => {
        // No open_workbook call: the scratchpad exists on first touch.
        const groups = await call<
            Array<{sheet_name: string; blocks: unknown[]}>
        >('list_blocks')
        expect(groups.length).toBeGreaterThan(0)
        expect(groups[0]?.blocks).toEqual([])
    })

    it('hands back a real .xlsx that reopens with its formulas alive', async () => {
        await call('create_block', {
            sheet: 'Model',
            name: 'orders',
            position: {row: 0, col: 0},
            fields: [
                {name: 'id'},
                {name: 'qty', field_type: 'number'},
                {name: 'unit', field_type: 'number'},
            ],
            initial_rows: [{key: 'A-1', values: {qty: 4, unit: 2.5}}],
        })

        // A live formula, not a precomputed constant. Raw cells are addressed
        // by zero-based (sheetIdx, row, col); Sheet1 (idx 0) has no block, so
        // this range is entirely ordinary cells.
        await call('set_cells', {
            sheetIdx: 0,
            cells: [
                {row: 0, col: 0, content: '10'},
                {row: 1, col: 0, content: '20'},
                {row: 2, col: 0, content: '=SUM(A1:A2)'},
            ],
        })

        const out = join(dir, 'model.xlsx')
        const saved = await call<{path: string; bytes: number}>(
            'save_workbook',
            {path: out}
        )
        expect(saved.path).toBe(out)
        expect(saved.bytes).toBeGreaterThan(0)

        // It is a real zip container, i.e. something Excel will open.
        const bytes = await readFile(out)
        expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04])

        // And LogiSheets can reopen what it wrote — block structure intact.
        const reopened = await call<{sheets: string[]}>('open_workbook', {
            path: out,
        })
        expect(reopened.sheets).toContain('Model')
        const blocks = await call<
            Array<{sheet_name: string; blocks: Array<{name: string}>}>
        >('list_blocks')
        const names = blocks.flatMap((g) => g.blocks.map((b) => b.name))
        expect(names).toContain('orders')

        // The formula is still a formula — not baked down to 30 — so the human
        // can change an input in Excel and watch it recalculate.
        const cells = await call<{
            cells: Array<{ref: string; value: unknown; formula?: string}>
        }>('get_cells', {
            sheetIdx: 0,
            startRow: 0,
            startCol: 0,
            endRow: 2,
            endCol: 0,
        })
        const a3 = cells.cells.find((c) => c.ref === 'A3')
        expect(a3?.value).toBe(30)
        expect(a3?.formula).toContain('SUM')
    })

    it('keeps a field rule alive across save and reload', async () => {
        // A field rule is a TEMPLATE, not a one-off write: every row, including
        // rows added later, is computed from it. So the template itself has to
        // survive the file, or the workbook silently stops being live — values
        // still sitting there, new rows coming back blank.
        //
        // It regressed exactly that way: the engine wrote the template into an
        // XML attribute without escaping it, and `#FIELD("units")` — the
        // documented syntax, always quoted — terminated the attribute at its
        // first quote. It reloaded as `#FIELD(` and computed nothing.
        const rule = '=#FIELD("units")*#FIELD("price")'
        await call('create_block', {
            sheet: 'Model',
            name: 'revenue',
            position: {row: 0, col: 0},
            fields: [
                {name: 'year'},
                {name: 'units', field_type: 'number'},
                {name: 'price', field_type: 'number'},
                {name: 'total', field_type: 'number'},
            ],
            initial_rows: [{key: '2024', values: {units: 100, price: 9.5}}],
        })
        await call('set_field_rule', {
            block: 'revenue',
            field: 'total',
            value_formula: rule,
        })

        const out = join(dir, 'rule.xlsx')
        await call('save_workbook', {path: out})
        await call('open_workbook', {path: out})

        type Described = {
            fields: Array<{name: string; value_formula?: string | null}>
            rows?: Array<{key: string; values: Record<string, unknown>}>
        }

        // The template came back intact, quotes and all.
        const after = await call<Described>('describe_block', {
            name: 'revenue',
            include_rows: true,
        })
        const total = after.fields.find((f) => f.name === 'total')
        expect(total?.value_formula).toBe('#FIELD("units")*#FIELD("price")')

        // The proof that matters: a row added AFTER the reload still computes.
        await call('add_block_rows', {
            block: 'revenue',
            rows: [{key: '2025', values: {units: 200, price: 12}}],
        })
        const grown = await call<Described>('describe_block', {
            name: 'revenue',
            include_rows: true,
        })
        const byKey = new Map(grown.rows?.map((r) => [r.key, r.values]))
        expect(byKey.get('2024')?.total).toBe(950)
        expect(byKey.get('2025')?.total).toBe(2400)
    })

    it("works on a file the user brought without destroying it", async () => {
        // The scenario the product is sold on: open the human's real .xlsx, add
        // something, hand it back. A file a user brings has NO blocks, which is
        // exactly why this went wrong — `list_blocks` derived its suggested
        // position from blocks alone, so it said row 0, and `create_block`
        // silently overwrote the first row the agent was told to build on.
        const src = resolve(
            dirname(fileURLToPath(import.meta.url)),
            '..',
            '..',
            'LogiSheets',
            'tests',
            '6.xlsx'
        )
        const work = join(dir, 'brought-by-user.xlsx')
        await copyFile(src, work)

        await call('open_workbook', {path: work})
        type Grid = {cells: Array<{ref: string; value: unknown}>}
        const before = await call<Grid>('get_cells', {
            sheetIdx: 0,
            startRow: 0,
            startCol: 0,
            endRow: 25,
            endCol: 5,
        })
        expect(before.cells.length).toBeGreaterThan(0)

        // The suggested position must clear the content, not just the blocks.
        const groups = await call<
            Array<{next_block_start: {row: number; col: number}}>
        >('list_blocks')
        const hint = groups[0]?.next_block_start
        expect(hint).toBeDefined()
        expect(hint!.row).toBeGreaterThan(0)

        // And landing on the data is refused however the agent got there.
        await expect(
            call('create_block', {
                sheet: 'Control',
                name: 'onto_data',
                position: {row: 0, col: 0},
                fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
                initial_rows: [{key: 'r1', values: {v: 1}}],
            })
        ).rejects.toThrow(/already hold data/)

        await call('create_block', {
            sheet: 'Control',
            name: 'analysis',
            position: hint,
            fields: [
                {name: 'metric'},
                {name: 'units', field_type: 'number'},
                {name: 'price', field_type: 'number'},
                {name: 'total', field_type: 'number'},
            ],
            initial_rows: [{key: 'q1', values: {units: 10, price: 2.5}}],
        })
        await call('set_field_rule', {
            block: 'analysis',
            field: 'total',
            value_formula: '=#FIELD("units")*#FIELD("price")',
        })

        await call('save_workbook', {path: work})
        await call('open_workbook', {path: work})

        // Their content came back exactly as it went in.
        const after = await call<Grid>('get_cells', {
            sheetIdx: 0,
            startRow: 0,
            startCol: 0,
            endRow: 25,
            endCol: 5,
        })
        expect(after.cells).toEqual(before.cells)

        // And the work the agent added computed and survived.
        const described = await call<{
            rows?: Array<{key: string; values: Record<string, unknown>}>
        }>('describe_block', {name: 'analysis', include_rows: true})
        expect(described.rows?.[0]?.values.total).toBe(25)
    })

    it('refuses duplicates that would break semantic addressing', async () => {
        // `(block, row_key, field)` IS the addressing scheme, so a duplicate at
        // any level makes some cell unreachable — and makes aggregates wrong
        // rather than merely odd. With two rows keyed "a", BLOCKREFS resolved
        // the first one twice and skipped the other: a block holding 1, 2 and
        // 99 summed to 4. Nothing complained.
        await call('create_block', {
            sheet: 'S',
            name: 'ledger',
            position: {row: 0, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [
                {key: 'a', values: {v: 1}},
                {key: 'b', values: {v: 2}},
            ],
        })

        // A second block under the same ref name would make every BLOCKREF to
        // it ambiguous, which surfaces as #VALUE! at evaluation time.
        await expect(
            call('create_block', {
                sheet: 'S',
                name: 'ledger',
                position: {row: 20, col: 0},
                fields: [{name: 'k'}, {name: 'v'}],
                initial_rows: [{key: 'x'}],
            })
        ).rejects.toThrow(/already exists/)

        await expect(
            call('create_block', {
                sheet: 'S',
                name: 'dup_fields',
                position: {row: 30, col: 0},
                fields: [{name: 'k'}, {name: 'v'}, {name: 'v'}],
                initial_rows: [{key: 'x'}],
            })
        ).rejects.toThrow(/field name/)

        await expect(
            call('create_block', {
                sheet: 'S',
                name: 'dup_keys',
                position: {row: 40, col: 0},
                fields: [{name: 'k'}, {name: 'v'}],
                initial_rows: [{key: 's'}, {key: 's'}],
            })
        ).rejects.toThrow(/row key/)

        // Appending a key the block already has, and repeating one in a batch.
        await expect(
            call('add_block_rows', {
                block: 'ledger',
                rows: [{key: 'a', values: {v: 99}}],
            })
        ).rejects.toThrow(/already has the row key/)
        await expect(
            call('add_block_rows', {
                block: 'ledger',
                rows: [{key: 'z'}, {key: 'z'}],
            })
        ).rejects.toThrow(/repeat the key/)

        // A genuinely new key still appends, and the total is the real total.
        await call('add_block_rows', {
            block: 'ledger',
            rows: [{key: 'c', values: {v: 3}}],
        })
        const total = await call<{value: unknown}>('eval_formula', {
            expr: '=SUM(BLOCKREFS("ledger","*","v"))',
        })
        expect(total.value).toBe(6)
    })

    it('can save block formulas as coordinates for Excel', async () => {
        // BLOCKREF is a LogiSheets function. Excel shows the saved numbers but
        // turns those cells into #NAME? the moment it recalculates, so a file
        // meant for Excel needs the coordinates instead. Both forms matter: the
        // named one is readable and reopens here intact, so neither is "the"
        // default for every purpose — the caller chooses.
        await call('create_block', {
            sheet: 'S',
            name: 't',
            position: {row: 0, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [
                {key: 'r1', values: {v: 10}},
                {key: 'r2', values: {v: 20}},
            ],
        })
        await call('set_cells', {
            sheetIdx: 0,
            cells: [
                {row: 5, col: 0, content: '=BLOCKREF("t","r2","v")'},
                {row: 5, col: 1, content: '=SUM(BLOCKREFS("t","*","v"))'},
            ],
        })

        const named = join(dir, 'named.xlsx')
        const excel = join(dir, 'excel.xlsx')
        await call('save_workbook', {path: named})
        await call('save_workbook', {path: excel, resolve_block_refs: true})

        const formulasIn = async (file: string): Promise<string[]> => {
            const xml = worksheetXml(await readFile(file))
            return [...xml.matchAll(/<f>(.*?)<\/f>/g)].map((m) =>
                (m[1] ?? '').replaceAll('&quot;', '"')
            )
        }

        // Default keeps the names.
        const kept = await formulasIn(named)
        expect(kept.some((f) => f.includes('BLOCKREF('))).toBe(true)
        expect(kept.some((f) => f.includes('BLOCKREFS('))).toBe(true)

        // Resolved leaves nothing Excel cannot evaluate: the single ref becomes
        // the cell it named, the scan becomes the range it covered.
        const resolved = await formulasIn(excel)
        expect(resolved.some((f) => f.includes('BLOCKREF'))).toBe(false)
        expect(resolved).toContain('B2')
        expect(resolved.some((f) => f.includes('B1:B2'))).toBe(true)
    })

    it('previews a change without committing it', async () => {
        // `preview_changes` is the what-if primitive: it runs the edits on the
        // engine's temp branch, reports every cell that would move, and discards
        // the branch. The discard was a silent no-op — the RPC was named
        // `cleanTempStatus` while the client interface said
        // `cleanupTempStatus`, so a client that forwards method names verbatim
        // called nothing at all and every "dry run" committed itself. A
        // sensitivity scan would have walked the model to the last probe value.
        await call('create_block', {
            sheet: 'S',
            name: 'assum',
            position: {row: 0, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [{key: 'rate', values: {v: 10}}],
        })
        await call('create_block', {
            sheet: 'S',
            name: 'out',
            position: {row: 5, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [{key: 'total'}],
        })
        await call('set_block_cells', {
            changes: [
                {
                    block: 'out',
                    row_key: 'total',
                    field: 'v',
                    value: '=BLOCKREF("assum","rate","v")*100',
                },
            ],
        })

        type Described = {rows?: Array<{key: string; values: Record<string, unknown>}>}
        const readTotal = async (): Promise<unknown> => {
            const d = await call<Described>('describe_block', {
                name: 'out',
                include_rows: true,
            })
            return d.rows?.[0]?.values.v
        }
        expect(await readTotal()).toBe(1000)

        const preview = await call<{
            diff: Array<{block: string | null; field: string | null; before: unknown; after: unknown}>
        }>('preview_changes', {
            changes: [{block: 'assum', row_key: 'rate', field: 'v', value: 25}],
        })

        // It reports the cascade, not just the cell written.
        const moved = preview.diff.find((d) => d.block === 'out' && d.field === 'v')
        expect(moved?.before).toBe(1000)
        expect(moved?.after).toBe(2500)

        // And the live workbook is untouched.
        expect(await readTotal()).toBe(1000)
    })

    it('runs a grid of scenarios in one call, watching only the answer', async () => {
        // A sensitivity table is the commonest thing anyone does to a model, and
        // doing it by mutating-and-reverting is both slow and unsafe: a failure
        // mid-scan leaves the model on the last probe. Scenarios run each
        // hypothetical on its own temp branch, and `watch` keeps the result to
        // the numbers asked for — a 4x4 grid over a model that cascades into
        // dozens of cells would otherwise be hundreds of diff rows to answer
        // sixteen questions.
        await call('create_block', {
            sheet: 'S',
            name: 'assum',
            position: {row: 0, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [
                {key: 'a', values: {v: 2}},
                {key: 'b', values: {v: 3}},
            ],
        })
        await call('create_block', {
            sheet: 'S',
            name: 'out',
            position: {row: 5, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [{key: 'product'}],
        })
        await call('set_block_cells', {
            changes: [
                {
                    block: 'out',
                    row_key: 'product',
                    field: 'v',
                    value: '=BLOCKREF("assum","a","v")*BLOCKREF("assum","b","v")',
                },
            ],
        })

        type Described = {rows?: Array<{key: string; values: Record<string, unknown>}>}
        const product = async (): Promise<unknown> => {
            const d = await call<Described>('describe_block', {
                name: 'out',
                include_rows: true,
            })
            return d.rows?.[0]?.values.v
        }
        expect(await product()).toBe(6)

        const grid = [4, 5].flatMap((a) =>
            [10, 20].map((b) => ({
                label: `a=${a},b=${b}`,
                changes: [
                    {block: 'assum', row_key: 'a', field: 'v', value: a},
                    {block: 'assum', row_key: 'b', field: 'v', value: b},
                ],
            }))
        )

        const res = await call<{
            scenarios: Array<{
                label?: string
                watched?: Array<{value: unknown}>
                diff?: unknown[]
            }>
        }>('preview_changes', {
            scenarios: grid,
            watch: [{block: 'out', row_key: 'product', field: 'v'}],
        })

        // One result per scenario, in order, each carrying just the watched cell.
        expect(res.scenarios).toHaveLength(4)
        const byLabel = new Map(
            res.scenarios.map((x) => [x.label, x.watched?.[0]?.value])
        )
        expect(byLabel.get('a=4,b=10')).toBe(40)
        expect(byLabel.get('a=4,b=20')).toBe(80)
        expect(byLabel.get('a=5,b=10')).toBe(50)
        expect(byLabel.get('a=5,b=20')).toBe(100)
        // `watch` replaces the diff rather than adding to it.
        expect(res.scenarios[0]?.diff).toBeUndefined()

        // Scenarios are independent, not cumulative, and none of them stuck.
        expect(await product()).toBe(6)

        // The single-hypothetical shape still returns a diff, as it always did.
        const one = await call<{diff: Array<{after: unknown}>}>('preview_changes', {
            changes: [{block: 'assum', row_key: 'a', field: 'v', value: 7}],
        })
        expect(one.diff.some((d) => d.after === 21)).toBe(true)
        expect(await product()).toBe(6)
    })

    it('traces dependencies in both directions, and says how precise it is', async () => {
        // The engine's trace only resolved normal ranges: every block cell and
        // every BLOCKREF came back with no dependencies at all. "Nothing depends
        // on this assumption" is the most dangerous possible wrong answer to
        // give before an edit.
        await call('create_block', {
            sheet: 'S',
            name: 'assum',
            position: {row: 0, col: 0},
            fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
            initial_rows: [
                {key: 'rate', values: {v: 3}},
                {key: 'other', values: {v: 9}},
            ],
        })
        await call('create_block', {
            sheet: 'S',
            name: 'proj',
            position: {row: 6, col: 0},
            fields: [
                {name: 'k'},
                {name: 'base', field_type: 'number'},
                {name: 'scaled', field_type: 'number'},
            ],
            initial_rows: [{key: 'r1', values: {base: 10}}],
        })
        await call('set_field_rule', {
            block: 'proj',
            field: 'scaled',
            value_formula: '=#FIELD("base")*BLOCKREF("assum","rate","v")',
        })

        type Traced = {
            approximate?: boolean
            precedents?: Array<{block: string | null; field: string | null; ref: string; scope: string}>
            dependents?: Array<{block: string | null; row_key: string | null; field: string | null; scope: string}>
        }

        // Inside a block the graph is per-cell, so this is exact: no
        // approximation flag, and the sibling is named rather than located.
        const inside = await call<Traced>('trace', {
            target: {block: 'proj', row_key: 'r1', field: 'scaled'},
            direction: 'precedents',
        })
        const base = inside.precedents?.find((x) => x.field === 'base')
        expect(base?.block).toBe('proj')
        expect(base?.scope).toBe('cell')
        // ...and it sees the BLOCKREF too, at field granularity.
        expect(
            inside.precedents?.some((x) => x.block === 'assum' && x.scope === 'field')
        ).toBe(true)
        expect(inside.approximate).toBe(true)

        // The reverse direction is the one formula text cannot answer.
        const reverse = await call<Traced>('trace', {
            target: {block: 'assum', row_key: 'rate', field: 'v'},
            direction: 'dependents',
        })
        expect(
            reverse.dependents?.some(
                (d) => d.block === 'proj' && d.field === 'scaled'
            )
        ).toBe(true)
        // Block dependencies are tracked per field, not per row, so this is an
        // over-approximation and has to admit as much.
        expect(reverse.approximate).toBe(true)
        expect(reverse.dependents?.every((d) => d.scope === 'field')).toBe(true)

        // Ordinary cells stay exact and unflagged.
        await call('set_cells', {
            sheetIdx: 0,
            cells: [
                {row: 20, col: 0, content: '5'},
                {row: 21, col: 0, content: '=A21*2'},
            ],
        })
        const plain = await call<Traced>('trace', {
            target: {row: 21, col: 0},
        })
        expect(plain.precedents?.map((x) => x.ref)).toEqual(['A21'])
        expect(plain.approximate).toBeUndefined()
    })

    it('rejects a range straddling a block boundary without dying', async () => {
        // A `Range` is wholly normal or wholly one block's, so B1:B10 with B1
        // inside a block has no representation. The engine used to panic here,
        // which killed the session — an agent writing an over-wide SUM is
        // routine, so it has to degrade to an ordinary bad-formula outcome.
        await call('create_block', {
            sheet: 'Model',
            name: 'items',
            position: {row: 0, col: 0},
            fields: [{name: 'id'}, {name: 'qty', field_type: 'number'}],
            initial_rows: [{key: 'x', values: {qty: 4}}],
        })
        await call('set_cells', {
            sheetIdx: 1,
            cells: [{row: 0, col: 4, content: '=SUM(B1:B10)*2'}],
        })

        // Session still alive and computing.
        expect(await call('eval_formula', {expr: '=6*7'})).toEqual({
            type: 'number',
            value: 42,
        })
        // And the block is intact.
        const d = await call<{keys: string[]}>('describe_block', {
            name: 'items',
        })
        expect(d.keys).toEqual(['x'])
    })

    it('exports base64 for hosts with no shared filesystem', async () => {
        await call('set_cells', {
            sheetIdx: 0,
            cells: [{row: 0, col: 0, content: 'hello'}],
        })
        const exported = await call<{base64: string; bytes: number}>(
            'export_xlsx'
        )
        expect(exported.bytes).toBeGreaterThan(0)
        const decoded = Buffer.from(exported.base64, 'base64')
        expect(decoded.length).toBe(exported.bytes)
        expect([...decoded.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04])

        // Round-trips back in through the same door.
        const reopened = await call<{source: string}>('open_workbook', {
            xlsx_base64: exported.base64,
        })
        expect(reopened.source).toBe('bytes')
    })

    it('reports a bad tool argument as a usable error, not a crash', async () => {
        await expect(
            call('describe_block', {name: 'does_not_exist'})
        ).rejects.toThrow(/does_not_exist/)
    })
})
