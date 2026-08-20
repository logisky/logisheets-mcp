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
        expect(names).toHaveLength(15)
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
