/**
 * The handover claim, tested: a session that did NOT build the model can pick
 * the file up and understand it.
 *
 * Every other test in this repo checks what one session can do while it still
 * remembers what it did. This one checks the opposite — the context is gone,
 * the only thing left is the `.xlsx`, and a new agent has to work out what the
 * model is, what computes what, and what it may safely change. That is the
 * common case in practice: the conversation that built a spreadsheet is almost
 * never the conversation that has to answer a question about it.
 *
 * Three things are measured, not asserted:
 *   1. WHAT comes back — schema in the model's own words, or coordinates the
 *      agent has to resolve against a label column before they mean anything.
 *   2. WHAT IT COSTS — how many calls and how many bytes of context orienting
 *      takes, and whether that grows with the size of the data.
 *   3. WHETHER IT WORKED — the new session writes a formula against a model it
 *      has never seen, from orientation alone, and the engine agrees with an
 *      answer computed independently here in TypeScript.
 *
 * Costs are read off the real wire: each session drives the server through an
 * MCP `Client`, and the meter counts the text a host would put in front of the
 * model. Nothing is estimated.
 */

import {mkdtemp, rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {Client} from '@modelcontextprotocol/sdk/client/index.js'
import {InMemoryTransport} from '@modelcontextprotocol/sdk/inMemory.js'
import type {CallToolResult} from '@modelcontextprotocol/sdk/types.js'
import {afterEach, beforeEach, describe, expect, it} from 'vitest'
import {createServer} from './server.js'
import type {WorkbookSession} from './session.js'

/** An A1-style coordinate: what a rule must NOT contain to be self-describing. */
const A1_REF = /\$?[A-Z]{1,2}\$?\d{1,4}/

interface Group {
    sheet_name: string
    blocks: Array<{
        name: string
        fields: string[]
        key_field: string | null
        derived_fields: string[]
        key_count: number
    }>
}

interface Described {
    block: string
    fields: Array<{name: string; value_formula: string | null}>
    keys: string[]
    rows?: Array<{key: string; values: Record<string, unknown>}>
}

interface Traced {
    approximate?: boolean
    dependents?: Array<{block: string | null; field: string | null}>
}

/**
 * One MCP session, with a meter on it.
 *
 * `calls` and `bytes` are the two numbers a host pays: one round trip and its
 * text, per call. Bytes are counted over exactly the text blocks the server
 * returns — the same string the host hands the model.
 */
class Session {
    calls = 0
    bytes = 0

    private constructor(
        private readonly client: Client,
        private readonly session: WorkbookSession
    ) {}

    static async open(): Promise<Session> {
        const created = createServer({mode: 'core', log: () => {}})
        const client = new Client({name: 'cold-read', version: '0.0.0'})
        const [clientTransport, serverTransport] =
            InMemoryTransport.createLinkedPair()
        await Promise.all([
            created.server.connect(serverTransport),
            client.connect(clientTransport),
        ])
        return new Session(client, created.session)
    }

    async call<T>(name: string, args: Record<string, unknown> = {}): Promise<T> {
        const res = (await this.client.callTool({
            name,
            arguments: args,
        })) as CallToolResult
        const text = res.content
            .filter((c): c is {type: 'text'; text: string} => c.type === 'text')
            .map((c) => c.text)
            .join('\n')
        this.calls += 1
        this.bytes += Buffer.byteLength(text, 'utf8')
        if (res.isError === true) throw new Error(text)
        for (const line of text.split('\n').reverse()) {
            const i = line.search(/[[{]/)
            if (i >= 0) {
                try {
                    return JSON.parse(line.slice(i)) as T
                } catch {
                    /* a display line, not the payload */
                }
            }
        }
        return undefined as T
    }

    reset(): void {
        this.calls = 0
        this.bytes = 0
    }

    async close(): Promise<void> {
        await this.client.close()
        this.session.close()
    }
}

/** The model under test: two assumptions and a projection that reads them. */
const UNITS = [100, 140, 196, 274, 384]
const PRICE = 9.5
const MARGIN = 0.3
const TAX = 0.25
const REVENUE_RULE = '=#FIELD("units")*#FIELD("price")'
const PROFIT_RULE =
    '=#FIELD("revenue")*BLOCKREF("assum","margin","v")*(1-BLOCKREF("assum","tax","v"))'

/** Build the model in `author` and save it. Returns the path. */
async function buildAndSave(
    author: Session,
    path: string,
    years: number
): Promise<string> {
    await author.call('create_block', {
        sheet: 'Model',
        name: 'assum',
        position: {row: 0, col: 0},
        fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
        initial_rows: [
            {key: 'margin', values: {v: MARGIN}},
            {key: 'tax', values: {v: TAX}},
        ],
    })
    await author.call('create_block', {
        sheet: 'Model',
        name: 'proj',
        position: {row: 6, col: 0},
        fields: [
            {name: 'year'},
            {name: 'units', field_type: 'number'},
            {name: 'price', field_type: 'number'},
            {name: 'revenue', field_type: 'number'},
            {name: 'profit', field_type: 'number'},
        ],
        initial_rows: Array.from({length: years}, (_, i) => ({
            key: `y${i + 1}`,
            values: {units: units(i), price: PRICE},
        })),
    })
    await author.call('set_field_rule', {
        block: 'proj',
        field: 'revenue',
        value_formula: REVENUE_RULE,
    })
    await author.call('set_field_rule', {
        block: 'proj',
        field: 'profit',
        value_formula: PROFIT_RULE,
    })
    await author.call('save_workbook', {path})
    return path
}

/** Units for year i, extended past the literal table by repeating the last. */
function units(i: number): number {
    return UNITS[i] ?? UNITS[UNITS.length - 1]!
}

describe('a session that did not build the model', () => {
    let author: Session
    let dir: string

    beforeEach(async () => {
        author = await Session.open()
        dir = await mkdtemp(join(tmpdir(), 'logisheets-mcp-cold-'))
    })

    afterEach(async () => {
        await author.close()
        await rm(dir, {recursive: true, force: true})
    })

    it('recovers the schema from the file, in the model’s own words', async () => {
        const path = await buildAndSave(author, join(dir, 'model.xlsx'), 5)

        // A genuinely new session: its own server, its own workbook, no memory
        // of the one above. Reusing the author's session would prove only that
        // the engine remembers, which is not the claim.
        const fresh = await Session.open()
        try {
            await fresh.call('open_workbook', {path})

            // ONE call, and the workbook has introduced itself: which tables
            // exist, what each column is called, which column is the row key,
            // which columns the engine owns, and how many records there are.
            const groups = await fresh.call<Group[]>('list_blocks')
            const blocks = groups.flatMap((g) => g.blocks)
            expect(blocks.map((b) => b.name).sort()).toEqual(['assum', 'proj'])

            const proj = blocks.find((b) => b.name === 'proj')!
            expect(proj.fields).toEqual([
                'year',
                'units',
                'price',
                'revenue',
                'profit',
            ])
            expect(proj.key_field).toBe('year')
            // The two the engine computes — writing to them would be refused,
            // and knowing that up front saves a failed call and a retry.
            expect(proj.derived_fields).toEqual(['revenue', 'profit'])
            expect(proj.key_count).toBe(5)

            // And the rules themselves survived the file intact.
            const described = await fresh.call<Described>('describe_block', {
                name: 'proj',
            })
            const rules = new Map(
                described.fields.map((f) => [f.name, f.value_formula])
            )
            expect(rules.get('revenue')).toBe(REVENUE_RULE.slice(1))
            expect(rules.get('profit')).toBe(PROFIT_RULE.slice(1))

            // The measurement that matters more than the answers: a rule comes
            // back saying what it MEANS. `revenue x margin x (1 - tax)` is an
            // explanation; `D8*$B$1*(1-$B$2)` is a second lookup problem, and
            // one the next session would have to solve over again because the
            // file records the coordinates but not what they stand for.
            for (const [field, rule] of rules) {
                if (rule === null) continue
                expect(
                    A1_REF.test(rule),
                    `rule for ${field} leaked a coordinate: ${rule}`
                ).toBe(false)
                expect(/#FIELD\(|BLOCKREF\(/.test(rule)).toBe(true)
            }
            // Nothing in the projection is anonymous: an input field has no
            // rule, and every field with one names its sources.
            expect(rules.get('units')).toBeNull()

            // Understanding is proven by acting, not by prose. The new session
            // now writes a formula against a model it has never seen, using
            // only the three names orientation gave it, and the engine has to
            // agree with the arithmetic done independently right here.
            const expected = units(2) * PRICE * MARGIN * (1 - TAX)
            const evaluated = await fresh.call<{value: unknown}>(
                'eval_formula',
                {expr: `=BLOCKREF("proj","y3","profit")`}
            )
            expect(evaluated.value).toBeCloseTo(expected, 6)
        } finally {
            await fresh.close()
        }
    })

    it('knows what reads an assumption before it changes one', async () => {
        const path = await buildAndSave(author, join(dir, 'model.xlsx'), 5)

        const fresh = await Session.open()
        try {
            await fresh.call('open_workbook', {path})
            // The question formula text cannot answer. Reading a cell tells you
            // what it depends on; nothing in the file tells you what depends on
            // IT — which is exactly what you need before editing an assumption
            // in a model you did not write.
            const back = await fresh.call<Traced>('trace', {
                target: {block: 'assum', row_key: 'margin', field: 'v'},
                direction: 'dependents',
            })
            expect(
                back.dependents?.some(
                    (d) => d.block === 'proj' && d.field === 'profit'
                )
            ).toBe(true)
            // Block dependencies are tracked per field rather than per row, so
            // the answer is an over-approximation and says so rather than
            // implying a precision it does not have.
            expect(back.approximate).toBe(true)
        } finally {
            await fresh.close()
        }
    })

    it('costs the same to orient on a big model as on a small one', async () => {
        // What makes a schema worth declaring is that reading it is O(columns),
        // not O(cells). A coordinate grid has no such summary: understanding it
        // means reading it, so the bill for arriving grows with the data and
        // is paid again every session.
        const small = await buildAndSave(author, join(dir, 'small.xlsx'), 5)
        await author.call('add_block_rows', {
            block: 'proj',
            rows: Array.from({length: 100}, (_, i) => ({
                key: `y${i + 6}`,
                values: {units: units(i), price: PRICE},
            })),
        })
        const bigPath = join(dir, 'large.xlsx')
        await author.call('save_workbook', {path: bigPath})

        const orient = async (path: string) => {
            const s = await Session.open()
            try {
                await s.call('open_workbook', {path})
                s.reset() // the open is fixed cost; measure the reading
                const groups = await s.call<Group[]>('list_blocks')
                const rows =
                    groups
                        .flatMap((g) => g.blocks)
                        .find((b) => b.name === 'proj')?.key_count ?? 0
                return {calls: s.calls, bytes: s.bytes, rows, session: s}
            } catch (err) {
                await s.close()
                throw err
            }
        }

        const a = await orient(small)
        const b = await orient(bigPath)
        try {
            expect(a.rows).toBe(5)
            expect(b.rows).toBe(105)

            // One call each, and the payload is the schema — so twenty-one
            // times the data costs a couple of extra digits, not twenty-one
            // times the context. Measured: 540 B for five rows, 545 B for a
            // hundred and five. The bounds are loose around those so a wording
            // change in the tool's display line doesn't fail the suite, but
            // anything that makes orientation grow with the data will.
            expect(a.calls).toBe(1)
            expect(b.calls).toBe(1)
            expect(b.bytes - a.bytes).toBeLessThan(32)
            expect(a.bytes).toBeLessThan(800)

            // The honest control: the cheap number is cheap because the schema
            // is small, not because the workbook is empty. Ask for the data and
            // it costs what the data costs — measured at 11 kB against the
            // 545 B schema read, twenty times the size, exactly as it should
            // be. What the schema buys is the CHOICE of not reading it.
            const full = await b.session.call<Described>('describe_block', {
                name: 'proj',
                include_rows: true,
            })
            expect(full.rows).toHaveLength(105)
            expect(b.session.bytes).toBeGreaterThan(10 * b.bytes)
        } finally {
            await a.session.close()
            await b.session.close()
        }
    })
})
