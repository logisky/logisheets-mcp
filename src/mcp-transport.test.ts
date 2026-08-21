/**
 * The protocol layer: a real MCP client talking to this server over the SDK's
 * in-memory transport. Everything a host does — handshake, tools/list,
 * tools/call — goes through the same code path the stdio binary uses.
 */

import {mkdtemp, rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {Client} from '@modelcontextprotocol/sdk/client/index.js'
import {InMemoryTransport} from '@modelcontextprotocol/sdk/inMemory.js'
import type {CallToolResult} from '@modelcontextprotocol/sdk/types.js'
import {afterEach, beforeEach, describe, expect, it} from 'vitest'
import {createServer, INSTRUCTIONS} from './server.js'
import {WORKBOOK_URI, XLSX_MIME} from './lifecycle.js'
import type {WorkbookSession} from './session.js'

/** First text block of a tool result. */
function text(result: CallToolResult): string {
    return result.content
        .filter((c): c is {type: 'text'; text: string} => c.type === 'text')
        .map((c) => c.text)
        .join('\n')
}

describe('MCP protocol surface', () => {
    let client: Client
    let session: WorkbookSession
    let dir: string

    beforeEach(async () => {
        const created = createServer({mode: 'core', log: () => {}})
        session = created.session
        client = new Client({name: 'test-host', version: '0.0.0'})
        const [clientTransport, serverTransport] =
            InMemoryTransport.createLinkedPair()
        await Promise.all([
            created.server.connect(serverTransport),
            client.connect(clientTransport),
        ])
        dir = await mkdtemp(join(tmpdir(), 'logisheets-mcp-proto-'))
    })

    afterEach(async () => {
        await client.close()
        session.close()
        await rm(dir, {recursive: true, force: true})
    })

    it('advertises tools with valid JSON Schema and correct hints', async () => {
        const {tools} = await client.listTools()
        expect(tools).toHaveLength(20)

        for (const t of tools) {
            expect(t.name).toMatch(/^[a-z][a-z0-9_]*$/)
            expect(t.description).toBeTruthy()
            expect(t.inputSchema.type).toBe('object')
        }

        const byName = new Map(tools.map((t) => [t.name, t]))
        // Reads are marked read-only so a host needn't prompt for them.
        expect(byName.get('eval_formula')?.annotations?.readOnlyHint).toBe(true)
        expect(byName.get('describe_block')?.annotations?.readOnlyHint).toBe(
            true
        )
        // Writes are not, so a host can gate them.
        expect(byName.get('set_block_cells')?.annotations?.readOnlyHint).toBe(
            false
        )
        expect(byName.get('create_block')?.annotations?.readOnlyHint).toBe(false)
        // Data-discarding tools are flagged more strongly than mere writes.
        expect(
            byName.get('delete_block_rows')?.annotations?.destructiveHint
        ).toBe(true)
        expect(byName.get('create_block')?.annotations?.destructiveHint).toBe(
            false
        )

        // Required fields survive the readonly->mutable conversion.
        expect(byName.get('describe_block')?.inputSchema.required).toEqual([
            'name',
        ])
    })

    it('sends the agent instructions on connect', () => {
        expect(client.getInstructions()).toBe(INSTRUCTIONS)
    })

    it('runs the whole loop over the protocol', async () => {
        const created = await client.callTool({
            name: 'create_block',
            arguments: {
                sheet: 'Budget',
                name: 'lines',
                position: {row: 0, col: 0},
                fields: [
                    {name: 'item'},
                    {name: 'qty', field_type: 'number'},
                    {name: 'unit', field_type: 'number'},
                    {name: 'total', field_type: 'number'},
                ],
                initial_rows: [{key: 'widget', values: {qty: 3, unit: 4}}],
            },
        })
        expect(created.isError).toBeFalsy()

        await client.callTool({
            name: 'add_block_rows',
            arguments: {
                block: 'lines',
                rows: [{key: 'gizmo', values: {qty: 10, unit: 2.5}}],
            },
        })

        // The engine does the arithmetic: 3*4 + 10*2.5 = 37.
        const evaluated = await client.callTool({
            name: 'eval_formula',
            arguments: {
                expr: '=BLOCKREF("lines","widget","qty")*BLOCKREF("lines","widget","unit")+BLOCKREF("lines","gizmo","qty")*BLOCKREF("lines","gizmo","unit")',
            },
        })
        expect(evaluated.isError).toBeFalsy()
        expect(text(evaluated as CallToolResult)).toContain('37')

        // Read the structured memory back by name.
        const described = await client.callTool({
            name: 'describe_block',
            arguments: {name: 'lines', include_rows: true},
        })
        const body = text(described as CallToolResult)
        expect(body).toContain('widget')
        expect(body).toContain('gizmo')

        // A field the engine computes for every row, present and future: the
        // agent declares the rule once and never writes those cells itself.
        const ruled = await client.callTool({
            name: 'set_field_rule',
            arguments: {
                block: 'lines',
                field: 'total',
                value_formula: '=#FIELD("qty")*#FIELD("unit")',
            },
        })
        expect(ruled.isError).toBeFalsy()

        const withTotals = await client.callTool({
            name: 'describe_block',
            arguments: {name: 'lines', include_rows: true},
        })
        const rows = (
            JSON.parse(
                text(withTotals as CallToolResult).split('\n').at(-1) ?? '{}'
            ) as {
                rows?: Array<{key: string; values: Record<string, unknown>}>
            }
        ).rows
        const byKey = new Map(rows?.map((r) => [r.key, r.values]))
        expect(byKey.get('widget')?.total).toBe(12) // 3 * 4
        expect(byKey.get('gizmo')?.total).toBe(25) // 10 * 2.5

        // Hand back a real file.
        const out = join(dir, 'budget.xlsx')
        const saved = await client.callTool({
            name: 'save_workbook',
            arguments: {path: out},
        })
        expect(saved.isError).toBeFalsy()
        expect(text(saved as CallToolResult)).toContain(out)
    })

    it('returns tool failures as errors the agent can read, not protocol faults', async () => {
        const result = await client.callTool({
            name: 'describe_block',
            arguments: {name: 'nope'},
        })
        expect(result.isError).toBe(true)
        expect(text(result as CallToolResult)).toContain('nope')

        // The session is still healthy afterwards.
        const ok = await client.callTool({
            name: 'eval_formula',
            arguments: {expr: '=1+1'},
        })
        expect(ok.isError).toBeFalsy()
        expect(text(ok as CallToolResult)).toContain('2')
    })

    it('serializes concurrent calls instead of letting them race', async () => {
        // MCP does not promise to serialize: JSON-RPC allows pipelining and the
        // SDK dispatches concurrently. Handlers read state and then write it
        // across an await — `create_block` checks whether its sheet exists and
        // creates it if not — so three in flight at once each saw the sheet
        // missing, each tried to create it, and two failed with a bare
        // "status code 1".
        const results = await Promise.all(
            ['alpha', 'beta', 'gamma'].map((name, i) =>
                client.callTool({
                    name: 'create_block',
                    arguments: {
                        sheet: 'Shared',
                        name,
                        position: {row: i * 4, col: 0},
                        fields: [{name: 'k'}, {name: 'v', field_type: 'number'}],
                        initial_rows: [{key: 'r1', values: {v: i}}],
                    },
                })
            )
        )
        for (const r of results) {
            expect(text(r as CallToolResult)).not.toContain('status code')
            expect(r.isError).toBeFalsy()
        }

        // All three landed, each with its own block id — the sheet was created
        // exactly once.
        const listed = await client.callTool({name: 'list_blocks', arguments: {}})
        const groups = JSON.parse(
            text(listed as CallToolResult).slice(
                text(listed as CallToolResult).indexOf('[')
            )
        ) as Array<{blocks: Array<{name: string; block_id: number}>}>
        const blocks = groups.flatMap((g) => g.blocks)
        expect(blocks.map((b) => b.name).sort()).toEqual([
            'alpha',
            'beta',
            'gamma',
        ])
        expect(new Set(blocks.map((b) => b.block_id)).size).toBe(3)
    })

    it("relays the engine's own reason for a rejected transaction", async () => {
        // Every engine rejection is status code 1, so the code alone says
        // nothing. The reason now rides along on the effect.
        const {server} = createServer({mode: 'full', log: () => {}})
        const host = new Client({name: 'test-host-full', version: '0.0.0'})
        const [c, s2] = InMemoryTransport.createLinkedPair()
        await Promise.all([server.connect(s2), host.connect(c)])

        await host.callTool({name: 'open_workbook', arguments: {}})
        const bad = await host.callTool({
            name: 'delete_sheet',
            arguments: {idx: 9},
        })
        expect(bad.isError).toBe(true)
        const msg = text(bad as CallToolResult)
        expect(msg).toContain('exceeds the maximum')
        expect(msg).not.toContain('status code')
        await host.close()
    })

    it('hands the workbook back as a resource, not as bytes in the reply', async () => {
        // A tool result goes into the model's context, so returning the file
        // there costs ~1.4 KB of text per KB of workbook and teaches the model
        // nothing. The protocol's answer is a resource: the tool returns a link,
        // and the host reads the bytes out of band.
        await client.callTool({name: 'open_workbook', arguments: {}})
        await client.callTool({
            name: 'set_cells',
            arguments: {
                sheetIdx: 0,
                cells: [{row: 0, col: 0, content: '=6*7'}],
            },
        })

        const out = join(dir, 'handed-back.xlsx')
        const saved = await client.callTool({
            name: 'save_workbook',
            arguments: {path: out},
        })
        const link = (saved as CallToolResult).content.find(
            (c) => c.type === 'resource_link'
        )
        expect(link).toBeDefined()
        expect((link as {uri: string}).uri).toBe(WORKBOOK_URI)
        expect((link as {mimeType?: string}).mimeType).toBe(XLSX_MIME)
        // Named after the file the human is being handed.
        expect((link as {name: string}).name).toBe(out)
        // The reply itself stays small — no base64 payload rode along.
        const text = JSON.stringify((saved as CallToolResult).content)
        expect(text.length).toBeLessThan(1000)

        // The resource is listed...
        const {resources} = await client.listResources()
        expect(resources.map((r) => r.uri)).toContain(WORKBOOK_URI)

        // ...and reading it yields a real .xlsx, formula and all.
        const read = await client.readResource({uri: WORKBOOK_URI})
        const blob = read.contents[0] as {blob: string; mimeType?: string}
        expect(blob.mimeType).toBe(XLSX_MIME)
        const bytes = Buffer.from(blob.blob, 'base64')
        expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04])
        expect(bytes.length).toBeGreaterThan(1000)
    })

    it('offers no resource until a workbook exists', async () => {
        const {server} = createServer({mode: 'core', log: () => {}})
        const host = new Client({name: 'test-host-fresh', version: '0.0.0'})
        const [c, s2] = InMemoryTransport.createLinkedPair()
        await Promise.all([server.connect(s2), host.connect(c)])
        const {resources} = await host.listResources()
        expect(resources).toEqual([])
        await host.close()
    })

    it('reports an unknown tool without dropping the connection', async () => {
        const missing = await client.callTool({
            name: 'no_such_tool',
            arguments: {},
        })
        expect(missing.isError).toBe(true)
        expect(text(missing as CallToolResult)).toContain('no_such_tool')

        const ok = await client.callTool({
            name: 'eval_formula',
            arguments: {expr: '=2*3'},
        })
        expect(text(ok as CallToolResult)).toContain('6')
    })
})
