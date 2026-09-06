/**
 * The file Excel actually opens.
 *
 * Every other test here asks the engine what it thinks the workbook contains.
 * This one asks the `.xlsx`, because the two can disagree in a way only Excel
 * notices: a package can hold every correct value and still be refused at the
 * door. That is not hypothetical — logisky/logisheets-mcp#1 was two such
 * defects at once, both invisible to the engine's own round-trip:
 *
 *   * `[Content_Types].xml` declared `/logisheets/data.xml` with an empty
 *     `ContentType`. OPC requires a media type, so the whole package was
 *     invalid — and because the fault is at the package layer rather than in
 *     any one part, Excel's repair report named nothing at all, which is what
 *     made it so hard to place.
 *   * A generated chart's `graphicFrame` had no `xdr:xfrm`. The schema requires
 *     it; Excel repaired the drawing part on open.
 *
 * Neither showed up in a save/reopen through LogiSheets, because LogiSheets
 * wrote and read them consistently. So these assertions deliberately do not go
 * through the engine: they unzip the bytes and check the package on its own
 * terms, the way a foreign consumer does.
 *
 * The fixes live in the LogiSheets engine, not here. This is the tripwire — an
 * engine bump that regresses either one fails at the boundary this server is
 * responsible for: the file a human is handed.
 */

import {mkdtemp, rm, readFile} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {inflateRawSync} from 'node:zlib'
import {afterEach, beforeEach, describe, expect, it} from 'vitest'
import type {Tool, ToolContext} from 'logisheets-logician'
import {createServer} from './server.js'
import {WorkbookSession} from './session.js'

/** Every part of an .xlsx, as name -> text. Directory entries are skipped. */
function xlsxParts(buf: Buffer): Map<string, string> {
    let i = 0
    const out = new Map<string, string>()
    while (i + 30 <= buf.length && buf.readUInt32LE(i) === 0x04034b50) {
        const method = buf.readUInt16LE(i + 8)
        const size = buf.readUInt32LE(i + 18)
        const nameLen = buf.readUInt16LE(i + 26)
        const extraLen = buf.readUInt16LE(i + 28)
        const name = buf.subarray(i + 30, i + 30 + nameLen).toString('utf8')
        const start = i + 30 + nameLen + extraLen
        const data = buf.subarray(start, start + size)
        if (!name.endsWith('/')) {
            out.set(
                name,
                (method === 0 ? data : inflateRawSync(data)).toString('utf8')
            )
        }
        i = start + size
    }
    return out
}

/**
 * Every content type the package declares, `Default`s and `Override`s alike.
 *
 * Returned as raw strings rather than checked here so a failure can name the
 * part that is wrong, not merely report that something is.
 */
function declaredContentTypes(
    contentTypesXml: string
): Array<{part: string; type: string}> {
    const out: Array<{part: string; type: string}> = []
    const re =
        /<(Default|Override)\s+(?:Extension|PartName)="([^"]*)"\s+ContentType="([^"]*)"/g
    for (const m of contentTypesXml.matchAll(re)) {
        out.push({part: m[2] ?? '', type: m[3] ?? ''})
    }
    return out
}

describe('the .xlsx as a foreign consumer sees it', () => {
    let session: WorkbookSession
    let call: <T = unknown>(
        name: string,
        args?: Record<string, unknown>
    ) => Promise<T>
    let dir: string

    beforeEach(async () => {
        const created = createServer({mode: 'core', log: () => {}})
        session = created.session
        const tools: Map<string, Tool> = created.tools
        const base: Omit<ToolContext, 'workbook'> = {
            signal: new AbortController().signal,
            confirm: async () => true,
            log: () => {},
        }
        call = async <T,>(name: string, args: Record<string, unknown> = {}) => {
            const tool = tools.get(name)
            if (tool === undefined) throw new Error(`no such tool: ${name}`)
            const r = await tool.handler(args, {
                ...base,
                workbook: session.client,
            })
            return r.data as T
        }
        dir = await mkdtemp(join(tmpdir(), 'logisheets-pkg-'))
    })

    afterEach(async () => {
        session.close()
        await rm(dir, {recursive: true, force: true})
    })

    /**
     * Build the workbook from the bug report: two blocks, a field rule whose
     * result is text for one row, and optionally a chart over the block.
     */
    async function buildAndSave(
        withChart: boolean,
        opts: {resolveBlockRefs: boolean}
    ): Promise<Buffer> {
        await call('create_block', {
            sheet: 'Model',
            name: 'params',
            description: 'Scenario multiplier.',
            position: {row: 0, col: 0},
            fields: [
                {name: 'key', field_type: 'string'},
                {name: 'value', field_type: 'number'},
            ],
            initial_rows: [{key: 'factor', values: {value: 1.2}}],
        })
        await call('create_block', {
            sheet: 'Model',
            name: 'quotes',
            description: 'Unit conversion.',
            position: {row: 4, col: 0},
            fields: [
                {name: 'key', field_type: 'string'},
                {name: 'usd_per_tonne', field_type: 'number'},
                {name: 'kg_per_1000', field_type: 'number'},
                {name: 'usd_per_1000', field_type: 'number', num_fmt: '0.00'},
            ],
            initial_rows: [
                {key: 'normal', values: {usd_per_tonne: 2400, kg_per_1000: 12.5}},
                {key: 'missing', values: {kg_per_1000: 12.5}},
            ],
        })
        await call('set_field_rule', {
            block: 'quotes',
            field: 'usd_per_1000',
            value_formula:
                '=IF(COUNT(#FIELD("usd_per_tonne"),#FIELD("kg_per_1000"))<2,"n/a",' +
                '#FIELD("usd_per_tonne")*#FIELD("kg_per_1000")/1000*BLOCKREF("params","factor","value"))',
        })

        if (withChart) {
            const d = await call<{sheet_idx: number; block_id: number}>(
                'describe_block',
                {name: 'quotes'}
            )
            await call('chart_from_block', {
                sheetIdx: d.sheet_idx,
                blockId: d.block_id,
                valueFields: ['usd_per_1000'],
                chartType: 'col',
                title: 'USD per 1000',
            })
        }

        const file = join(dir, `${withChart ? 'chart' : 'plain'}.xlsx`)
        await call('save_workbook', {
            path: file,
            resolve_block_refs: opts.resolveBlockRefs,
        })
        return readFile(file)
    }

    // Both save modes and both shapes: the original report blamed
    // `resolve_block_refs`, but the malformed package had nothing to do with
    // it — pinning all four keeps that conclusion from quietly rotting.
    for (const withChart of [false, true]) {
        for (const resolveBlockRefs of [false, true]) {
            const label = `${withChart ? 'with a chart' : 'blocks only'}, ${
                resolveBlockRefs ? 'refs resolved' : 'refs kept'
            }`

            it(`declares a real media type for every part (${label})`, async () => {
                const parts = xlsxParts(
                    await buildAndSave(withChart, {resolveBlockRefs})
                )
                const xml = parts.get('[Content_Types].xml')
                expect(xml).toBeDefined()

                const declared = declaredContentTypes(xml!)
                expect(declared.length).toBeGreaterThan(0)
                for (const {part, type} of declared) {
                    // `type/subtype`, both halves non-empty. An empty string
                    // here is what made Excel refuse the whole file.
                    expect(
                        type,
                        `content type for ${part} is not a media type`
                    ).toMatch(/^[^/\s]+\/[^/\s]+$/)
                }

                // Our own part is the one that regressed, so name it directly.
                const appData = declared.find(
                    (d) => d.part === '/logisheets/data.xml'
                )
                if (appData !== undefined) {
                    expect(appData.type).not.toBe('')
                }
            })

            it(`gives every chart frame an xfrm (${label})`, async () => {
                const parts = xlsxParts(
                    await buildAndSave(withChart, {resolveBlockRefs})
                )
                const drawings = [...parts.entries()].filter(([n]) =>
                    n.startsWith('xl/drawings/drawing')
                )
                // A chartless workbook has no drawing at all; that is the
                // correct outcome, not a skipped assertion.
                expect(drawings.length).toBe(withChart ? 1 : 0)

                for (const [name, xml] of drawings) {
                    const frames =
                        xml.match(/<xdr:graphicFrame[\s\S]*?<\/xdr:graphicFrame>/g) ??
                        []
                    expect(frames.length, `${name} has no graphicFrame`).toBe(1)
                    for (const frame of frames) {
                        // May be empty — Excel recomputes the geometry from the
                        // anchor — but it has to be there, or Excel repairs the
                        // drawing part on open.
                        expect(frame, `${name}: graphicFrame without xfrm`).toMatch(
                            /<xdr:xfrm\s*\/>|<xdr:xfrm[\s>]/
                        )
                    }
                }
            })
        }
    }

    /**
     * Cheap structural sanity on the rest of the package, so a future engine
     * bump that drops a part is caught here rather than by a human in Excel.
     */
    it('ships a package whose declared parts are all present', async () => {
        const parts = xlsxParts(await buildAndSave(true, {resolveBlockRefs: true}))
        const xml = parts.get('[Content_Types].xml')!

        for (const {part} of declaredContentTypes(xml)) {
            // Defaults are extensions, Overrides are absolute part names.
            if (!part.startsWith('/')) continue
            expect(parts.has(part.slice(1)), `declared but absent: ${part}`).toBe(
                true
            )
        }

        // The workbook root and its relationships, without which nothing opens.
        expect(parts.has('[Content_Types].xml')).toBe(true)
        expect(parts.has('_rels/.rels')).toBe(true)
        expect(parts.has('xl/workbook.xml')).toBe(true)
        expect(parts.has('xl/_rels/workbook.xml.rels')).toBe(true)
    })
})
