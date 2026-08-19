#!/usr/bin/env node
/**
 * What an agent can do with a real spreadsheet engine.
 *
 * Runs the exact tool sequence an LLM would, over real stdio against the built
 * binary — the same code path Claude Desktop drives. No LLM is involved: the
 * point is what the ENGINE guarantees, and hard-coding the calls makes those
 * guarantees checkable rather than a story about a chat session.
 *
 *   npm run build && npm run demo
 *
 * Three claims, each asserted rather than asserted-at:
 *
 *   1. No number here was produced by a language model. Every total is a
 *      formula the engine evaluated, and the .xlsx carries the formulas, so a
 *      human can change an input in Excel and watch it recalculate.
 *   2. Structured memory survives structural edits. The agent addresses cells
 *      as (block, row key, field) and never tracks an A1 coordinate. Growing
 *      the model physically moves the blocks below it, and every reference
 *      still resolves — so every total is still right afterwards.
 *   3. The output is a real file, not a transcript. It reopens with structure
 *      and live formulas intact.
 */

import {mkdir, readFile} from 'node:fs/promises'
import {inflateRawSync} from 'node:zlib'
import {dirname, join, resolve} from 'node:path'
import {fileURLToPath} from 'node:url'
import {Client} from '@modelcontextprotocol/sdk/client/index.js'
import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js'

const here = dirname(fileURLToPath(import.meta.url))
const repo = resolve(here, '..')
const outDir = join(repo, 'out')
const outFile = join(outDir, 'revenue-model.xlsx')

// ── plumbing ────────────────────────────────────────────────────────────────

let client
let failures = 0

/** Call a tool, print it the way a host transcript would, return parsed data. */
async function call(name, args = {}) {
    const res = await client.callTool({name, arguments: args})
    const text = res.content
        .filter((c) => c.type === 'text')
        .map((c) => c.text)
        .join('\n')
    if (res.isError) {
        throw new Error(`${name} failed: ${text}`)
    }
    process.stdout.write(`  \x1b[2m→ ${name}\x1b[0m ${summarize(args)}\n`)
    // Handlers return a human line then the JSON payload; take the last line
    // that parses as JSON.
    for (const line of text.split('\n').reverse()) {
        try {
            return JSON.parse(line)
        } catch {
            /* not the payload */
        }
    }
    return undefined
}

function summarize(args) {
    const s = JSON.stringify(args)
    return s === '{}' ? '' : `\x1b[2m${s.length > 76 ? `${s.slice(0, 73)}...` : s}\x1b[0m`
}

function step(n, title) {
    process.stdout.write(`\n\x1b[1m${n}. ${title}\x1b[0m\n`)
}

/** Assert, but keep going so one demo run reports everything. */
function check(label, actual, expected) {
    const ok = JSON.stringify(actual) === JSON.stringify(expected)
    if (!ok) failures++
    const mark = ok ? '\x1b[32m✓\x1b[0m' : '\x1b[31m✗\x1b[0m'
    const detail = ok
        ? `\x1b[2m${JSON.stringify(actual)}\x1b[0m`
        : `\x1b[31mgot ${JSON.stringify(actual)}, expected ${JSON.stringify(expected)}\x1b[0m`
    process.stdout.write(`  ${mark} ${label} ${detail}\n`)
}

/** Every row of a block as {key: {field: value}}. */
async function rows(block) {
    const d = await call('describe_block', {name: block, include_rows: true})
    return new Map((d.rows ?? []).map((r) => [r.key, r.values]))
}


/**
 * Pull one entry out of a zip. An .xlsx IS a zip, and the claim being checked
 * below is about what Excel will find in the file — so read the bytes rather
 * than asking the server that just wrote them.
 *
 * Walks local file headers, which is enough for archives this engine writes
 * (no data descriptors, sizes present in the header).
 */
function unzipEntry(buf, wanted) {
    let i = 0
    while (i + 30 <= buf.length && buf.readUInt32LE(i) === 0x04034b50) {
        const method = buf.readUInt16LE(i + 8)
        const compressedSize = buf.readUInt32LE(i + 18)
        const nameLen = buf.readUInt16LE(i + 26)
        const extraLen = buf.readUInt16LE(i + 28)
        const name = buf.subarray(i + 30, i + 30 + nameLen).toString('utf8')
        const start = i + 30 + nameLen + extraLen
        const data = buf.subarray(start, start + compressedSize)
        if (name === wanted) {
            return (method === 0 ? data : inflateRawSync(data)).toString('utf8')
        }
        i = start + compressedSize
    }
    return undefined
}

/** Every worksheet part in the workbook, by entry name. */
function worksheetParts(buf) {
    const parts = []
    for (let n = 1; n <= 8; n++) {
        const xml = unzipEntry(buf, `xl/worksheets/sheet${n}.xml`)
        if (xml !== undefined) parts.push(xml)
    }
    return parts
}

// ── the demo ────────────────────────────────────────────────────────────────

async function main() {
    await mkdir(outDir, {recursive: true})

    const transport = new StdioClientTransport({
        command: process.execPath,
        args: [join(repo, 'dist', 'cli.js')],
        // Keep the server's progress lines off this transcript.
        stderr: 'ignore',
    })
    client = new Client({name: 'logisheets-demo', version: '0.1.0'})
    await client.connect(transport)

    process.stdout.write(
        '\x1b[1mA revenue model an agent builds, computes, and edits\x1b[0m\n' +
            '\x1b[2mReal MCP over stdio against dist/cli.js. No LLM: the engine is the subject.\x1b[0m\n'
    )

    step(1, 'Open a workbook and a structured block')
    await call('open_workbook')
    // fields[0] is the row key. The agent names the block; from here on it
    // addresses data by that name, never by cell coordinates.
    await call('create_block', {
        sheet: 'Model',
        name: 'revenue',
        position: {row: 0, col: 0},
        fields: [
            {name: 'year'},
            {name: 'units', field_type: 'number'},
            {name: 'price', field_type: 'number'},
            {name: 'revenue', field_type: 'number'},
        ],
        initial_rows: [
            {key: '2023', values: {units: 1200, price: 8.5}},
            {key: '2024', values: {units: 1500, price: 9.0}},
        ],
    })

    step(2, 'Give the engine the arithmetic, as a rule')
    // A rule is a template, not a per-row write: it applies to every row,
    // including rows that do not exist yet. This is the whole trick — the agent
    // states the relationship once and stops doing arithmetic.
    await call('set_field_rule', {
        block: 'revenue',
        field: 'revenue',
        value_formula: '=#FIELD("units")*#FIELD("price")',
    })

    const built = await rows('revenue')
    check('2023 revenue = 1200 × 8.5', built.get('2023')?.revenue, 10200)
    check('2024 revenue = 1500 × 9', built.get('2024')?.revenue, 13500)

    step(3, 'A total that reads the block by name, not by range')
    // BLOCKREFS(block, keyCondition, fieldCondition) scans a block: "*" takes
    // every row, and the third argument picks the field. Note what is absent:
    // no A1:A3, so nothing depends on where the block sits or how tall it is.
    await call('create_block', {
        sheet: 'Model',
        name: 'summary',
        position: {row: 6, col: 0},
        fields: [{name: 'metric'}, {name: 'value', field_type: 'number'}],
        initial_rows: [{key: 'total_revenue'}, {key: 'best_year'}],
    })
    await call('set_block_cells', {
        changes: [
            {
                block: 'summary',
                row_key: 'total_revenue',
                field: 'value',
                value: '=SUM(BLOCKREFS("revenue","*","revenue"))',
            },
            {
                block: 'summary',
                row_key: 'best_year',
                field: 'value',
                value: '=MAX(BLOCKREFS("revenue","*","revenue"))',
            },
        ],
    })
    const summary = await rows('summary')
    check('total = 10200 + 13500', summary.get('total_revenue')?.value, 23700)
    check('best year revenue', summary.get('best_year')?.value, 13500)

    step(4, 'The part a raw grid cannot do: grow the model underneath a formula')
    // Adding a year to `revenue` inserts sheet rows, so the summary block below
    // it physically MOVES — its cells are not where they were a moment ago. An
    // agent tracking A1 coordinates is now wrong about every one of them. This
    // agent never learned any, and the engine moves the references with the
    // cells, so the total is simply right afterwards.
    const before = await call('describe_block', {name: 'summary'})
    await call('add_block_rows', {
        block: 'revenue',
        rows: [{key: '2025', values: {units: 1800, price: 9.5}}],
    })
    const after = await call('describe_block', {name: 'summary'})

    const grown = await rows('revenue')
    check('the new row computed on arrival', grown.get('2025')?.revenue, 17100)
    check(
        `summary moved down the sheet (row ${before.position.row} -> ${after.position.row})`,
        after.position.row > before.position.row,
        true
    )
    const afterInsert = await rows('summary')
    check(
        'and its total picked up the new row (23700 + 17100)',
        afterInsert.get('total_revenue')?.value,
        40800
    )
    check('as did the max', afterInsert.get('best_year')?.value, 17100)

    step(5, 'Change one input; every dependent recalculates')
    await call('set_block_cells', {
        changes: [{block: 'revenue', row_key: '2025', field: 'units', value: 2000}],
    })
    const repriced = await rows('revenue')
    check('2025 revenue = 2000 × 9.5', repriced.get('2025')?.revenue, 19000)
    const retotalled = await rows('summary')
    check('total followed', retotalled.get('total_revenue')?.value, 42700)
    check('best year followed', retotalled.get('best_year')?.value, 19000)

    step(6, 'Hand the human a real .xlsx')
    const saved = await call('save_workbook', {path: outFile})
    process.stdout.write(`  \x1b[2m${saved.bytes} bytes\x1b[0m\n`)

    // Reopen it: this is the claim that the file is the deliverable, not a
    // rendering of one. Structure, rules and formulas all have to come back.
    await call('open_workbook', {path: outFile})
    const reopened = await rows('revenue')
    check('values survived the file', reopened.get('2025')?.revenue, 19000)

    const described = await call('describe_block', {name: 'revenue'})
    const rule = described.fields.find((f) => f.name === 'revenue')?.value_formula
    check('the RULE survived, not just its output', rule, '#FIELD("units")*#FIELD("price")')

    // And it is still live after the round trip: a row added to the reopened
    // file computes from the reloaded rule.
    await call('add_block_rows', {
        block: 'revenue',
        rows: [{key: '2026', values: {units: 2200, price: 10}}],
    })
    const stillLive = await rows('revenue')
    check('row added after reload computes', stillLive.get('2026')?.revenue, 22000)

    // Formulas, not baked numbers — that is the difference between a live model
    // and a screenshot of one. Checked twice: what the reopened workbook reports,
    // and what the file itself contains. The second is the one that settles it,
    // since the claim is about what Excel will see rather than what the server
    // says it wrote.
    const summaryAt = await call('describe_block', {name: 'summary'})
    const reread = await call('get_cells', {
        sheetIdx: summaryAt.sheet_idx,
        startRow: summaryAt.position.row,
        startCol: summaryAt.position.col,
        endRow: summaryAt.position.row + summaryAt.row_count - 1,
        endCol: summaryAt.position.col + summaryAt.col_count - 1,
    })
    const live = reread.cells.filter((c) => c.formula?.includes('BLOCKREFS'))
    check(
        'the reopened workbook still knows both totals are formulas',
        live.length,
        2
    )

    const xlsx = await readFile(outFile)
    const sheets = worksheetParts(xlsx)
    const formulas = sheets
        .flatMap((xml) => [...xml.matchAll(/<f>(.*?)<\/f>/g)])
        .map((m) => m[1])

    check('the sheet stores formulas', formulas.length > 0, true)
    check(
        'one per model row, computing revenue',
        formulas.filter((f) => /^\(B\d+\) \* \(C\d+\)$/.test(f)).length,
        3
    )
    check(
        'and the totals still reference the block by name',
        formulas.filter((f) => f.includes('BLOCKREFS')).length,
        2
    )

    process.stdout.write(
        `\n\x1b[1m${failures === 0 ? '\x1b[32mAll checks passed' : `\x1b[31m${failures} check(s) failed`}\x1b[0m\n` +
            `\nOpen it in Excel: \x1b[4m${outFile}\x1b[0m\n` +
            '\x1b[2mColumn D holds a formula per row, and the totals hold BLOCKREFS formulas.\n' +
            'Change units or price in Excel and it recomputes — the model is live,\n' +
            'which a pasted table of numbers would not be.\x1b[0m\n'
    )

    await client.close()
    process.exitCode = failures === 0 ? 0 : 1
}

main().catch(async (err) => {
    process.stderr.write(`\n\x1b[31mdemo failed:\x1b[0m ${err.message}\n`)
    try {
        await client?.close()
    } catch {
        /* already gone */
    }
    process.exitCode = 1
})
