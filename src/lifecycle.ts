/**
 * Workbook lifecycle tools — the one part of the surface logician doesn't
 * already provide.
 *
 * logician's tools operate on an *open* workbook (it was written for the
 * in-app assistant, where the host owns the file). A standalone MCP server has
 * to own opening and handing back the file itself, so those three tools live
 * here. They are ordinary logician `Tool`s that close over the session, which
 * means the MCP adapter treats them identically to every engine tool.
 *
 * File paths are the primary interface: this server runs next to the agent, and
 * a `.xlsx` round-tripped through base64 would cost tens of thousands of tokens
 * in the model's context for no benefit. base64 stays available for hosts with
 * no shared filesystem.
 */

import type {Tool} from 'logisheets-logician'
import type {OpenResult, SaveResult, WorkbookSession} from './session.js'

/**
 * The active workbook, addressable as an MCP resource.
 *
 * A tool result goes into the model's context, so returning the file itself
 * there costs roughly 1.4 KB of text per KB of workbook and teaches the model
 * nothing. The protocol's answer is a resource: the tool returns a *link* — uri,
 * mime type, size — and a host that wants the bytes fetches them with
 * `resources/read`, outside the conversation. That is how the workbook gets
 * handed back without the model paying for it.
 */
export const WORKBOOK_URI = 'workbook://current.xlsx'

/** The OOXML spreadsheet media type, as Excel and every host expect it. */
export const XLSX_MIME =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/**
 * Tools whose result should carry a link to the workbook, so the host can offer
 * the finished file to the human.
 */
export const TOOLS_YIELDING_WORKBOOK: ReadonlySet<string> = new Set([
    'save_workbook',
])

export interface OpenWorkbookInput {
    path?: string
    xlsx_base64?: string
    name?: string
}

function openWorkbook(session: WorkbookSession): Tool<
    OpenWorkbookInput,
    OpenResult
> {
    return {
        namespace: 'workbook',
        name: 'open_workbook',
        description: [
            'Start the workbook you will work in. Call with no arguments for a fresh, empty one; pass `path` to load an existing .xlsx from disk and work on the human\'s real file.',
            '',
            'This replaces whatever workbook the session currently holds, discarding unsaved changes — so call it once at the start, not between steps. You do NOT have to call it at all: an empty workbook appears automatically the moment any other tool touches the session.',
            '',
            'Prefer `path` over `xlsx_base64`. `xlsx_base64` exists for hosts with no shared filesystem and costs enormous context for a file of any size.',
        ].join('\n'),
        mutates: true,
        confirmation: 'always',
        inputSchema: {
            properties: {
                path: {
                    type: 'string',
                    description:
                        'Path to an existing .xlsx to load. Omit for an empty workbook.',
                },
                xlsx_base64: {
                    type: 'string',
                    description:
                        'A .xlsx as base64, for when no shared filesystem exists. Use `path` when you can.',
                },
                name: {
                    type: 'string',
                    description:
                        'File name to report to the engine when loading from base64.',
                },
            },
        },
        handler: async (input) => {
            const result = await session.open({
                path: input.path,
                xlsxBase64: input.xlsx_base64,
                name: input.name,
            })
            const where =
                result.path !== undefined
                    ? ` from ${result.path}`
                    : result.source === 'bytes'
                      ? ' from uploaded bytes'
                      : ''
            return {
                data: result,
                display: `Opened ${result.source} workbook${where}: ${
                    result.sheets.length
                } sheet(s) — ${result.sheets.join(', ')}`,
            }
        },
    }
}

function saveWorkbook(
    session: WorkbookSession
): Tool<{path?: string; resolve_block_refs?: boolean}, SaveResult> {
    return {
        namespace: 'workbook',
        name: 'save_workbook',
        description: [
            'Write the workbook to a real .xlsx file the human can open in Excel. This is how you hand your work back — do it when the task is done.',
            '',
            'Defaults to the path it was opened from, or last saved to; pass `path` to write somewhere else. Values, formulas and the block structure are all saved.',
            '',
            'IMPORTANT — if the human is going to work on this in Excel, pass `resolve_block_refs: true`. Formulas that read blocks are written as BLOCKREF/BLOCKREFS, which only LogiSheets understands: Excel shows the saved numbers but turns those cells into #NAME? the moment it recalculates. Resolving rewrites them as ordinary A1 references so Excel can recompute the model. Leave it off when the file is coming back here — the named form is readable and survives rows moving.',
            '',
            'The result carries a link to the workbook rather than its bytes, so the host can offer the human the file without any of it passing through your context.',
        ].join('\n'),
        // Writes to the filesystem: not a workbook mutation, but very much an
        // effect on the world, so hosts should gate it.
        mutates: true,
        confirmation: 'always',
        inputSchema: {
            properties: {
                path: {
                    type: 'string',
                    description:
                        'Destination .xlsx path. Defaults to the path the workbook was opened from.',
                },
                resolve_block_refs: {
                    type: 'boolean',
                    default: false,
                    description:
                        'Rewrite BLOCKREF/BLOCKREFS as ordinary A1 references. Set this when the human will open the file in Excel; Excel has no BLOCKREF function and would show #NAME? on recalculation. One-way: a resolved file is an export, not a round trip.',
                },
            },
        },
        handler: async (input) => {
            const result = await session.saveTo(input.path, {
                resolveBlockRefs: input.resolve_block_refs === true,
            })
            return {
                data: result,
                display:
                    `Saved ${result.bytes} bytes to ${result.path}` +
                    (input.resolve_block_refs === true
                        ? ' (block formulas resolved to A1 — Excel can recalculate it)'
                        : ''),
            }
        },
    }
}

function exportXlsx(
    session: WorkbookSession
): Tool<Record<string, never>, {base64: string; bytes: number}> {
    return {
        namespace: 'workbook',
        name: 'export_xlsx',
        description: [
            'Return the workbook as base64-encoded .xlsx bytes.',
            '',
            'Last resort. The bytes land in your context and cost roughly 1.4 KB of text per KB of file. Prefer `save_workbook`, which writes a real file and hands the host a link to it — the host can give the human the workbook without any of it passing through you.',
        ].join('\n'),
        mutates: false,
        confirmation: 'never',
        cost: 'expensive',
        inputSchema: {properties: {}},
        handler: async () => {
            const result = session.exportBase64()
            return {
                data: result,
                display: `Exported ${result.bytes} bytes as base64`,
            }
        },
    }
}

/** The lifecycle tools, bound to one session. */
export function createLifecycleTools(session: WorkbookSession): Tool[] {
    return [
        openWorkbook(session) as Tool,
        saveWorkbook(session) as Tool,
        exportXlsx(session) as Tool,
    ]
}
