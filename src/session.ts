/**
 * The session's workbook — the thing that makes this server *memory* rather
 * than a calculator.
 *
 * One MCP session owns one active workbook, held here and persistent across
 * tool calls. Every tool operates on it, so state the agent built in step 3 is
 * still there in step 30. (Multiple named workbooks per session is a later
 * feature; a single active one keeps the tool surface small and the agent's
 * mental model simple.)
 */

import {resolve} from 'node:path'
import {stat} from 'node:fs/promises'
import {SpreadsheetRuntime, Workbook} from 'logisheets-runtime'

/** The async engine client, as logician's tool handlers consume it. */
export type WorkbookClient = Workbook['client']

export interface OpenResult {
    /** Where the workbook came from. */
    source: 'new' | 'file' | 'bytes'
    /** Absolute path, when opened from (or destined for) a file. */
    path?: string
    /** Sheet names in order. */
    sheets: string[]
}

export interface SaveResult {
    path: string
    bytes: number
}

export class WorkbookSession {
    private readonly runtime = new SpreadsheetRuntime()
    private active: Workbook | undefined

    /**
     * The active workbook, created empty on first touch.
     *
     * Lazy creation is deliberate: an agent that dives straight into
     * `create_block` shouldn't fail because it skipped `open_workbook`. The
     * scratchpad simply exists as soon as anything reaches for it.
     */
    public get workbook(): Workbook {
        if (this.active === undefined) {
            this.active = this.runtime.createWorkbook()
        }
        return this.active
    }

    public get client(): WorkbookClient {
        return this.workbook.client
    }

    /** Path the active workbook was opened from, if any. */
    public get path(): string | undefined {
        return this.workbook.path
    }

    /** True once a workbook exists — i.e. anything has touched the session. */
    public get isOpen(): boolean {
        return this.active !== undefined
    }

    /**
     * Replace the active workbook.
     *
     * `path` loads a real `.xlsx` from disk; `xlsxBase64` loads one from bytes
     * (for hosts with no shared filesystem); neither starts an empty workbook.
     */
    public async open(opts: {
        path?: string
        xlsxBase64?: string
        name?: string
    }): Promise<OpenResult> {
        if (opts.path !== undefined && opts.xlsxBase64 !== undefined) {
            throw new Error('pass either path or xlsx_base64, not both')
        }

        let next: Workbook
        let source: OpenResult['source']
        if (opts.path !== undefined) {
            next = await this.runtime.loadWorkbook(opts.path)
            source = 'file'
        } else if (opts.xlsxBase64 !== undefined) {
            const bytes = Buffer.from(opts.xlsxBase64, 'base64')
            if (bytes.length === 0) {
                throw new Error('xlsx_base64 decoded to zero bytes')
            }
            next = this.runtime.loadWorkbookFromBytes(
                bytes,
                opts.name ?? 'workbook.xlsx'
            )
            source = 'bytes'
        } else {
            next = this.runtime.createWorkbook()
            source = 'new'
        }

        // Release the workbook being displaced. The runtime dedups loads by
        // path, so re-opening the current file hands back the same handle —
        // closing it then would release the workbook we just "opened".
        const previous = this.active
        this.active = next
        if (previous !== undefined && previous !== next) {
            this.runtime.close(previous)
        }

        return {source, path: next.path, sheets: await this.sheetNames()}
    }

    /** Sheet names of the active workbook, in order. */
    public async sheetNames(): Promise<string[]> {
        const infos = await this.workbook.client.getAllSheetInfo()
        if (isErrorMessage(infos)) {
            throw new Error(`getAllSheetInfo failed: ${infos.msg}`)
        }
        return infos.map((s) => s.name)
    }

    /**
     * Write the workbook to `path`, defaulting to where it was opened from.
     * Returns the absolute path actually written and the file size.
     */
    public async saveTo(path?: string): Promise<SaveResult> {
        const target = path ?? this.workbook.path
        if (target === undefined) {
            throw new Error(
                'no path given, and this workbook was not opened from a file — ' +
                    'pass an explicit path'
            )
        }
        const absolute = resolve(target)
        await this.workbook.saveAs(absolute)
        // Size from disk rather than a second `save()` — serializing a whole
        // workbook twice just to report a number is not worth it.
        return {path: absolute, bytes: (await stat(absolute)).size}
    }

    /** Serialize the workbook to base64 `.xlsx` for transports with no shared disk. */
    public exportBase64(): {base64: string; bytes: number} {
        const data = this.workbook.save()
        return {
            base64: Buffer.from(data).toString('base64'),
            bytes: data.length,
        }
    }

    /** Release every engine resource this session holds. */
    public close(): void {
        this.runtime.closeAll()
        this.active = undefined
    }
}

/** Local copy of the engine's error-shape guard (avoids a direct engine dep). */
function isErrorMessage(v: unknown): v is {msg: string; ty: number} {
    return (
        typeof v === 'object' &&
        v !== null &&
        'msg' in v &&
        typeof (v as {msg: unknown}).msg === 'string'
    )
}
