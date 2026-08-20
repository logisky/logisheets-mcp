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
    /** Tail of the serialization chain — see {@link run}. */
    private lane: Promise<unknown> = Promise.resolve()
    /** Where {@link saveTo} last wrote, so the workbook resource can be named
     *  after the file the human is actually being handed. */
    private lastSaved: string | undefined

    /**
     * Run `fn` after everything already queued on this session, and before
     * anything queued later. One workbook, one lane.
     *
     * Tool handlers are read-then-write against shared state: `create_block`
     * asks whether its sheet exists and creates it if not; it asks the engine
     * for a free block id and then claims it. Those are two awaits with a gap
     * in between, and MCP does not promise to serialize requests — JSON-RPC
     * allows pipelining and the SDK dispatches concurrently. Three
     * `create_block` calls in flight at once therefore each saw the sheet
     * missing, each tried to create it, and two failed.
     *
     * Serializing costs nothing here. The engine is a single synchronous WASM
     * instance, so concurrent handlers never bought throughput — they only
     * interleaved. Reads are queued too: a read overlapping a half-applied
     * transaction would report state that never existed.
     *
     * Not reentrant. Wrap once, at the dispatch boundary; a handler that called
     * back into `run` would wait on itself forever.
     */
    public run<T>(fn: () => Promise<T>): Promise<T> {
        // Run `fn` whether or not its predecessor settled cleanly, then keep the
        // lane resolved so one failed tool call can't poison the queue.
        const result = this.lane.then(fn, fn)
        this.lane = result.then(
            () => undefined,
            () => undefined
        )
        return result
    }

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

    /**
     * The file this workbook belongs to: where it was last saved, else where it
     * was opened from. Undefined for a scratch workbook that has never been
     * written.
     */
    public get path(): string | undefined {
        return this.lastSaved ?? this.workbook.path
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
        this.lastSaved = undefined
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
    public async saveTo(
        path?: string,
        opts: {resolveBlockRefs?: boolean} = {}
    ): Promise<SaveResult> {
        // `this.path`, not `workbook.path`: a second bare save should go back to
        // wherever the last one went, even for a workbook that started empty.
        const target = path ?? this.path
        if (target === undefined) {
            throw new Error(
                'no path given, and this workbook was not opened from a file — ' +
                    'pass an explicit path'
            )
        }
        const absolute = resolve(target)
        await this.workbook.saveAs(absolute, '', opts.resolveBlockRefs ?? false)
        this.lastSaved = absolute
        // Size from disk rather than a second `save()` — serializing a whole
        // workbook twice just to report a number is not worth it.
        return {path: absolute, bytes: (await stat(absolute)).size}
    }

    /** Serialize the workbook to base64 `.xlsx` for transports with no shared disk. */
    public exportBase64(resolveBlockRefs = false): {base64: string; bytes: number} {
        const data = this.workbook.save('', resolveBlockRefs)
        return {
            base64: Buffer.from(data).toString('base64'),
            bytes: data.length,
        }
    }

    /** Release every engine resource this session holds. */
    public close(): void {
        this.runtime.closeAll()
        this.active = undefined
        this.lastSaved = undefined
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
