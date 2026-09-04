/**
 * The MCP shell: logician tools in, MCP protocol out.
 *
 * We use the SDK's low-level `Server` rather than `McpServer` on purpose.
 * `McpServer.registerTool` wants Zod schemas, while logician tools already
 * carry hand-written JSON Schema — which is what MCP puts on the wire anyway.
 * Going low-level passes those straight through instead of round-tripping them
 * through a Zod translation layer.
 */

import {Server} from '@modelcontextprotocol/sdk/server/index.js'
import {
    CallToolRequestSchema,
    ListResourcesRequestSchema,
    ListToolsRequestSchema,
    ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js'
import type {
    CallToolResult,
    Tool as McpTool,
} from '@modelcontextprotocol/sdk/types.js'
import type {Tool, ToolContext} from 'logisheets-logician'
import {WorkbookSession} from './session.js'
import {mcpName, selectTools, toolModeFromEnv, type ToolMode} from './surface.js'
import {validateToolInput} from './validate.js'
import {
    TOOLS_YIELDING_WORKBOOK,
    WORKBOOK_URI,
    XLSX_MIME,
} from './lifecycle.js'

export const SERVER_NAME = 'logisheets'
export const SERVER_VERSION = '0.1.0'

/**
 * How the agent should approach this server. Sent as MCP `instructions`, so a
 * host can put it in front of the model before it starts guessing.
 */
export const INSTRUCTIONS = [
    'A real, Excel-compatible spreadsheet engine you can compute in and remember in.',
    '',
    'Three things it is for:',
    '  1. Arithmetic you should not do yourself. Write a formula and let the engine evaluate it — `eval_formula` for a one-off, or store the formula in a cell so it keeps recalculating.',
    '  2. Structured memory that survives the whole task. A *block* is a named table; you address its cells by (block name, row key, field name), never by A1 coordinates. Inserting rows never breaks a reference, so you can keep building without tracking where anything sits.',
    '  3. Charts that stay live. `chart_from_block` plots a block by field name and the chart follows the data — rows added later appear on their own. Never draw a picture of the numbers; make the workbook hold the chart, so it redraws when the human changes an assumption in Excel.',
    '',
    'The loop: `list_blocks` to see what you have, `create_block` to open a structured workspace, `add_block_rows` / `set_block_cells` to fill it, formulas for the math, `describe_block` to read results back, `save_workbook` to hand the human a real .xlsx.',
    '',
    'Prefer blocks over raw cells. `get_cells` / `set_cells` exist for data that genuinely has no structure.',
    '',
    'Write down what a block is for — a sentence on `create_block`, or `set_block_description` afterwards. `describe_block` gives it back, so it is what the next session (yours or someone else\'s) reads instead of guessing the meaning from column names.',
].join('\n')

export interface CreateServerOptions {
    /** Reuse an existing session (tests, embedding). Defaults to a fresh one. */
    session?: WorkbookSession
    /** Override the tool surface; defaults to `LOGISHEETS_MCP_TOOLS` or `core`. */
    mode?: ToolMode
    /**
     * Where tool progress lines go. Defaults to stderr — on the stdio
     * transport, stdout carries JSON-RPC frames and writing anything else
     * there corrupts the stream.
     */
    log?: (msg: string) => void
}

export interface CreatedServer {
    server: Server
    session: WorkbookSession
    /** The tools exposed, by MCP name. */
    tools: Map<string, Tool>
}

/** Translate a logician tool into its MCP wire description. */
function toMcpTool(t: Tool): McpTool {
    // logician declares `required` readonly; the wire type wants it mutable.
    const {required, ...schema} = t.inputSchema
    return {
        // logician namespaces tools (`build__create_block`) to keep crafts from
        // colliding. MCP hosts already prefix by server, so the namespace would
        // just be noise in the model's context — `mcpName` drops it, except
        // where the bare name would be ambiguous. Names are asserted unique
        // below.
        name: mcpName(t),
        description: t.description,
        // Spread first: a tool's own `type` must not shadow 'object', which is
        // what MCP requires at the top level of a tool's input schema.
        inputSchema: {
            ...schema,
            type: 'object',
            ...(required !== undefined ? {required: [...required]} : {}),
        },
        annotations: {
            readOnlyHint: !t.mutates,
            // logician's 'destructive' policy marks the tools that discard data
            // (clear a block, delete rows/sheets) rather than merely write.
            destructiveHint: t.confirmation === 'destructive',
        },
    }
}

/**
 * Build the server. Nothing is started — hand the result to a transport.
 */
export function createServer(opts: CreateServerOptions = {}): CreatedServer {
    const session = opts.session ?? new WorkbookSession()
    const mode = opts.mode ?? toolModeFromEnv()
    const log = opts.log ?? ((msg: string) => process.stderr.write(`${msg}\n`))

    const tools = new Map<string, Tool>()
    for (const t of selectTools(session, mode)) {
        const name = mcpName(t)
        if (tools.has(name)) {
            throw new Error(
                `two tools share the MCP name "${name}" — namespaces differ but names must be unique`
            )
        }
        tools.set(name, t)
    }

    const server = new Server(
        {name: SERVER_NAME, version: SERVER_VERSION},
        {
            capabilities: {tools: {}, resources: {}},
            instructions: INSTRUCTIONS,
        }
    )

    // The workbook as a resource. This is how the finished file reaches the
    // human: `save_workbook` returns a link, and a host that wants the bytes
    // reads them here — outside the model's context, so a 200 KB workbook costs
    // the conversation nothing.
    server.setRequestHandler(ListResourcesRequestSchema, async () => {
        if (!session.isOpen) return {resources: []}
        return {
            resources: [
                {
                    uri: WORKBOOK_URI,
                    name: session.path ?? 'workbook.xlsx',
                    title: 'The active workbook',
                    description:
                        'The workbook this session is working in, serialized as a real .xlsx.',
                    mimeType: XLSX_MIME,
                    // Marked for the human: it is a file to open, not text to
                    // reason over.
                    annotations: {audience: ['user']},
                },
            ],
        }
    })

    server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
        if (request.params.uri !== WORKBOOK_URI) {
            throw new Error(`unknown resource: ${request.params.uri}`)
        }
        // Serialize through the same lane as the tools, so a read can't observe
        // a half-applied transaction.
        const {base64, bytes} = await session.run(async () =>
            session.exportBase64()
        )
        return {
            contents: [
                {
                    uri: WORKBOOK_URI,
                    mimeType: XLSX_MIME,
                    blob: base64,
                    _meta: {bytes},
                },
            ],
        }
    })

    server.setRequestHandler(ListToolsRequestSchema, async () => ({
        tools: [...tools.values()].map(toMcpTool),
    }))

    server.setRequestHandler(
        CallToolRequestSchema,
        async (request, extra): Promise<CallToolResult> => {
            const tool = tools.get(request.params.name)
            if (tool === undefined) {
                return {
                    content: [
                        {type: 'text', text: `unknown tool: ${request.params.name}`},
                    ],
                    isError: true,
                }
            }

            // Check arguments against the tool's declared schema first. Nothing
            // else does — the SDK passes `arguments` through untouched — so a
            // wrong parameter name would otherwise surface as whatever
            // TypeError the handler happens to throw, naming neither the tool
            // nor the parameter the agent got wrong.
            const invalid = validateToolInput(tool, request.params.arguments)
            if (invalid !== undefined) {
                return {
                    content: [{type: 'text', text: invalid}],
                    isError: true,
                }
            }

            const ctx: ToolContext = {
                workbook: session.client,
                signal: extra.signal,
                // The MCP host owns approval: it decided to dispatch this call,
                // and a server-side prompt has nobody to ask. Mutating tools
                // are flagged via annotations so the host can gate them there.
                confirm: async () => true,
                log,
            }

            try {
                // One workbook, one lane. Handlers read state and then write it
                // across an await, and the host may have several calls in
                // flight, so without this they interleave and lose. See
                // WorkbookSession.run.
                const result = await session.run(() =>
                    tool.handler(request.params.arguments ?? {}, ctx)
                )
                if (result.canceled === true) {
                    return {
                        content: [{type: 'text', text: 'canceled'}],
                        isError: true,
                    }
                }
                const content: CallToolResult['content'] = []
                if (result.display !== undefined && result.display !== '') {
                    content.push({type: 'text', text: result.display})
                }
                if (result.data !== undefined) {
                    content.push({
                        type: 'text',
                        text: JSON.stringify(result.data),
                    })
                }
                if (content.length === 0) {
                    content.push({type: 'text', text: 'ok'})
                }
                // Hand back a reference to the file, not the file. The host can
                // turn this into a download for the human; the model just sees a
                // short link.
                if (TOOLS_YIELDING_WORKBOOK.has(tool.name)) {
                    content.push({
                        type: 'resource_link',
                        uri: WORKBOOK_URI,
                        name: session.path ?? 'workbook.xlsx',
                        mimeType: XLSX_MIME,
                        description:
                            'The saved workbook. Read this resource to get the .xlsx bytes.',
                        annotations: {audience: ['user']},
                    })
                }
                return {content}
            } catch (err) {
                // Tool failures are results, not protocol errors: the agent
                // should see the message and get a chance to correct itself.
                return {
                    content: [
                        {
                            type: 'text',
                            text: err instanceof Error ? err.message : String(err),
                        },
                    ],
                    isError: true,
                }
            }
        }
    )

    return {server, session, tools}
}
