/**
 * logisheets-mcp — a real, Excel-compatible spreadsheet engine as an MCP server.
 *
 * Library entry point, for embedding the server in another host (an HTTP
 * transport, an agent framework, a test). The stdio binary lives in ./cli.ts.
 */

export {
    createServer,
    INSTRUCTIONS,
    SERVER_NAME,
    SERVER_VERSION,
    type CreatedServer,
    type CreateServerOptions,
} from './server.js'
export {
    selectTools,
    toolModeFromEnv,
    type ToolMode,
} from './surface.js'
export {
    WorkbookSession,
    type OpenResult,
    type SaveResult,
    type WorkbookClient,
} from './session.js'
export {createLifecycleTools, type OpenWorkbookInput} from './lifecycle.js'
