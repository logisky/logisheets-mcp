#!/usr/bin/env node
/**
 * stdio entry point — what an MCP host (Claude Desktop, Cursor, Cline) spawns.
 *
 * stdout is the JSON-RPC channel and must carry nothing else, so every
 * diagnostic in this process goes to stderr.
 */

import {StdioServerTransport} from '@modelcontextprotocol/sdk/server/stdio.js'
import {createServer, SERVER_VERSION} from './server.js'
import {toolModeFromEnv} from './surface.js'

async function main(): Promise<void> {
    const argv = process.argv.slice(2)
    if (argv.includes('--help') || argv.includes('-h')) {
        process.stderr.write(
            [
                'logisheets-mcp — a real spreadsheet engine as an MCP server (stdio).',
                '',
                'Environment:',
                '  LOGISHEETS_MCP_TOOLS=core|full   tool surface (default: core)',
                '',
                'Configure your MCP host to run this command; it speaks MCP on stdio,',
                'not something you interact with directly in a terminal.',
                '',
            ].join('\n')
        )
        return
    }
    if (argv.includes('--version') || argv.includes('-v')) {
        process.stderr.write(`${SERVER_VERSION}\n`)
        return
    }

    const {server, session, tools} = createServer()
    process.stderr.write(
        `logisheets-mcp ${SERVER_VERSION}: ${tools.size} tools (${toolModeFromEnv()} surface)\n`
    )

    const shutdown = (): void => {
        session.close()
        process.exit(0)
    }
    process.on('SIGINT', shutdown)
    process.on('SIGTERM', shutdown)

    await server.connect(new StdioServerTransport())
}

main().catch((err: unknown) => {
    process.stderr.write(
        `logisheets-mcp failed to start: ${
            err instanceof Error ? (err.stack ?? err.message) : String(err)
        }\n`
    )
    process.exit(1)
})
