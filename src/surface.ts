/**
 * Which tools this server exposes.
 *
 * logician ships ~55 tools, built for an in-app assistant with a UI. Handing an
 * agent all of them is a real cost: tool-selection accuracy falls as the list
 * grows, and every description is context the agent pays for on every turn. So
 * the default is a deliberate core — the loop from the design doc and nothing
 * else — with the rest available behind an env flag.
 *
 *   LOGISHEETS_MCP_TOOLS=core   (default) the 17 below
 *   LOGISHEETS_MCP_TOOLS=full   everything except the browser-only tools
 */

import {
    BLOCK_OPS_TOOLS,
    BUILDER_TOOLS,
    CELL_TOOLS,
    COMMENT_TOOLS,
    EDIT_TOOLS,
    FORMAT_TOOLS,
    HISTORY_TOOLS,
    INSPECT_TOOLS,
    LINK_TOOLS,
    STRUCTURE_TOOLS,
    toolId,
} from 'logisheets-logician'
import type {Tool} from 'logisheets-logician'
import {createLifecycleTools} from './lifecycle.js'
import type {WorkbookSession} from './session.js'

export type ToolMode = 'core' | 'full'

/**
 * The core surface, in the order an agent meets it: orient, build structured
 * memory, compute, read back, hand over.
 *
 * Everything here is addressed semantically — `(block, row_key, field)` rather
 * than `C7` — because positional addressing is exactly what agents get wrong.
 * The raw-cell pair is the escape hatch, kept last on purpose.
 */
const CORE_IDS: readonly string[] = [
    // Lifecycle
    'workbook__open_workbook',
    'workbook__save_workbook',
    'workbook__export_xlsx',
    // Orient
    'build__list_blocks',
    'build__describe_block',
    // Compute
    'build__eval_formula',
    // Structured memory
    'build__create_block',
    'build__add_block_rows',
    'build__delete_block_rows',
    'edit__set_block_cells',
    'build__set_field_rule',
    // Sheets
    'build__create_sheet',
    // Raw-cell escape hatch
    'cell__get_cells',
    'cell__set_cells',
    // `set_field_rule` can attach a validation rule, and this is the only way
    // to see what breaks it — without it, validation is write-only and the
    // agent has no way to check its own work.
    'inspect__list_violations',
    // Answering "what would happen if…" without changing anything. Read-only:
    // it runs the edits on the engine's temp branch, reports the whole cascade
    // and discards them. Without it the only way to explore is to mutate and
    // put back, which walks the model somewhere else if anything goes wrong
    // mid-scan — and a sensitivity scan is dozens of probes.
    'edit__preview_changes',
    // Auditing a number and predicting the blast radius of an edit, from the
    // engine's own dependency graph. The alternative is reading every formula
    // in the workbook and parsing it — and that still cannot answer the reverse
    // direction, which is the one you want before changing an assumption.
    'inspect__trace',
]

/**
 * Tools that cannot work here, excluded from `full` as well as `core`.
 *
 * The craft-interaction tools register cell widgets in the browser app; they
 * detect a headless host and return "not available", so exposing them would
 * only spend context to advertise failures. `get_active_selection` reads what
 * the user has selected on a canvas that doesn't exist in an MCP session.
 */
const NEVER_IDS: ReadonlySet<string> = new Set([
    'craft__register_radio_group',
    'craft__register_multi_select_group',
    'craft__register_point_allocator',
    'craft__register_percent_allocator',
    'craft__register_number_slider',
    'craft__clear_interaction',
    'craft__read_selection',
    'inspect__get_active_selection',
])

/** Every logician tool this server is willing to expose, core first. */
function allEngineTools(): Tool[] {
    return [
        ...BUILDER_TOOLS,
        ...EDIT_TOOLS,
        ...CELL_TOOLS,
        ...INSPECT_TOOLS,
        ...STRUCTURE_TOOLS,
        ...FORMAT_TOOLS,
        ...HISTORY_TOOLS,
        ...COMMENT_TOOLS,
        ...BLOCK_OPS_TOOLS,
        ...LINK_TOOLS,
    ]
}

/** Read the mode from the environment, defaulting to `core`. */
export function toolModeFromEnv(
    env: Record<string, string | undefined> = process.env
): ToolMode {
    const raw = env.LOGISHEETS_MCP_TOOLS?.trim().toLowerCase()
    if (raw === undefined || raw === '') return 'core'
    if (raw === 'core' || raw === 'full') return raw
    throw new Error(
        `LOGISHEETS_MCP_TOOLS must be "core" or "full", got "${raw}"`
    )
}

/**
 * Resolve the tool list for a session.
 *
 * Core tools come first and in the declared order, so a host that renders the
 * list in order shows the agent the loop rather than an alphabet soup.
 */
export function selectTools(session: WorkbookSession, mode: ToolMode): Tool[] {
    const available = new Map<string, Tool>()
    for (const t of [...createLifecycleTools(session), ...allEngineTools()]) {
        const id = toolId(t)
        if (available.has(id)) {
            throw new Error(`duplicate tool id from logician: ${id}`)
        }
        available.set(id, t)
    }

    const core: Tool[] = CORE_IDS.map((id) => {
        const t = available.get(id)
        if (t === undefined) {
            // A rename upstream must fail loudly at startup rather than
            // silently shrink the agent's surface.
            throw new Error(
                `core tool "${id}" not found — logisheets-logician may have renamed it`
            )
        }
        return t
    })

    if (mode === 'core') return core

    const coreIds = new Set(CORE_IDS)
    const rest = [...available.entries()]
        .filter(([id]) => !coreIds.has(id) && !NEVER_IDS.has(id))
        .map(([, t]) => t)
    return [...core, ...rest]
}
