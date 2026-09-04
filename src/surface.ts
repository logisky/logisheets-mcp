/**
 * Which tools this server exposes.
 *
 * logician ships ~55 tools, built for an in-app assistant with a UI. Handing an
 * agent all of them is a real cost: tool-selection accuracy falls as the list
 * grows, and every description is context the agent pays for on every turn. So
 * the default is a deliberate core — the loop from the design doc and nothing
 * else — with the rest available behind an env flag.
 *
 *   LOGISHEETS_MCP_TOOLS=core   (default) the 26 below
 *   LOGISHEETS_MCP_TOOLS=full   everything except the browser-only tools
 */

import {
    BLOCK_OPS_TOOLS,
    BUILDER_TOOLS,
    CELL_TOOLS,
    CHART_TOOLS,
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
 * Namespaces whose tools keep their prefix in the model-facing name.
 *
 * This server drops logician's namespace (`build__create_block` becomes
 * `create_block`) because MCP hosts already prefix every tool by server, so the
 * namespace is context spent twice. That holds while the bare name still says
 * what the tool does — and it stops holding at charts, which logician names
 * `list`, `insert`, `update`, `delete` and `suggest` inside their namespace.
 * Those would sit in the same flat list as `list_blocks`, `insert_rows` and
 * `delete_rows`, and a bare `delete` next to those is a coin flip for exactly
 * the reason this file exists. They keep the prefix.
 */
const PREFIXED_NAMESPACES: ReadonlySet<string> = new Set(['chart'])

/**
 * The name a tool goes on the wire under — what the model sees and calls.
 *
 * Callers must key their tool map by this, not by `tool.name`, or a call comes
 * back for a name nothing is registered under.
 */
export function mcpName(t: Tool): string {
    return PREFIXED_NAMESPACES.has(t.namespace)
        ? `${t.namespace}_${t.name}`
        : t.name
}

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
    // The counterpart for a workbook someone hands you: `create_block` refuses
    // to write over existing data, and this takes data that is already there and
    // makes it addressable in place. Without it an agent given a legacy file has
    // only the destructive half of the pair.
    'build__convert_to_block',
    // The schema says what shape the records are; only this says what they
    // mean. `create_block` takes a description inline, so this is for the case
    // that inline cannot cover: a table adopted with `convert_to_block`, or a
    // block whose purpose is only clear once it has been built.
    'build__set_block_description',
    'build__add_block_rows',
    'build__delete_block_rows',
    // Row order is presentation, not model — but the presentation is part of
    // the deliverable. A person can drag rows around in the app; without this
    // an agent handed the same file cannot, and cannot put a table into the
    // order someone asked for.
    'build__move_block_row',
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
    // Reverse the model: what input lands the answer on a given number. Runs the
    // whole search on the temp branch inside one call — as a conversation it is
    // one round trip per bisection step, and it changes nothing either way.
    'edit__goal_seek',
    // Charts. A chart stores references, not numbers, so it is part of the
    // model rather than a picture of it: the .xlsx the human opens recomputes
    // its chart when they change an assumption, exactly as its formulas do.
    // That is the whole reason charting belongs next to the engine and not in
    // an image the model draws.
    //
    // `from_block` first because it is the idiom this server pushes — name the
    // fields, and the chart follows rows being added and columns moving, the
    // same guarantee blocks give formulas. `insert` is its raw-range
    // counterpart, for the same data that would use `get_cells` / `set_cells`.
    'chart__from_block',
    'chart__insert',
    // `list` is orientation — the chart half of `list_blocks`, and the only
    // source of the chart ids the next two take. Then the pair that makes a
    // chart correctable: `update` for a chart pointing at the wrong series or
    // drawn as the wrong type, `delete` for one on the wrong sheet, which
    // `update` cannot move. Without them a first guess is permanent in the
    // file the human is handed.
    'chart__list',
    'chart__update',
    'chart__delete',
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
        ...CHART_TOOLS,
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
