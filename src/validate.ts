/**
 * Check tool arguments against the tool's own declared JSON Schema, before the
 * handler ever sees them.
 *
 * MCP puts each tool's `inputSchema` on the wire, but nothing enforces it: the
 * SDK hands `params.arguments` straight through. Handlers then read fields that
 * aren't there, and the agent gets whatever TypeError falls out —
 * `Cannot read properties of undefined (reading 'startsWith')` for a missing
 * `expr`. That names neither the tool nor the parameter, so the agent has no
 * way to correct itself and burns turns guessing.
 *
 * Agents mostly get arguments wrong in a few predictable ways: they omit a
 * required parameter, invent a plausible synonym for its name (`formula` for
 * `expr`), pass a string where a number belongs, or guess an enum variant. So
 * the messages here name the parameter, say what was expected, and — the part
 * that actually saves a turn — suggest the declared name a stray key looks like.
 *
 * Deliberately a subset of draft-07: the keywords logician's schemas actually
 * use (type incl. unions, required, properties, items, enum, bounds). No $ref,
 * no anyOf/allOf. An unrecognized keyword is ignored rather than guessed at.
 */

import type {JSONSchema, JSONSchemaType, Tool} from 'logisheets-logician'

/** Levenshtein distance, iterative two-row. Only used on short key names. */
function editDistance(a: string, b: string): number {
    if (a === b) return 0
    let prev = Int32Array.from({length: b.length + 1}, (_, i) => i)
    let row = new Int32Array(b.length + 1)
    // Every index below is in bounds by construction; `?? 0` is only there to
    // satisfy noUncheckedIndexedAccess, which applies to typed arrays too.
    for (let i = 1; i <= a.length; i++) {
        row[0] = i
        for (let j = 1; j <= b.length; j++) {
            const drop = (prev[j] ?? 0) + 1
            const add = (row[j - 1] ?? 0) + 1
            const swap = (prev[j - 1] ?? 0) + (a[i - 1] === b[j - 1] ? 0 : 1)
            row[j] = Math.min(drop, add, swap)
        }
        // Swap the rows instead of reallocating; the old `prev` becomes scratch.
        const done = row
        row = prev
        prev = done
    }
    return prev[b.length] ?? 0
}

/**
 * The declared name `key` was probably meant to be, or undefined. Case-blind
 * exact match first (`Expr` → `expr`), then near-misses, with the threshold
 * scaled to length so short names don't match everything.
 */
function didYouMean(
    key: string,
    candidates: readonly string[]
): string | undefined {
    const lower = key.toLowerCase()
    const exact = candidates.find((c) => c.toLowerCase() === lower)
    if (exact !== undefined) return exact

    let best: string | undefined
    let bestDist = Infinity
    for (const c of candidates) {
        const d = editDistance(lower, c.toLowerCase())
        if (d < bestDist) {
            bestDist = d
            best = c
        }
    }
    if (best === undefined) return undefined
    const maxLen = Math.max(key.length, best.length)
    const limit = Math.min(3, Math.max(1, Math.ceil(maxLen / 3)))
    return bestDist <= limit ? best : undefined
}

/** The JSON type name for a value, in the vocabulary schemas use. */
function jsonTypeOf(v: unknown): JSONSchemaType {
    if (v === null) return 'null'
    if (Array.isArray(v)) return 'array'
    if (typeof v === 'number') return Number.isInteger(v) ? 'integer' : 'number'
    if (typeof v === 'string') return 'string'
    if (typeof v === 'boolean') return 'boolean'
    return 'object'
}

/** Does `v` satisfy a single declared type? `integer` implies `number`. */
function matchesType(v: unknown, want: JSONSchemaType): boolean {
    const got = jsonTypeOf(v)
    if (got === want) return true
    // An integer is a valid number; a whole-valued float is a valid integer.
    if (want === 'number' && got === 'integer') return true
    return false
}

function describe(v: unknown): string {
    if (typeof v === 'string') return JSON.stringify(v)
    if (v === null || v === undefined) return String(v)
    if (Array.isArray(v)) return `an array of ${v.length}`
    if (typeof v === 'object') return 'an object'
    return String(v)
}

/** Human-readable form of a schema's declared type(s). */
function typeNames(t: JSONSchemaType | readonly JSONSchemaType[]): string {
    return Array.isArray(t) ? t.join(' or ') : String(t)
}

/**
 * Collect every problem with `value` against `schema`. `path` is the argument
 * name as the agent wrote it (`changes[0].row_key`), so a message points at the
 * exact spot in a nested payload.
 */
function collect(
    value: unknown,
    schema: JSONSchema,
    path: string,
    out: string[]
): void {
    const label = path === '' ? 'argument' : `\`${path}\``

    if (schema.type !== undefined) {
        const wanted = Array.isArray(schema.type) ? schema.type : [schema.type]
        if (!wanted.some((w) => matchesType(value, w))) {
            out.push(
                `${label} must be ${typeNames(schema.type)}, got ${describe(value)}`
            )
            // The type is wrong, so every nested check below would just be
            // noise about a value that has to be replaced wholesale.
            return
        }
    }

    if (schema.enum !== undefined && Array.isArray(schema.enum)) {
        if (!schema.enum.includes(value as string | number)) {
            const allowed = schema.enum.map((e) => JSON.stringify(e)).join(', ')
            const hint =
                typeof value === 'string'
                    ? didYouMean(
                          value,
                          schema.enum.filter(
                              (e): e is string => typeof e === 'string'
                          )
                      )
                    : undefined
            out.push(
                `${label} must be one of ${allowed}, got ${describe(value)}` +
                    (hint !== undefined ? ` — did you mean ${JSON.stringify(hint)}?` : '')
            )
        }
    }

    if (typeof value === 'number') {
        if (schema.minimum !== undefined && value < schema.minimum) {
            out.push(`${label} must be >= ${schema.minimum}, got ${value}`)
        }
        if (schema.maximum !== undefined && value > schema.maximum) {
            out.push(`${label} must be <= ${schema.maximum}, got ${value}`)
        }
    }

    if (Array.isArray(value)) {
        if (schema.minItems !== undefined && value.length < schema.minItems) {
            out.push(
                `${label} needs at least ${schema.minItems} item(s), got ${value.length}`
            )
        }
        if (schema.maxItems !== undefined && value.length > schema.maxItems) {
            out.push(
                `${label} allows at most ${schema.maxItems} item(s), got ${value.length}`
            )
        }
        if (schema.items !== undefined) {
            for (let i = 0; i < value.length; i++) {
                collect(value[i], schema.items, `${path}[${i}]`, out)
            }
        }
        return
    }

    // Objects: missing required keys, then recurse into declared properties.
    if (
        value !== null &&
        typeof value === 'object' &&
        (schema.properties !== undefined || schema.required !== undefined)
    ) {
        const obj = value as Record<string, unknown>
        const props = schema.properties ?? {}
        const declared = Object.keys(props)
        const unknown = Object.keys(obj).filter((k) => !declared.includes(k))
        const missing = (schema.required ?? []).filter(
            (k) => obj[k] === undefined
        )
        // Stray keys already blamed on a missing parameter; not repeated below.
        const paired = new Set<string>()

        for (const key of missing) {
            // The commonest agent mistake: the value IS there, under a name the
            // agent invented. Point at the stray key rather than just the gap.
            //
            // Two ways to spot it. A lexical near-miss catches typos
            // (`feild`/`field`). But the frequent case is a *synonym* —
            // `formula` for `expr` — which shares no letters, so fall back to
            // structure: exactly one parameter missing and exactly one key the
            // schema doesn't know is almost certainly the same one renamed.
            let stray = unknown.find(
                (k) => !paired.has(k) && didYouMean(k, [key]) === key
            )
            if (
                stray === undefined &&
                missing.length === 1 &&
                unknown.length === 1
            ) {
                const only = unknown[0] as string
                // ...unless that key is plainly a typo of some OTHER declared
                // parameter (`feild` for `field`). Blaming it on this one would
                // send the agent to the wrong place; the unknown-key hint below
                // names the right one.
                const elsewhere = didYouMean(only, declared)
                if (elsewhere === undefined || elsewhere === key) stray = only
            }
            if (stray !== undefined) paired.add(stray)
            const at = path === '' ? '' : ` on \`${path}\``
            out.push(
                stray !== undefined
                    ? `missing required parameter \`${key}\`${at} — you passed \`${stray}\`, did you mean \`${key}\`?`
                    : `missing required parameter \`${key}\`${at}`
            )
        }

        for (const [key, sub] of Object.entries(props)) {
            const v = obj[key]
            if (v === undefined) continue // absent optional; required handled above
            collect(v, sub, path === '' ? key : `${path}.${key}`, out)
        }

        // Unknown keys are a hard error, not a hint.
        //
        // This started out the other way round — reported only alongside some
        // other problem, on the theory that a handler might tolerate extras
        // and that rejecting a call which would have worked is worse than
        // staying quiet. Testing settled it: passing `after_key` to a tool
        // that had no such parameter returned success, the rows went
        // somewhere else entirely, and everything downstream reasoned from a
        // false premise. A rejected call costs one retry; a silently ignored
        // parameter costs the agent its model of what the workbook contains.
        //
        // Free-form objects are unaffected — this whole branch is only
        // entered for schemas that actually declare `properties`/`required`,
        // so a value-bag keyed by arbitrary field names never reaches here.
        if (declared.length > 0) {
            for (const key of unknown) {
                if (paired.has(key)) continue
                const guess = didYouMean(key, declared)
                out.push(
                    guess !== undefined
                        ? `unknown parameter \`${key}\` — did you mean \`${guess}\`?`
                        : `unknown parameter \`${key}\``
                )
            }
        }
    }
}

/**
 * Validate a tool call's arguments. Returns an agent-facing message listing
 * every problem, or undefined when the arguments are acceptable.
 */
export function validateToolInput(
    tool: Pick<Tool, 'name' | 'inputSchema'>,
    args: unknown
): string | undefined {
    // A schema declaring nothing constrains nothing.
    const schema = tool.inputSchema
    if (schema.properties === undefined && schema.required === undefined) {
        return undefined
    }

    if (args !== undefined && (typeof args !== 'object' || args === null || Array.isArray(args))) {
        return `${tool.name}: arguments must be an object, got ${describe(args)}`
    }

    const problems: string[] = []
    collect(args ?? {}, {...schema, type: 'object'}, '', problems)
    if (problems.length === 0) return undefined

    const lines = problems.map((p) => `  - ${p}`).join('\n')
    return `${tool.name}: invalid arguments\n${lines}`
}
