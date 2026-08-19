import {describe, it, expect} from 'vitest'
import type {Tool} from 'logisheets-logician'
import {validateToolInput} from './validate.js'

/** A stand-in tool carrying only the parts the validator reads. */
const tool = (inputSchema: Tool['inputSchema']): Pick<Tool, 'name' | 'inputSchema'> => ({
    name: 'demo',
    inputSchema,
})

describe('validateToolInput', () => {
    it('accepts arguments that satisfy the schema', () => {
        const t = tool({
            properties: {expr: {type: 'string'}, sheet: {type: 'integer'}},
            required: ['expr'],
        })
        expect(validateToolInput(t, {expr: '=1+1'})).toBeUndefined()
        expect(validateToolInput(t, {expr: '=1+1', sheet: 0})).toBeUndefined()
    })

    it('names the missing parameter instead of throwing downstream', () => {
        // The motivating failure: eval_formula reads `expr` unconditionally, so
        // omitting it produced "Cannot read properties of undefined (reading
        // 'startsWith')" — which names neither the tool nor the parameter.
        const t = tool({properties: {expr: {type: 'string'}}, required: ['expr']})
        const msg = validateToolInput(t, {})
        expect(msg).toContain('missing required parameter `expr`')
        expect(msg).not.toContain('startsWith')
    })

    it('points at a stray key that was probably the missing one', () => {
        const t = tool({properties: {expr: {type: 'string'}}, required: ['expr']})
        const msg = validateToolInput(t, {formula: '=1+1'})
        // Both halves matter: what is missing, and where the agent put it.
        expect(msg).toContain('missing required parameter `expr`')
        expect(msg).toContain('you passed `formula`')
    })

    it('suggests the intended name for a near-miss key', () => {
        const t = tool({
            properties: {block: {type: 'string'}, field: {type: 'string'}},
            required: ['block'],
        })
        // `feild` is a typo, not a synonym, so it survives to the hint pass.
        const msg = validateToolInput(t, {feild: 'total'})
        expect(msg).toContain('unknown parameter `feild`')
        expect(msg).toContain('did you mean `field`?')
    })

    it('reports a wrong type with what it got', () => {
        const t = tool({properties: {row: {type: 'integer'}}, required: ['row']})
        const msg = validateToolInput(t, {row: 'first'})
        expect(msg).toContain('`row` must be integer, got "first"')
    })

    it('accepts an integer where a number is declared, but not the reverse', () => {
        expect(
            validateToolInput(tool({properties: {w: {type: 'number'}}}), {w: 3})
        ).toBeUndefined()
        expect(
            validateToolInput(tool({properties: {w: {type: 'integer'}}}), {w: 3.5})
        ).toContain('must be integer')
    })

    it('honours type unions', () => {
        // set_block_cells declares value as string|number|boolean|null.
        const t = tool({
            properties: {value: {type: ['string', 'number', 'boolean', 'null']}},
        })
        for (const value of ['x', 1, true, null]) {
            expect(validateToolInput(t, {value})).toBeUndefined()
        }
        expect(validateToolInput(t, {value: {nested: 1}})).toContain(
            'must be string or number or boolean or null'
        )
    })

    it('checks enum membership and suggests a variant', () => {
        const t = tool({
            properties: {field_type: {type: 'string', enum: ['string', 'number', 'boolean']}},
        })
        expect(validateToolInput(t, {field_type: 'number'})).toBeUndefined()
        const msg = validateToolInput(t, {field_type: 'numeric'})
        expect(msg).toContain('must be one of')
        expect(msg).toContain('did you mean "number"?')
    })

    it('reaches into nested objects and names the path', () => {
        const t = tool({
            properties: {
                position: {
                    type: 'object',
                    properties: {row: {type: 'integer'}, col: {type: 'integer'}},
                    required: ['row', 'col'],
                },
            },
            required: ['position'],
        })
        expect(validateToolInput(t, {position: {row: 0, col: 0}})).toBeUndefined()
        const msg = validateToolInput(t, {position: {row: 0}})
        expect(msg).toContain('missing required parameter `col` on `position`')
    })

    it('reaches into array items and names the index', () => {
        const t = tool({
            properties: {
                changes: {
                    type: 'array',
                    minItems: 1,
                    items: {
                        type: 'object',
                        properties: {row_key: {type: 'string'}, field: {type: 'string'}},
                        required: ['row_key', 'field'],
                    },
                },
            },
            required: ['changes'],
        })
        expect(
            validateToolInput(t, {changes: [{row_key: 'a', field: 'b'}]})
        ).toBeUndefined()

        const msg = validateToolInput(t, {
            changes: [{row_key: 'a', field: 'b'}, {row_key: 'c'}],
        })
        expect(msg).toContain('`changes[1]`')
        expect(msg).toContain('missing required parameter `field`')
    })

    it('enforces item-count and numeric bounds', () => {
        expect(
            validateToolInput(
                tool({properties: {fields: {type: 'array', minItems: 1}}}),
                {fields: []}
            )
        ).toContain('needs at least 1 item(s)')
        expect(
            validateToolInput(
                tool({properties: {row: {type: 'integer', minimum: 0}}}),
                {row: -1}
            )
        ).toContain('must be >= 0')
    })

    it('tolerates an unknown key when nothing else is wrong', () => {
        // A handler may accept extras, so a stray key alone must not reject a
        // call that would have succeeded. It only shows up as a hint next to a
        // real problem (see the near-miss case above).
        const t = tool({properties: {expr: {type: 'string'}}, required: ['expr']})
        expect(validateToolInput(t, {expr: '=1+1', extra: true})).toBeUndefined()
    })

    it('constrains nothing when the schema declares nothing', () => {
        expect(validateToolInput(tool({}), {anything: 1})).toBeUndefined()
        expect(validateToolInput(tool({}), undefined)).toBeUndefined()
    })

    it('rejects a non-object argument payload', () => {
        const t = tool({properties: {expr: {type: 'string'}}, required: ['expr']})
        expect(validateToolInput(t, 'just a string')).toContain(
            'arguments must be an object'
        )
        expect(validateToolInput(t, [1, 2])).toContain('arguments must be an object')
    })

    it('lists every problem at once, so one round trip fixes them all', () => {
        const t = tool({
            properties: {block: {type: 'string'}, row: {type: 'integer'}},
            required: ['block', 'row'],
        })
        const msg = validateToolInput(t, {row: 'x'})
        expect(msg).toContain('missing required parameter `block`')
        expect(msg).toContain('`row` must be integer')
    })
})
