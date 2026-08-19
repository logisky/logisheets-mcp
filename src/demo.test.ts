/**
 * The demo is documentation people will actually run, so it has to keep
 * working. Run it as a test: it exits non-zero if any of its own checks fail.
 *
 * It drives the built binary over stdio, so `dist/` must be current — the
 * pretest build handles that.
 */

import {execFile} from 'node:child_process'
import {promisify} from 'node:util'
import {dirname, join, resolve} from 'node:path'
import {fileURLToPath} from 'node:url'
import {describe, it, expect} from 'vitest'

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..')

describe('examples/revenue-model.mjs', () => {
    it('passes every check it makes', async () => {
        const {stdout} = await promisify(execFile)(
            process.execPath,
            [join(repo, 'examples', 'revenue-model.mjs')],
            {cwd: repo}
        )
        // execFile rejects on a non-zero exit, so reaching here means the demo
        // reported success; assert the summary line as well so a demo that
        // stops checking things can't pass by doing nothing.
        expect(stdout).toContain('All checks passed')
        expect(stdout).not.toContain('failed')
    }, 60_000)
})
