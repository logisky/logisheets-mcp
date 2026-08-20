/**
 * The registry manifest duplicates facts that live in package.json — the server
 * name, the package identifier, and the version in two places. Nothing keeps
 * them in step, and a mismatch fails at publish time with an ownership error
 * rather than anywhere useful, so check it here.
 */

import {readFile} from 'node:fs/promises'
import {dirname, join, resolve} from 'node:path'
import {fileURLToPath} from 'node:url'
import {describe, it, expect} from 'vitest'

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..')

async function readJson(name: string): Promise<Record<string, unknown>> {
    return JSON.parse(await readFile(join(repo, name), 'utf8')) as Record<
        string,
        unknown
    >
}

describe('registry manifest', () => {
    it('agrees with package.json', async () => {
        const pkg = await readJson('package.json')
        const server = (await readJson('server.json')) as {
            name: string
            version: string
            packages: Array<{
                identifier: string
                version: string
                registryType: string
                registryBaseUrl: string
            }>
        }

        // npm ownership is proved by matching `mcpName` against the server name.
        expect(pkg.mcpName).toBe(server.name)
        expect(server.version).toBe(pkg.version)

        const npm = server.packages[0]
        expect(npm).toBeDefined()
        if (npm === undefined) return
        expect(npm.identifier).toBe(pkg.name)
        expect(npm.version).toBe(pkg.version)
        // Only the public registry is accepted.
        expect(npm.registryBaseUrl).toBe('https://registry.npmjs.org')
        expect(npm.registryType).toBe('npm')
    })

    it('declares the environment variables the server actually reads', async () => {
        const server = (await readJson('server.json')) as {
            packages: Array<{
                environmentVariables?: Array<{name: string}>
            }>
        }
        const declared = (server.packages[0]?.environmentVariables ?? []).map(
            (e) => e.name
        )
        // Anything the server branches on should be discoverable from the
        // manifest; `surface.ts` is the only place that reads the environment.
        const source = await readFile(join(repo, 'src', 'surface.ts'), 'utf8')
        const read = [...source.matchAll(/process\.env\.(\w+)/g)].map(
            (m) => m[1] as string
        )
        for (const name of new Set(read)) {
            expect(declared).toContain(name)
        }
    })
})
