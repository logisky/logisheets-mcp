#!/usr/bin/env node
//
// Gate a release on the four identities lining up. Run it before anything is
// published, because `npm publish` cannot be taken back: a version number is
// spent the moment it lands, so a mismatch found afterwards means burning the
// next one.
//
//   node scripts/check-release.mjs            # check the files agree
//   node scripts/check-release.mjs 0.2.0      # ...and agree with this version
//
// What can drift, and has:
//   * server.json's `version` and its `packages[0].version` are separate
//     fields that both have to track package.json.
//   * the registry cross-checks `mcpName` in the PUBLISHED package against
//     server.json's `name`; a rename in one place fails at publish time.
//   * server.json is schema-constrained (description is capped at 100
//     characters) and nothing in a normal build ever reads it.

import {readFileSync} from 'node:fs'
import {dirname, join} from 'node:path'
import {fileURLToPath} from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const read = (name) => JSON.parse(readFileSync(join(root, name), 'utf8'))
const pkg = read('package.json')
const server = read('server.json')

/** Version to hold everything to. From the tag in CI, else package.json. */
const expected = (process.argv[2] ?? pkg.version).replace(/^v/, '')

const problems = []
const eq = (label, got, want) => {
    if (got !== want) problems.push(`${label}: ${got ?? '(missing)'} != ${want}`)
}

const npmPackage = server.packages?.[0]
if (npmPackage === undefined) {
    problems.push('server.json has no packages[0]')
} else {
    eq('server.json packages[0].version', npmPackage.version, expected)
    eq('server.json packages[0].identifier', npmPackage.identifier, pkg.name)
    eq('server.json packages[0].registryType', npmPackage.registryType, 'npm')
}
eq('package.json version', pkg.version, expected)
eq('server.json version', server.version, expected)
eq('package.json mcpName', pkg.mcpName, server.name)

// Schema limits worth failing on here rather than at the registry. The full
// schema is checked by the registry itself; these are the ones a human edit
// trips over.
if (typeof server.description !== 'string') {
    problems.push('server.json description is missing')
} else if (server.description.length > 100) {
    problems.push(
        `server.json description is ${server.description.length} characters, max 100`
    )
}
if (!/^io\.github\.[^/]+\/[^/]+$/.test(server.name ?? '')) {
    problems.push(
        `server.json name ${server.name} is not of the form io.github.<owner>/<name>, ` +
            'which is what GitHub-based registry auth grants'
    )
}
// `files` decides what a user actually receives; a missing dist means an
// installable package with no server in it.
if (!(pkg.files ?? []).includes('dist')) {
    problems.push("package.json files does not include 'dist'")
}
if (pkg.bin?.['logisheets-mcp'] === undefined) {
    problems.push('package.json bin.logisheets-mcp is missing — hosts spawn it by name')
}

if (problems.length > 0) {
    console.error('check-release: not ready to publish')
    for (const p of problems) console.error(`  - ${p}`)
    process.exit(1)
}
console.log(
    `check-release: ok — ${pkg.name}@${expected} as ${server.name} ` +
        `(description ${server.description.length}/100 chars)`
)
