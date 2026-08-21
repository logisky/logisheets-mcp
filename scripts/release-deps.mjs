#!/usr/bin/env node
//
// Flip the local `file:` dependencies to registry ranges before publishing.
//
// Development points logisheets-runtime / logisheets-logician at a sibling
// LogiSheets checkout, because this server is usually built alongside
// unreleased engine changes. A published tarball must depend on the registry
// instead, so run this (and commit the result) as part of a release:
//
//   node scripts/release-deps.mjs 1.12.0     # pin to a released engine version
//   npm publish
//
// Pass no argument to reuse whatever caret range is already recorded below.

import {readFileSync, writeFileSync} from 'node:fs'
import {dirname, join} from 'node:path'
import {fileURLToPath} from 'node:url'

const PKG = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json')
/** Local deps that must become registry ranges, with their fallback version. */
const LOCAL_DEPS = {
    'logisheets-logician': '1.12.0',
    'logisheets-runtime': '1.12.0',
}

const version = process.argv[2]
if (version !== undefined && !/^\d+\.\d+\.\d+/.test(version)) {
    console.error(`release-deps: "${version}" is not a semver version`)
    process.exit(1)
}

const pkg = JSON.parse(readFileSync(PKG, 'utf8'))
let changed = false
for (const [name, fallback] of Object.entries(LOCAL_DEPS)) {
    const current = pkg.dependencies?.[name]
    if (current === undefined) {
        console.error(`release-deps: ${name} is not a dependency`)
        process.exit(1)
    }
    const range = `^${version ?? fallback}`
    if (current === range) continue
    pkg.dependencies[name] = range
    console.log(`release-deps: ${name} ${current} -> ${range}`)
    changed = true
}

if (!changed) {
    console.log('release-deps: nothing to change')
} else {
    writeFileSync(PKG, `${JSON.stringify(pkg, null, 4)}\n`)
}

const stillLocal = Object.entries(pkg.dependencies ?? {}).filter(([, v]) =>
    String(v).startsWith('file:')
)
if (stillLocal.length > 0) {
    console.error(
        `release-deps: still local: ${stillLocal
            .map(([k]) => k)
            .join(', ')} — do not publish`
    )
    process.exit(1)
}
