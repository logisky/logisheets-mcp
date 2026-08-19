#!/usr/bin/env node
//
// Point the three LogiSheets packages at a sibling checkout instead of the
// registry, for developing the server and the engine together.
//
// `npm install` fetches the published versions; running this afterwards
// replaces those three directories with symlinks into ../LogiSheets. Node then
// resolves each package's own dependencies from the LogiSheets monorepo, so its
// workspace links (and the freshly built WASM in packages/node/wasm) are what
// actually runs. Re-run it after any `npm install`.
//
//   npm install && npm run link:local
//
// Set LOGISHEETS_REPO to a checkout elsewhere.

import {existsSync, lstatSync, rmSync, symlinkSync, mkdirSync} from 'node:fs'
import {dirname, join, resolve} from 'node:path'
import {fileURLToPath} from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const repoRoot = resolve(here, '..')
const logisheets = resolve(
    process.env.LOGISHEETS_REPO ?? join(repoRoot, '..', 'LogiSheets')
)

/** package name → its directory inside the LogiSheets monorepo. */
const LINKS = {
    logisheets: 'packages/node',
    'logisheets-runtime': 'packages/runtime',
    'logisheets-logician': 'packages/logician',
}

if (!existsSync(logisheets)) {
    console.error(
        `link-local: no LogiSheets checkout at ${logisheets}\n` +
            `Set LOGISHEETS_REPO to point at one.`
    )
    process.exit(1)
}

const nodeModules = join(repoRoot, 'node_modules')
mkdirSync(nodeModules, {recursive: true})

for (const [pkg, dir] of Object.entries(LINKS)) {
    const target = join(logisheets, dir)
    if (!existsSync(target)) {
        console.error(`link-local: missing ${target}`)
        process.exit(1)
    }
    const link = join(nodeModules, pkg)
    // lstat, not existsSync: a broken symlink must still be replaced.
    try {
        lstatSync(link)
        rmSync(link, {recursive: true, force: true})
    } catch {
        /* nothing there yet */
    }
    symlinkSync(target, link, 'dir')
    console.log(`link-local: ${pkg} -> ${target}`)
}

console.log(
    'link-local: done. Build the linked packages if you changed them:\n' +
        '  (in LogiSheets) yarn workspace logisheets-runtime build\n' +
        '  (in LogiSheets) yarn workspace logisheets-logician build'
)
