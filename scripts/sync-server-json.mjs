#!/usr/bin/env node
//
// Carry package.json's version into server.json. Wired to npm's `version`
// lifecycle hook, which runs after the bump and before the commit, so
// `npm version 0.2.0` updates both files in one commit.
//
// The registry keeps the version in two places — the server's own `version`
// and the version of the npm package it points at — and both have to match
// what actually got published. Editing them by hand is the step most likely
// to be forgotten, and `npm publish` cannot be undone.

import {readFileSync, writeFileSync} from 'node:fs'
import {dirname, join} from 'node:path'
import {fileURLToPath} from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const serverPath = join(root, 'server.json')
const {version} = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const server = JSON.parse(readFileSync(serverPath, 'utf8'))

const before = [server.version, server.packages?.[0]?.version]
server.version = version
if (server.packages?.[0] !== undefined) server.packages[0].version = version
writeFileSync(serverPath, `${JSON.stringify(server, null, 4)}\n`)

console.log(
    `sync-server-json: ${before.join(' / ')} -> ${version} (server.json version / packages[0].version)`
)
