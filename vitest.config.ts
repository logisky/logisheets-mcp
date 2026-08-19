import {defineConfig} from 'vitest/config'

export default defineConfig({
    test: {
        // The engine is a single WASM instance shared per process; workbooks are
        // isolated by book id, so tests can share it. Keep them serial anyway so
        // a panic in one file can't be blamed on another.
        fileParallelism: false,
        testTimeout: 30_000,
    },
})
