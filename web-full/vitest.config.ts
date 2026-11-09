import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

const root = resolve(__dirname, '..')

export default defineConfig({
  root: __dirname,
  resolve: {
    alias: {
      // this checkout's engine, also in a worktree whose node_modules links to another checkout
      '@genoffice/docx-engine/lazy-media': resolve(root, 'packages/docx-engine/src/lazy-media.ts'),
      '@genoffice/docx-engine/zip-splice': resolve(root, 'packages/docx-engine/src/zip-splice.ts'),
      '@genoffice/docx-engine': resolve(root, 'packages/docx-engine/src/index.ts'),
      '@genoffice/font-metrics/resolve': resolve(root, 'packages/font-metrics/src/resolve.ts'),
      '@genoffice/font-metrics/word-faces.json': resolve(root, 'packages/font-metrics/word-faces.json'),
    },
  },
  test: {
    name: 'web-full',
    include: ['__tests__/**/*.test.ts'],
    environment: 'node',
  },
})
