import { defineConfig } from 'vitest/config';
import { fileURLToPath } from 'node:url';

export default defineConfig({
  resolve: {
    alias: [{ find: /^\.\/core\//, replacement: fileURLToPath(new URL('./shared/scripts/core/', import.meta.url)).replace(/\\/g, '/') + '/' }],
  },
  test: {
    environment: 'jsdom',
    setupFiles: ['./shared/tests/setup.js'],
  },
});
