import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
  root: resolve(__dirname),
  publicDir: resolve(__dirname, 'public'),
  resolve: {
    alias: {
      'hwpkit': resolve(__dirname, '../src/index.ts'),
    },
  },
  define: {
    'process.env': {},
    global: 'globalThis',
  },
});
