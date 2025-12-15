import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
  root: resolve(__dirname),
  resolve: {
    alias: {
      'cellify': resolve(__dirname, '../src/index.ts'),
    },
  },
  server: {
    port: 5432,
    open: true,
  },
});
