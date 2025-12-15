import { defineConfig } from 'vite';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

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
