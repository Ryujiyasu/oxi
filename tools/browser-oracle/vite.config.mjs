// Multi-page build: the docx harness (index.html) and the pptx harness
// (index_pptx.html) ship together in dist/ so one static server serves both.
import { resolve } from 'path';
import { defineConfig } from 'vite';

export default defineConfig({
  base: './',
  build: {
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        pptx: resolve(__dirname, 'index_pptx.html'),
      },
    },
  },
});
