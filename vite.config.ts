import path from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  base: "./",
  build: {
    // minify: false,
    outDir: 'build',
    rollupOptions: {
      input: {
        index: path.resolve(__dirname, 'index.html'),
        settings: path.resolve(__dirname, 'settings.html'),
        connectionError: path.resolve(__dirname, 'connection-error.html'),
      },
      external: [/Asc/],
      output: {
        entryFileNames: 'scripts/[name].js',
      }
    }
  }
});
