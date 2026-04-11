import { defineConfig } from 'vite';
import { resolve } from 'path';
import { copyFileSync, cpSync, mkdirSync, existsSync } from 'fs';

export default defineConfig({
  root: '.',
  base: './',
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      input: resolve(__dirname, 'index.html'),
    },
  },
  server: {
    port: 3000,
    https: true, // Office.js requires HTTPS even in dev
  },
  plugins: [
    {
      name: 'copy-manifest-and-assets',
      closeBundle() {
        const dist = resolve(__dirname, 'dist');
        // Copy manifests
        copyFileSync(resolve(__dirname, 'manifest.xml'), resolve(dist, 'manifest.xml'));
        copyFileSync(resolve(__dirname, 'manifest.json'), resolve(dist, 'manifest.json'));
        // Copy assets
        const assetsDir = resolve(dist, 'assets');
        if (!existsSync(assetsDir)) mkdirSync(assetsDir, { recursive: true });
        cpSync(resolve(__dirname, 'assets'), assetsDir, { recursive: true });
      },
    },
  ],
});
