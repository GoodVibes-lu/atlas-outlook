import { defineConfig } from 'vite';
import { resolve } from 'path';
import { copyFileSync, cpSync, mkdirSync, existsSync } from 'fs';
import { buildSync } from 'esbuild';

export default defineConfig({
  root: '.',
  base: './',
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        commands: resolve(__dirname, 'commands.html'),
        // Tableau de bord de la barre de gauche (application Microsoft 365, teams-app/manifest.json).
        tableauDeBord: resolve(__dirname, 'tableau-de-bord.html'),
      },
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
        // Smart Alerts (OnMessageSend) sur Outlook CLASSIQUE Windows : runtime JavaScript seul, qui
        // exige UN fichier .js sans import (manifest.xml, bt:Url « smartAlertsJsUrl »). IIFE esbuild.
        buildSync({
          entryPoints: [resolve(__dirname, 'src/launch-event.ts')],
          bundle: true,
          format: 'iife',
          platform: 'browser',
          target: 'es2019',
          minify: true,
          outfile: resolve(dist, 'launch-event.js'),
          logLevel: 'error',
        });
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
