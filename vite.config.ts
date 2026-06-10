import { defineConfig } from "vite";

// Renderer build only. The Electron main/preload are built separately with esbuild
// (see scripts/dev.mjs and scripts/build.mjs). Using a relative base so the packaged
// app can load assets from the local filesystem.
export default defineConfig({
  root: "src",
  base: "./",
  build: {
    outDir: "../dist",
    emptyOutDir: true,
    target: "es2022",
  },
  server: {
    port: 5123,
    strictPort: true,
  },
});
