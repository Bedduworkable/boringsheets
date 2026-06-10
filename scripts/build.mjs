// Production build: bundle the renderer with Vite and the Electron main+preload
// with esbuild. Output goes to dist/ (renderer) and dist-electron/ (shell).
import { build as viteBuild } from "vite";
import { build as esbuildBuild } from "esbuild";

await viteBuild();

await esbuildBuild({
  entryPoints: {
    main: "electron/main.ts",
    preload: "electron/preload.ts",
  },
  bundle: true,
  platform: "node",
  format: "cjs",
  outdir: "dist-electron",
  outExtension: { ".js": ".cjs" },
  external: ["electron"],
  minify: true,
  target: "node20",
});

console.log("[build] done -> dist/ + dist-electron/");
