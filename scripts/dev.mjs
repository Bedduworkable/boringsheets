// Dev orchestrator: builds the Electron main+preload with esbuild (watch),
// starts the Vite dev server for the renderer, then launches Electron pointing
// at the dev server. No magic plugins — just the three pieces wired together.
import { createServer } from "vite";
import { context as esbuildContext } from "esbuild";
import { spawn } from "node:child_process";
import electronPath from "electron";

const isWatch = true;

// 1. Build main + preload (CommonJS for Electron's require-based runtime).
const buildCtx = await esbuildContext({
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
  sourcemap: true,
  target: "node20",
});
await buildCtx.rebuild();
if (isWatch) await buildCtx.watch();

// 2. Start Vite dev server for the renderer.
const viteServer = await createServer();
await viteServer.listen();
const info = viteServer.config.server;
const url = `http://localhost:${info.port}`;
console.log(`[dev] Vite serving renderer at ${url}`);

// 3. Launch Electron, telling it where the dev server lives.
let electronProc = null;
const launchElectron = () => {
  electronProc = spawn(electronPath, ["."], {
    stdio: "inherit",
    env: { ...process.env, VITE_DEV_SERVER_URL: url },
  });
  electronProc.on("close", async () => {
    await buildCtx.dispose();
    await viteServer.close();
    process.exit(0);
  });
};
launchElectron();
