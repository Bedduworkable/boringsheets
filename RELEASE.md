# Releasing BoringSheets

Production installers are built with [electron-builder](https://www.electron.build/).
Configuration lives in [`electron-builder.yml`](./electron-builder.yml).

## Prerequisites

```bash
npm install
```

The packaging step always runs the project build first (`node scripts/build.mjs`),
which bundles the Vite renderer into `dist/` and esbuilds the Electron main/preload
into `dist-electron/*.cjs`. electron-builder then packages those artifacts plus
`package.json` into an installer.

## Building installers

| Command            | Output                                   |
| ------------------ | ---------------------------------------- |
| `npm run dist`     | Installer for the **current** platform   |
| `npm run dist:mac` | macOS `.dmg`                             |
| `npm run dist:win` | Windows NSIS `.exe` installer            |
| `npm run pack`     | Unpacked app in a directory (no installer; fast config check) |

All output is written to **`release/`**:

- macOS: `release/BoringSheets-<version>-<arch>.dmg` (+ unpacked `release/mac*/BoringSheets.app`)
- Windows: `release/BoringSheets Setup <version>.exe`
- Linux (optional): `release/BoringSheets-<version>.AppImage`

> Cross-building (e.g. `dist:win` from macOS) requires the relevant toolchain
> (Wine for Windows NSIS). Building each platform on its native OS / CI runner is
> the most reliable path.

## Signing status: UNSIGNED

Current builds are **unsigned** — no Apple Developer ID and no Windows
code-signing certificate are configured.

Consequences:

- **macOS**: `mac.identity` is `null` and `hardenedRuntime` is `false`. The app is
  not notarized, so Gatekeeper will warn. Users must right-click → Open, or run
  `xattr -dr com.apple.quarantine /Applications/BoringSheets.app`.
- **Windows**: the NSIS installer is unsigned, so SmartScreen will show an
  "unknown publisher" warning.

## Adding signing later

### macOS (signing + notarization)

1. Obtain an **Apple Developer ID Application** certificate; export it as a `.p12`.
2. Provide it to electron-builder via env vars (preferred for CI):

   ```bash
   export CSC_LINK=/path/to/developer-id.p12   # or a base64 data URL
   export CSC_KEY_PASSWORD='p12-password'
   ```

   …or set `mac.identity` in `electron-builder.yml` to the certificate's common
   name and remove `identity: null`.
3. In `electron-builder.yml`, enable hardened runtime and notarization:

   ```yaml
   mac:
     hardenedRuntime: true
     gatekeeperAssess: false
     entitlements: build/entitlements.mac.plist
     entitlementsInherit: build/entitlements.mac.plist
     notarize: true            # electron-builder >= 24 native notarization
   ```
4. Supply notarization credentials via env vars:

   ```bash
   export APPLE_ID='you@example.com'
   export APPLE_APP_SPECIFIC_PASSWORD='abcd-efgh-ijkl-mnop'
   export APPLE_TEAM_ID='XXXXXXXXXX'
   ```

   (Alternatively use `APPLE_API_KEY` / `APPLE_API_KEY_ID` / `APPLE_API_ISSUER`
   for App Store Connect API-key notarization, or wire a custom `afterSign`
   notarize hook with `@electron/notarize`.)

### Windows (code signing)

1. Obtain a code-signing certificate (`.pfx` / `.p12`). EV certs avoid the
   SmartScreen reputation warning.
2. Provide it via the same env vars electron-builder reads on Windows:

   ```bash
   export CSC_LINK=/path/to/codesign.pfx
   export CSC_KEY_PASSWORD='pfx-password'
   ```

   electron-builder signs the app and NSIS installer automatically. For
   cloud/HSM signing (e.g. Azure Trusted Signing), configure `win.signtoolOptions`
   or a custom `win.sign` hook instead.

Re-run `npm run dist:mac` / `npm run dist:win` after exporting these variables.
