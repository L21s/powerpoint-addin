# CLAUDE.md

Guidance for AI assistants working in this repository.

## Project Overview

**L21s PowerPoint Add-in** — a Microsoft Office Task Pane add-in for PowerPoint,
built with TypeScript and bundled by Webpack. It augments PowerPoint with
L21s-branded productivity features: sticky notes, row/column guides, icon
search (via a Freepik proxy), employee photo insertion, branded background
fills, logo insertion, and banner creation.

- Host: PowerPoint (desktop/web) via Office.js Task Pane API
- UI framework: plain HTML + [Shoelace](https://shoelace.style) web components
  (loaded from CDN); no React despite `jsx: "react"` in `tsconfig.json` (legacy)
- Auth: MSAL.js (`@azure/msal-browser`) against Microsoft Entra ID
- Backend: Ktor proxy hosted on DigitalOcean
  (`powerpoint-addin-ktor-pq9vk.ondigitalocean.app`) that wraps Freepik and an
  employee directory
- Deployment: GitHub Pages (`https://l21s.github.io/powerpoint-addin/`) via
  `.github/workflows/static.yml` on pushes to `master`

## Commands

```bash
npm install              # install deps (first time only)
npm run build:dev        # webpack dev build to dist/
npm run build            # webpack production build (rewrites manifest URLs to prod)
npm run watch            # webpack --watch (development)
npm run dev-server       # webpack-dev-server on https://localhost:3000
npm run start:desktop    # side-load manifest.xml into desktop PowerPoint
npm run start:web        # side-load into PowerPoint for the web
npm run stop             # stop a running side-loaded add-in
npm run validate         # validate manifest.xml
npm run lint             # office-addin-lint check
npm run lint:fix         # office-addin-lint fix
npm run prettier         # format with office-addin-prettier-config
```

There are no tests in this repository.

### Typical dev loop
1. `npm install`
2. `npm run build:dev`
3. `npm run start:desktop` — PowerPoint launches with the add-in side-loaded;
   HMR keeps it in sync while you edit.

Mac debugging console is enabled via:
`defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`
(not available on Mac App Store Office builds).

## Repository Layout

```
manifest.xml                 Office Add-in manifest (TaskPane, icons, MSAL client id)
webpack.config.js            Build config; rewrites localhost → GitHub Pages URL in prod
tsconfig.json                target es5, allowJs, noEmitOnError
babel.config.json            @babel/preset-typescript
.eslintrc.json               plugin:office-addins/recommended
.github/workflows/static.yml GitHub Pages deploy on push to master
assets/                      icons, fonts, logos
src/
  app/
    bootstrap.ts             Office.onReady entrypoint; lazy-imports taskpane
    taskpane.ts              Central DOM-query module; exports element refs and the
                             top-level initializeTaskPaneListener()
    taskpane.html            Shoelace-based UI for all features
    taskpane.css             Styling
    listener/                DOM event wiring per feature (one file per feature)
    actions/                 Business logic that calls PowerPoint/Office.js APIs
    services/                HTTP/auth wrappers (authService, iconApiService,
                             employeeApiService)
    shared/
      consts.ts              ShapeType map, SLIDE_WIDTH/HEIGHT/MARGIN, FALLBACK_COLOR
      enums.ts               BannerPosition
      types.ts               FetchIconResponse, Employee, BannerOptions, ShapeTypeKey
      utils/
        powerPointUtil.ts    runPowerPoint(), getSelectedShapeWith()
        imageUtils.ts        getImageAsBase64()
  security/
    authConfig.ts            Entra ID clientId / authority / scope constants
    authClient.ts            MSAL PublicClientApplication singleton + helpers
```

## Architecture

The code is organized as a strict **listener → action → service** pipeline per
feature. When adding or modifying a feature, keep that separation.

1. **`src/app/taskpane.ts`** — single source of DOM element references. Every
   element used elsewhere is queried once here and exported. It also exports
   `initializeTaskPaneListener()`, which calls each feature's initializer.
2. **`src/app/listener/<feature>.ts`** — imports the element refs from
   `taskpane.ts`, attaches DOM event handlers, and delegates to action
   functions. Listeners should contain no PowerPoint API calls.
3. **`src/app/actions/<feature>.ts`** — business logic. Opens
   `PowerPoint.run` contexts (usually through `runPowerPoint` in
   `shared/utils/powerPointUtil.ts`), manipulates slides/shapes, and calls
   services for network access.
4. **`src/app/services/*.ts`** — network + auth. All authenticated calls go
   through `getRequestHeadersWithAuthorization()` in `authService.ts`, which
   acquires a token silently via MSAL.

Features currently wired up (one listener + action file each):
`stickyNotes`, `rowsColumns`, `searchDrawer` (icons + employees tabs),
`backgroundFills`, `logos`, `banner`. There's also `iconsPreview.ts`,
`employeesPreview.ts`, and `errorPopup.ts` as pure action modules (no
listener counterpart).

### PowerPoint API conventions

- Always run PowerPoint work inside `PowerPoint.run(async (context) => { ... })`
  or the thin wrapper `runPowerPoint(fn)` in `shared/utils/powerPointUtil.ts`.
- `load()` properties before reading them, then `await context.sync()`.
- Named shapes are how features find their own output later. Known names in
  use: `"RowLine"`, `"ColumnLine"`, `"Banner"`, `"Square"` (sticky notes),
  and the `ShapeTypeKey` name (e.g. `"Rectangle"`) used by background fills
  to remember the previous background shape type.
- Slide geometry assumes a standard 960×540 canvas — see `SLIDE_WIDTH`,
  `SLIDE_HEIGHT`, `SLIDE_MARGIN` in `shared/consts.ts`.

### Auth flow

- `src/security/authClient.ts` instantiates a `PublicClientApplication` with
  `supportsNestedAppAuth: true` (required for the Office dialog API) and
  `sessionStorage` caching with cookie fallback.
- `authService.loginWithDialog()` uses `msalApp.loginPopup()` and stores the
  account name under `localStorage["initials"]` (used by sticky notes).
- Service calls reuse `getRequestHeadersWithAuthorization()` which calls
  `acquireTokenSilent`. No interactive fallback on silent failure — expect
  an unhandled rejection if the session is dead. Login is triggered lazily
  by the search drawer (`listener/searchDrawer.ts`) and sticky-notes action.

### Build specifics

- Production build in `webpack.config.js` rewrites `https://localhost:3000/`
  in `manifest*.xml` to `https://l21s.github.io/powerpoint-addin/` and injects
  `process.env.BUILD_NUMBER` into the version (`1.0.0.<n>`). CI sets
  `BUILD_NUMBER=${{ github.run_number }}`.
- Webpack entry points: `polyfill` (core-js + regenerator) and `taskpane`
  (bootstrap + html). `browserslist: ["ie 11"]` plus `target: "es5"` — avoid
  syntax the target can't handle.
- Dev server runs on https with certs from `office-addin-dev-certs`.

### External dependencies to know

- `@shoelace-style/shoelace@2.20.1` is loaded from CDN in `taskpane.html`
  (not via npm). When querying Shoelace elements, cast to `any` or use the
  documented property names (e.g. `drawer["open"] = true`).
- `@azure/msal-browser` is used but **not declared in `package.json`** as of
  this writing — it resolves transitively or is installed manually. If you
  touch auth code and `npm ci` fails to resolve it, add it to `dependencies`.
- `@types/office-js-preview` provides the PowerPoint typings, including
  preview APIs like `parentGroup`, `addGroup`, and `getSelectedShapes`.

## Conventions

- 2-space indent, double quotes, semicolons — enforced by
  `office-addin-prettier-config`.
- File naming is camelCase (`backgroundFills.ts`, `employeeApiService.ts`).
  Feature names are consistent across `listener/`, `actions/`, and (where
  relevant) `services/`.
- Exported initializer functions follow `initialize<Feature>Listener()`.
- Constants live in `shared/consts.ts`; reusable types in `shared/types.ts`;
  enums in `shared/enums.ts`.
- Keep PowerPoint API calls out of listeners; keep DOM queries out of actions
  (import from `taskpane.ts` instead of re-querying).
- Errors surface via `showErrorPopup()` in `actions/errorPopup.ts`, which
  drives the `<sl-alert>` at the top of `taskpane.html`.

## Git / Branching

- Default branch: `master`. Pushes to `master` trigger the GitHub Pages
  deployment.
- Feature branches follow `feature/<kebab-name>` or `<scope>/<kebab-name>`
  (see recent history: `advanced-background-grouping`, `bugfixes`).
- **This session must commit to `claude/add-claude-documentation-rV4Fr`**
  per the harness instructions. Do not push to `master`.
- Only the `l21s/powerpoint-addin` GitHub repo is reachable via MCP tools.

## Gotchas

- `Office.onReady` may fire before Office.js is fully initialized on first
  launch — the lazy `await import("./taskpane")` in `bootstrap.ts` is the fix
  for that race (see commit `3357ead`). Don't move DOM work back into
  `bootstrap.ts`.
- `taskpane.ts` runs `document.querySelector*` at module load. If you add new
  UI elements, register them there; otherwise listeners will see `null`.
- The icon search filters server results to `author.name === "Smashicons"`
  and `family.name === "Basic Miscellany Lineal"` — this is intentional
  branding, not a bug.
- `ShapeType` in `shared/consts.ts` is a plain object (not an enum) keyed by
  friendly names that map to `PowerPoint.GeometricShapeType` values. Use
  `ShapeTypeKey` when typing inputs.
- Recent-icons and recent-colors are persisted in `localStorage`
  (`recentIcons` key) and in the DOM (`fixedColors` buttons), respectively.
