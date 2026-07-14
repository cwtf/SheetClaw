# Repository Guidelines

## Project Structure & Module Organization

SheetClaw is a React/TypeScript Excel task-pane add-in. Application code lives in `src/`: `taskpane/` contains the UI and Office entry point, `agent/` owns the agent loop, `workbook/` implements Excel tools and snapshots, `adapters/` integrates model providers, `web/` handles search/fetch providers, and `store/`, `auth/`, `pricing/`, and `types/` provide shared services. Keep unit tests beside their modules in `__tests__/`. Live endpoint checks belong in `src/web/__integration__/`. Static icons and privacy content live in `public/`; architecture notes and specifications live in `docs/`. Treat `dist/` as generated output.

## Build, Test, and Development Commands

- `npm ci` installs the lockfile-pinned dependencies (CI uses Node 22).
- `npm run install-certs` installs local Office HTTPS certificates once.
- `npm run dev` starts Vite at `https://localhost:3000`.
- `npm run sideload` opens/registers the add-in in desktop Excel; keep the dev server running.
- `npm run build` runs strict TypeScript checks and creates `dist/`.
- `npm test` runs all unit tests once; `npm run test:watch` runs them interactively.
- `npm run test:providers` runs live-network/CORS integration checks and requires internet access.
- `npm run validate-manifest` validates `manifest.xml`.

Run one test with `npx vitest run src/agent/__tests__/loop.test.ts`.

## Coding Style & Naming Conventions

Use two-space indentation, single quotes, semicolons, and trailing commas in multiline literals. TypeScript is strict: avoid `any`, unused declarations, and unsafe Office globals. Use `PascalCase` for React components/classes/types, `camelCase` for functions and variables, and kebab-case filenames such as `context-builder.ts`; component filenames use `PascalCase.tsx`. Keep Office.js work inside `Excel.run()` and `ctx.sync()`. No standalone formatter or linter is configured, so follow surrounding code and rely on `npm run build`.

## Testing Guidelines

Vitest runs in a Node environment. Name unit files `*.test.ts` and integration files `*.integration.ts`. Mock Office.js or inject a `LoopRunner`; do not assume `Excel` or `Office` exists in tests. Add focused regression tests for behavior changes. There is no enforced coverage threshold, but new branches and failure paths should be exercised.

## Commit & Pull Request Guidelines

Recent history follows Conventional Commits, commonly scoped: `feat(web): ...`, `refactor(agent): ...`, or `chore: ...`. Keep commits focused and imperative. Pull requests should explain user-visible impact, link relevant issues, list verification commands, and include screenshots for task-pane UI changes. Call out manifest, provider, credential-storage, or network/CORS changes explicitly; never commit API keys or tokens.
