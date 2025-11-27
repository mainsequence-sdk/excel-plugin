# Developer Onboarding

Welcome to the MainSequence Excel Add-in. This guide is aimed at getting a new contributor productive quickly and keeping our workflow consistent.

## Quick Setup
- Prereqs: Node 18+, npm, Excel desktop with sideloading enabled, VS Code.
- Install deps: `npm install`.
- Local dev: `npm run start:local` (uses `manifest.local.xml` against https://localhost:3000 via webpack dev server).
- Dev/staging sideload: `npm run start:dev` (uses `manifest.dev.xml` built from the dev branch pipeline).
- Production check: `npm run start` (uses `manifest.xml`, assets published by CI/CD).

## Branching & Workflow
- Always branch from `dev` and create a dedicated feature branch for every change.
- Keep commits small and descriptive; open PRs back into `dev`.
- Before opening a PR, run the relevant npm scripts and add/update tests where possible.

## VS Code Expectations
- We cross-develop in VS Code. For every new test scenario or feature, add or update a launch configuration in `.vscode/launch.json` so others can debug the same entrypoint.
- Keep launch configurations named after the feature or ticket for easy discovery.

## Manifests & Builds
- `manifest.local.xml`: Local sideload; points to https://localhost:3000 assets from `npm run start:local`. Good for rapid iteration.
- `manifest.dev.xml`: Dev/staging sideload; uses artifacts produced from the `dev` branch CI/CD. Use when validating with shared test data.
- `manifest.xml`: Production sideload; served from the production bucket/https://main-sequence.app via CI/CD. Only updated after merges to main/production.
- When changing URLs or capabilities, keep the three manifests in sync and document the intended environment in your PR.

## Repository Layout
- `src/functions/functions.ts`: Office custom functions (e.g., `GET_DATA`) plus auth/token refresh helpers.
- `src/taskpane/`: React task pane app (sign-in flow, routing, UI components).
- `assets/`: Icons and static assets referenced by manifests.
- `deployment/`: Deployment scripts/configs for CI/CD buckets.
- `webpack.config.js`, `babel.config.json`, `tsconfig.json`: Build and tooling configuration.

## Feature Notes
- Custom functions currently centered on `GET_DATA` for pulling backend data into Excel with pagination and token refresh.
- Task pane handles user auth and wiring to the custom functions runtime.
- When adding functions, keep input flattening and boolean normalization consistent with the existing patterns.

## Development Guidelines
- Follow TypeScript/ESLint defaults in the repo; prefer small, testable units.
- Update manifests as needed when endpoints change and call out any environment-specific behavior in PRs.
- Document new endpoints, payloads, and required Excel sheet setup in README sections you touch.

## Excel Targets
- All features must work in both local Excel (desktop) and Excel on the web. Validate both unless explicitly exempted.
- Exception: Excel oil/QuantLib-style integrations may be desktop-only; note this in PRs when applicable.
- Maintain the online Excel example workbook demonstrating each function. Current shared link: https://mainsequence2185-my.sharepoint.com/:x:/g/personal/jose_mainsequence2185_onmicrosoft_com/IQCKHtHX4SWITbq5n1IoU4zUAX8bGAfeIqRPbnFjikpiAuY?e=WkW0lW. Keep the sheet examples up to date as new functions land.

## Upcoming Work Checklist
- [ ] Integrate Main Sequence Add-in with QuantLib (targeting functionality similar to https://bnikolic.co.uk/ql/qloil.html).
- [ ] Fixes for deployment (https://main-sequence.atlassian.net/browse/MSEAI-3?atlOrigin=eyJpIjoiNGE4ODExZjI5YjkzNGM2NGIyYWVhMjAxMDQwMzM1NmQiLCJwIjoiaiJ9).
- [ ] Implement Curve Inflation Function (https://main-sequence.atlassian.net/browse/MSEAI-3?atlOrigin=eyJpIjoiNGE4ODExZjI5YjkzNGM2NGIyYWVhMjAxMDQwMzM1NmQiLCJwIjoiaiJ9).
- [ ] Build `GET_ASSET` function (https://main-sequence.atlassian.net/browse/MSEAI-3?atlOrigin=eyJpIjoiNGE4ODExZjI5YjkzNGM2NGIyYWVhMjAxMDQwMzM1NmQiLCJwIjoiaiJ9).
