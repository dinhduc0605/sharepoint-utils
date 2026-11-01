# Repository Guidelines

## Project Structure & Module Organization
This repository is a Manifest V3 Chrome extension with a flat layout. [`manifest.json`](manifest.json) declares permissions and the background service worker. [`background.js`](background.js) injects the modal on icon click, checks that SharePoint tabs are open, and proxies SharePoint REST calls by running `fetch` inside an open SharePoint tab (adding `X-RequestDigest` for non-GET requests). [`content-utils.js`](content-utils.js) holds pure, testable helpers exposed as `globalThis.SharePointUtils` and is injected before [`content.js`](content.js), which builds the modal UI, search, settings, backup, and check-in/check-out flows. [`content.css`](content.css) contains all injected styles. [`tests/`](tests) contains `node:test` suites for the helpers. [`README.md`](README.md) documents installation and user-facing behavior.

## Build, Test, and Development Commands
There is no bundler or package-based build step in this repo. Work directly on the checked-in files and reload the unpacked extension in Chrome.

- Open `chrome://extensions/`, enable **Developer mode**, load this folder with **Load unpacked**, then click **Reload** after changes.
- `node --check background.js`, `node --check content-utils.js`, `node --check content.js`: syntax checks.
- `node --test tests/content-utils.test.mjs`: helper unit tests.
- `git diff`: review the exact script and manifest changes before testing.

Manual verification is still required for UI and SharePoint flows: open a SharePoint tab, click the extension icon, exercise search, settings, import/export, and check-in/check-out.

## Coding Style & Naming Conventions
Use 4-space indentation in JavaScript, JSON, and CSS to match the existing files. Prefer `camelCase` for variables and functions, `UPPER_SNAKE_CASE` for shared constants, and keep CSS class names prefixed with `sp-`. Preserve the current plain JavaScript style; do not introduce frameworks, build tooling, or module splitting unless the repo is explicitly being restructured. Put new pure logic in `content-utils.js` and cover it with a test.

Escape any dynamic value (file names, paths, queries, error messages) with `escapeHtml` before inserting it through `innerHTML`, including messages passed to `showStatus` and `showSettingsStatus`. Pass user search terms through `toKqlSearchTerm` before building KQL queries.

## Testing Guidelines
Run the syntax checks and helper tests above for every change. For behavior changes, also verify manually:

- modal injection from the toolbar button
- SharePoint connection checks (per configured tenant origin) and fetch flows
- local storage reads/writes, import/export
- UI regressions in both Search and Settings tabs

When fixing edge cases, document the repro steps in the pull request.

## Commit & Pull Request Guidelines
Recent commits use short, imperative summaries such as `increase input height` and `deleted unused code`. Keep commit messages concise, lower-case, and focused on one change. Pull requests should include a brief description, manual test steps, and screenshots or short recordings for UI changes. Link the relevant issue or task when one exists.

## Security & Configuration Tips
Be careful with SharePoint URLs, host permissions, and stored data. Avoid widening `host_permissions` or logging sensitive document paths unless the change is necessary and reviewed. The default tenant (`TENANT_ORIGIN`) and site (`DEFAULT_SITE`) are hard-coded at the top of `content.js`.
