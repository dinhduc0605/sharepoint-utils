# Repository Guidelines

## Project Structure & Module Organization
This repository is a Manifest V3 Chrome extension with a flat layout. [`manifest.json`](/home/dinhduc/Projects/sharepoint-utils/manifest.json) declares permissions and the background service worker. [`background.js`](/home/dinhduc/Projects/sharepoint-utils/background.js) handles icon clicks, SharePoint requests, and alarm-based auto-refresh. [`content.js`](/home/dinhduc/Projects/sharepoint-utils/content.js) builds the injected modal UI and search logic. [`content.css`](/home/dinhduc/Projects/sharepoint-utils/content.css) contains all injected styles. [`README.md`](/home/dinhduc/Projects/sharepoint-utils/README.md) documents installation and user-facing behavior.

## Build, Test, and Development Commands
There is no bundler or package-based build step in this repo. Work directly on the checked-in files and reload the unpacked extension in Chrome.

- `xdg-open chrome://extensions/` or open `chrome://extensions/` manually: manage the unpacked extension.
- Load this folder with **Developer mode** enabled, then click **Reload** after changes.
- `git diff`: review the exact script and manifest changes before testing.

Manual verification is the current test path: open a SharePoint tab, click the extension icon, exercise search, settings, import/export, and refresh flows.

## Coding Style & Naming Conventions
Use 4-space indentation in JavaScript, JSON, and CSS to match the existing files. Prefer `camelCase` for variables and functions, `UPPER_SNAKE_CASE` for shared constants, and keep CSS class names prefixed with `sp-`. Preserve the current plain JavaScript style; do not introduce frameworks, build tooling, or module splitting unless the repo is explicitly being restructured.

## Testing Guidelines
There is no automated test suite yet. For behavior changes, verify:

- modal injection from the toolbar button
- SharePoint connection checks and fetch flows
- local storage reads/writes and auto-refresh behavior
- UI regressions in both Search and Settings tabs

When fixing edge cases, document the repro steps in the pull request.

## Commit & Pull Request Guidelines
Recent commits use short, imperative summaries such as `increase input height` and `deleted unused code`. Keep commit messages concise, lower-case, and focused on one change. Pull requests should include a brief description, manual test steps, and screenshots or short recordings for UI changes. Link the relevant issue or task when one exists.

## Security & Configuration Tips
Be careful with SharePoint URLs, host permissions, and stored data. Avoid widening `host_permissions` or logging sensitive document paths unless the change is necessary and reviewed.
