# Think & Make PWA — Claude Code Instructions

## Project Overview
Google Apps Script web app (PWA) for the Think & Make programme. The codebase lives in two GitHub repos and is synced to a Google Apps Script project via `clasp`. Any code change must be pushed to both Git remotes AND deployed to Apps Script to take effect live.

## After Every Code Change — Always Do These 3 Steps

### 1. Push to Apps Script
```
clasp push --force
```

### 2. Deploy the live versioned web app
```
clasp deploy --deploymentId AKfycbyRWOzhC12yIn9Lnlrzo088n1t-mhsuzMK7Ua0ZEuQjmht-N7VuvZ1jq6b4bfVcK_lX7g --description "<short description of change>"
```
`clasp push` alone does NOT update the live app — the `clasp deploy` step is mandatory every time.

### 3. Push to both Git remotes
This repo has two remotes — always push to both:
```
git push origin <branch>
git push iif <branch>
```

## Git Remotes
| Remote | Repo URL | Purpose |
|--------|----------|---------|
| `origin` | https://github.com/Geetha-Inquilab/tm_pwa_test | Primary repo (Inqui-Lab org) |
| `iif` | https://github.com/IIF-2026/tm-form | IIF partner org mirror |

To verify both remotes are configured on your machine:
```
git remote -v
```
If `iif` is missing, add it:
```
git remote add iif https://github.com/IIF-2026/tm-form.git
```

## Google Apps Script — Project Info
| Field | Value |
|-------|-------|
| Script ID | `1b1wvGHy_712iFlA6Ay-rEkF3MSfLlrD54QP1T6R3_aMzInJpW5pQYmsY` |
| Live deployment ID | `AKfycbyRWOzhC12yIn9Lnlrzo088n1t-mhsuzMK7Ua0ZEuQjmht-N7VuvZ1jq6b4bfVcK_lX7g` |
| Live web app URL | `https://script.google.com/macros/s/AKfycbyRWOzhC12yIn9Lnlrzo088n1t-mhsuzMK7Ua0ZEuQjmht-N7VuvZ1jq6b4bfVcK_lX7g/exec` |
| IIF portal (GitHub Pages) | https://iif-2026.github.io/tm-form/ |

The Apps Script project serves the entire PWA — `Code.gs` is the backend and `index.html` is the frontend. `clasp` syncs this repo to the script project. The live deployment ID is the versioned endpoint that the IIF portal and all users hit; every `clasp push` without a `clasp deploy` only updates the HEAD (editor) version, not the live app.

To set up `clasp` on a new machine:
```
npm install -g @google/clasp
clasp login
```
Then open this repo — `.clasp.json` already has the script ID configured.

## Git Workflow
- Work on feature branches, not directly on `main` or `multiobservations`
- Create PRs into `multiobservations` (not directly into `main`)
- Commit messages should be descriptive (what changed and why)

## Key Files
- `Code.gs` — all server-side logic (Google Apps Script)
- `index.html` — entire frontend (single-file PWA)
- `appsscript.json` — Apps Script manifest
- `.clasp.json` — clasp config linking this repo to the Apps Script project
