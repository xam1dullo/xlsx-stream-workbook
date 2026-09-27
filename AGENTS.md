# AGENTS.md

Guidance for coding agents working in this repository.

## Project

`xlsx-stream-workbook` — streaming Excel (`.xlsx`) writer with multi-sheet support, built on `xlsx-write-stream` + `jszip`.

- **Entry point**: `index.js` → `lib/StreamingWorkbook.js`
- **Type definitions**: `index.d.ts`
- **Tests**: `node test/test.js` (`npm test`) — a script that writes real files to `test/output/`, not an assertion suite
- **Node**: `>=14`

## Agent skills

### Issue tracker

Issues live in this repo's GitHub Issues; skills use the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

Five canonical roles, each label string equal to its name (`needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: one `CONTEXT.md` and `docs/adr/` at the repo root. See `docs/agents/domain.md`.
