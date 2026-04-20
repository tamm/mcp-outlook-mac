# mcp-outlook-mac — agent notes

## Tests

Always run the suite with:

```
npm test
```

That runs `node --test 'test/*.test.js'` under the hood — don't invoke `node --test` directly. Keeps one canonical entry point.

## Pre-commit hook

Every commit runs `npm test` via `.githooks/pre-commit`. The repo sets `core.hooksPath=.githooks` locally, but a fresh clone needs:

```
git config core.hooksPath .githooks
```

If you're onboarding a fresh checkout and commits aren't triggering tests, that's the first thing to check.

## Tool descriptions are the source of truth

Tool names, schemas, and descriptions live in `index.js`'s `TOOLS` array. `SPEC.md` documents them — keep the two in sync when changing either.
