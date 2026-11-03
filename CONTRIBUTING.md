# Contributing to dsh-onlyoffice-bridge

Thanks for your interest in making ONLYOFFICE a first-class citizen in the
DeepSeek Harness ecosystem. This document explains how to set up the project
and what we expect from a pull request.

## Getting a working checkout

The plugin is self-contained. A fresh clone already carries every document
dependency it needs under `vendor/`, so **no `npm install` is required to run
the tests**. That property is a core feature of the repository and CI
deliberately enforces it.

```bash
git clone https://github.com/lianghui-cmd/dsh-onlyoffice-bridge.git
cd dsh-onlyoffice-bridge
npm test
```

You should see two summaries:

```
ℹ tests 36
ℹ pass 36
ℹ fail 0
...
全部用例通过
```

If either suite fails on your machine, please open an issue before submitting
a PR so we can see whether it is environmental.

### When `npm install` is genuinely needed

Only if you want to touch the vendored dependencies themselves:

```bash
npm install         # restores node_modules/
npm run vendor      # regenerates vendor/ from node_modules/
npm run vendor:check # preview size without writing
```

`vendor/` is checked in on purpose. Commit any regeneration together with the
dependency bump, and keep `vendor/MANIFEST.json` in sync.

## Project layout

| Path | What lives there |
| --- | --- |
| `index.js` | Plugin entry: routes, JWT, callbacks, cache, tool registration |
| `lib/definitions.js` | The 22 tool definitions (name, description, parameter schema) |
| `lib/tools.js` | Tool handlers: validation, locks, batch, template-fill, insert-image |
| `lib/document.js` | Pure document helpers: markdown->docx, xlsx, docx XML patching |
| `lib/optional-deps.js` | Lazy loading of document dependencies (vendor first) |
| `lib/vendor-resolver.js` | Makes vendor packages resolve each other |
| `scripts/vendor-deps.mjs` | Regenerates `vendor/` with a pruned runtime closure |
| `scripts/validate-patch.mjs` | Validates `cordis.patch.yml` without extra dependencies |
| `test/` | `node:test` suites; `validate-patch.test.mjs` is standalone |
| `vendor/` | Bundled document dependencies (must be committed) |

## Style

- Plain JavaScript with JSDoc types; `npm run typecheck` must stay clean.
- Comments and error messages are written in the user's language (Chinese),
  commit messages in English (Conventional Commits).
- No new runtime dependency is added without a strong reason. Prefer
  `node:*` built-ins.

## Commits and pull requests

- Use [Conventional Commits](https://www.conventionalcommits.org/):
  `feat: ...`, `fix: ...`, `docs: ...`, `chore: ...`.
- Keep commits focused and tested. Every PR must pass CI on all four
  matrices (Node 22/24 x Linux/Windows).
- If your change touches `vendor/`, mention which dependency moved and why.
- Update `README.md` when behavior, tool names, or configuration change.

## Adding a tool

1. Add the definition to `lib/definitions.js` (name, description, schema).
2. Implement the handler in `lib/tools.js` with the same key.
3. Add at least one test in `test/tools.test.mjs`.
4. Bump the tool count in `README.md` and `package.json` description.

## Reporting bugs

Please include:

- The tool name and arguments you used
- The exact error text (from `onlyoffice_status` or the DSH log)
- Whether the plugin was loaded from `vendor/` or an external `node_modules`
- Your DSH and ONLYOFFICE Document Server versions

## License

By contributing, you agree that your contributions will be licensed under the
[MIT License](LICENSE).
