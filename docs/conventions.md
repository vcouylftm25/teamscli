# Code conventions

Rules for anyone (human or agent) changing `teamscli`. They are short on purpose: when in doubt, copy the reference files listed under each section.

## Runtime and tooling

- Node.js 24+, ESM, TypeScript executed directly by Node's type stripping. No build step in development; `tsc` only type-checks (`npm run check`) and builds `dist/` for publishing.
- Only erasable TypeScript syntax (`erasableSyntaxOnly`): no `enum`, no `namespace`, no constructor parameter properties. Use string-literal unions and `as const` objects.
- Relative imports carry the `.ts` extension: `import { CliError } from '../errors.ts'`. Type-only imports use `import type`.
- Built-ins first: `node:sqlite`, `fetch`, `node:test`, `node:util` `parseArgs`, `node:fs`, `node:child_process`.

## Dependencies

No runtime dependencies. A new dependency (runtime or dev) needs a justification recorded in the table below, in the same commit that adds it.

| Package | Kind | Why |
|---|---|---|
| `typescript` | dev | Type checking and the `dist/` build. |
| `@types/node` | dev | Types for Node built-ins. |
| `@resvg/resvg-wasm` | runtime | Rasterize discovered SVG favicons to small PNGs for Teams inline images, without native modules. |

## Layout

```text
src/
  main.ts            # Executable: calls runCli and sets the exit code. Nothing else.
  models.ts          # Domain types shared across layers.
  errors.ts          # CliError, error codes and exit codes.
  cli/               # Argument parsing, command registry, output. No HTTP, no SQL.
    commands/        # One file per command.
  auth/              # Login, token refresh, secret stores.
  teams/             # Everything about Teams web endpoints and their formats.
  sync/              # Orchestrates teams/ -> store/.
  store/             # SQLite: schema, migrations, repositories. No HTTP.
  runtime/           # Paths and process-level helpers.
tests/               # Mirrors src/ (tests/store/conversations.test.ts tests src/store/conversations.ts).
  helpers/           # Fakes and builders shared by tests.
```

- Dependencies point inwards: `cli -> sync -> teams, store -> models`. `teams/` never imports `store/` and vice versa; `sync/` joins them.
- External formats (Teams JSON field names, `threadType`, `skypetoken`) stay in `teams/`. Everything else sees only `models.ts` types.
- Create a file when a feature needs it. No empty modules, no `index.ts` barrels.

## Naming

- Files: `kebab-case.ts`. Command files: the command name with spaces replaced by `-` (`messages search` -> `cli/commands/messages-search.ts`).
- Types and interfaces: `PascalCase`. Functions and variables: `camelCase`. Constants that are part of a contract: `UPPER_SNAKE_CASE`.
- SQL: `snake_case` tables and columns, plural table names. Repositories map rows to `camelCase` models.
- JSON output keys: `camelCase`, identical to the model fields.

## Functions, not classes

- Modules export plain functions. A group of operations that share dependencies is a factory returning an object (`createTeamsClient(deps)`), not a class.
- Classes only for errors (`CliError`).
- No singletons or module-level state. Anything with side effects arrives as a parameter.

## Dependency injection

Pass side effects explicitly so tests never touch the network, the real clock, the real home directory or the system secret store:

| Dependency | Parameter | Test replacement |
|---|---|---|
| HTTP | `fetch: typeof fetch`, `session: Session` | `fakeFetch()` and `fakeSession()` in `tests/helpers/fake-fetch.ts` |
| Time | `now: () => Date`, `sleep: (ms) => Promise<void>` | fixed date, no-op sleep |
| Database | `db: DatabaseSync` | `openDatabase(':memory:')` |
| Filesystem location | `home: string` (from `runtime/paths.ts`) | `mkdtemp` directory |
| Environment | `env: NodeJS.ProcessEnv` | plain object |
| External commands | `exec: (cmd, args, input) => Promise<{ stdout, stderr, code }>` | function returning canned output |

Defaults (the real `fetch`, `process.env`) are wired once, at the edge: `main.ts` and the command files.

## Models and data

- Model fields are always present; unknown values are `null`, not `undefined` or missing. This keeps `--json` output stable.
- Timestamps are ISO 8601 UTC strings (`2026-10-27T12:04:00.000Z`) in models, SQLite and JSON. They sort lexicographically.
- Teams ids are kept verbatim (`19:…@thread.v2`, `8:orgid:…`, `48:notes`). Message ids are unique only within a conversation: the key is `(conversationId, id)`.

## Errors and exit codes

- Expected failures throw `CliError(code, message, details?)` from `src/errors.ts`. The message is one sentence for the user, in English, saying what to do when possible.
- Exit codes come from the code (`EXIT_CODES` in `errors.ts`); never call `process.exit` outside `main.ts`.
- Anything else is a bug: it exits with `internal` (1). Set `TEAMSCLI_DEBUG=1` to print the stack.
- Never put tokens, cookies, message content or full URLs with query strings in error messages.

## Output

- stdout carries only the result. Progress, warnings and diagnostics go to stderr through `ctx.log`.
- Commands return data; the CLI renders it (see `cli/output.ts`). Commands do not write to stdout themselves, except event streams through `ctx.emit`.
- Human output is for people and may change. `--json` and `--events` are contracts: adding fields is fine, renaming or removing them is a breaking change.
- Human rendering formats valid ISO values whose field names end in `At` in the process's local timezone. `--json` keeps the original UTC ISO strings.

## Calling Teams

- All HTTP goes through `createTeamsClient` (`src/teams/client.ts`): it adds credentials from the `Session`, retries 429 honouring `Retry-After`, and maps 401 to `auth_required` and other failures to `remote`.
- Build URLs from `session.endpoints()`; continue paging with the full `backwardLink` or `syncState` URL Teams returned, never by rebuilding it.
- A `teams/<resource>.ts` module exports functions that take a `TeamsClient` (and `Endpoints`) and return `models.ts` types. Raw payload types are local to that module and not exported.
- Test with `fakeFetch` routes and small synthetic payloads in `tests/fixtures/teams/`.

## Storage

- One SQLite database per account (`runtime/paths.ts` `databasePath`), opened with `openDatabase` (`src/store/database.ts`): WAL, busy timeout, foreign keys on.
- Commands with `mutates: false` open it with `{ readOnly: true }`; they never create or migrate it.
- The schema lives in `src/store/migrations/NNN_name.sql`, applied in order and tracked by `PRAGMA user_version`. Never edit a migration that is in `main`; add the next number.
- A repository is `src/store/<table>.ts` with plain functions taking `db` and returning models; SQL stays there. Batch writes go through `transaction(db, …)`. Follow `src/store/conversations.ts`.
- Writes are upserts keyed by the Teams id (`(conversation_id, id)` for messages): sync pages arrive in any order and repeat. A message row is only replaced by an equal or higher `version`.
- Upserts are always `INSERT … ON CONFLICT (…) DO UPDATE`. Never `INSERT OR REPLACE` or `REPLACE`: they delete and reinsert the row, changing its rowid, and without `recursive_triggers` the delete trigger does not fire, so `messages_fts` drifts out of sync.
- `messages_fts` is kept in sync by triggers; repositories write only to `messages`.

## Adding a command

1. Create `src/cli/commands/<words-joined-by-dash>.ts` that default-exports a `Command` (`src/cli/command.ts`). The registry loads every file in that directory; no shared file is edited.
2. Declare `mutates: true` if it changes anything, local (database, files, secrets) or remote. `--read-only` refuses those before `run` is called. Under `--read-only` nothing is written at all: remote commands get their session from `teamsSession(ctx)`, which renews tokens in memory without saving the rotated refresh token.
3. Parse and validate the command's own options inside `run`, and throw `CliError('usage', …)` for bad input.
4. Return data (models or plain objects); add `render` only when the default rendering in `cli/output.ts` is not good enough.
   Set `optionHelp` by long option name to show a description and optional value name in command help. Options without help metadata show only their name. Set `arguments` for positional argument text.
   For message output, use `renderMessage`, `renderMessages` or `renderMessageDetail` from `cli/output.ts`; they use plain `text`, never HTML `content`.
5. Test it in-process with `runTestCli` (`tests/helpers/cli.ts`), always with `--store` pointing to a temporary directory.

```ts
import type { Command } from '../command.ts';

const command: Command<{ deleted: number }> = {
  name: 'things remove',
  summary: 'Remove a thing',
  usage: '<id>',
  maxPositionals: 1,
  mutates: true,
  async run(ctx, { positionals }) {
    // build real dependencies from ctx (ctx.home, ctx.account, ctx.fetch) and call the layer below
    return { deleted: 1 };
  },
};

export default command;
```

The CLI contract (global options, `--json` envelope, `--events` envelope, exit codes) is defined in `src/cli/command.ts`, `src/cli/output.ts` and `src/errors.ts` and tested in `tests/cli/program.test.ts`. Changing it is a contract change.

## Tests

- `node:test` and `node:assert/strict`. One test file per source file, under `tests/` with the same path.
- Test behaviour through the module's public functions; no mocking of internals.
- Data is synthetic. Build it with the builders in `tests/helpers/` and override only the fields the test cares about. Raw Teams payloads, when needed, live in `tests/fixtures/` as small hand-written JSON: invented names (`Alice Example`), invented ids, `example.com` addresses. Never paste captured traffic.
- No test talks to Teams, Azure AD, the real home directory or the system secret store.
- `npm run check` and `npm test` must pass before every commit.

## Commits

- Conventional Commits with a scope: `feat(store): …`, `fix(cli): …`, `docs(project): …`, `test(sync): …`.
- One coherent, verified step per commit. Only the files of that step.
- Never commit secrets, sessions, real databases, captured traffic or private media.
