# Command reference

Run `teamscli --help` for the command list or `teamscli <command> --help` for a command's usage. Global options may appear before or after the command. Examples use fictional names and IDs; replace them with values from your own local store. Run mutating examples only against destinations you are authorized to use.

## Global options

| Option | Meaning |
|---|---|
| `--account NAME` | Select an account. Resolution order: this option, `TEAMSCLI_ACCOUNT`, the saved default, then `default`. |
| `--store DIR` | Use a local state root. Default: `TEAMSCLI_HOME`, then `~/.teamscli`. |
| `--json` | Write one JSON result or error envelope to stdout. See [agent integration](agents.md#json-output). |
| `--events` | Emit NDJSON events from `sync`. See the [event contract](events.md). Do not combine with `--json`. |
| `--read-only` | Refuse commands that change local or remote data, with exit code `4`. Read-only commands remain available. |
| `-h`, `--help` | Show help. |

Message lookups and search read the selected account's local database and work offline. `sync`, `history fetch`, authentication checks and sends contact the Teams service. By default each account's database is `~/.teamscli/accounts/<account>/messages.db`.

Every command can return `0` on success, `1` for an unexpected internal error and `2` for invalid usage or global options. Each command below lists other relevant outcomes. The full [exit-code table](#exit-codes) is defined in [`src/errors.ts`](../src/errors.ts).

## Accounts and authentication

### `accounts add`

- **Usage:** `teamscli accounts add <name>`
- **Options:** Account name is positional; no command-specific options.
- **Example:** `teamscli accounts add assistant`
- **Changes:** Yes. Creates a local account directory; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `4` when blocked by `--read-only`.

### `accounts list`

- **Usage:** `teamscli accounts list`
- **Options:** No command-specific options.
- **Example:** `teamscli --json accounts list`
- **Changes:** No. Reads local account and secret-store metadata.
- **Exit codes:** `0`, `1`, `2`.

### `accounts remove`

- **Usage:** `teamscli accounts remove <name> --yes`
- **Options:** `--yes` confirms removal.
- **Example:** `teamscli accounts remove demo-only --yes`
- **Changes:** Yes. Deletes that account's local database and saved secret; blocked by `--read-only`. Treat this as destructive.
- **Exit codes:** `0`, `1`, `2`, `4` when blocked by `--read-only`, `5` if the account does not exist.

### `accounts use`

- **Usage:** `teamscli accounts use <name>`
- **Options:** Account name is positional; no command-specific options.
- **Example:** `teamscli accounts use assistant`
- **Changes:** Yes. Changes the saved default account; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `4` when blocked by `--read-only`, `5` if the account does not exist.

### `auth`

- **Usage:** `teamscli auth`
- **Options:** No command-specific options. Use `--account NAME` to select the account.
- **Example:** `teamscli --account assistant auth`
- **Changes:** Yes. Completes device-code sign-in and saves the refresh token in the selected secret store; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `3`, `4`, `7`, `8`.

### `auth status`

- **Usage:** `teamscli auth status`
- **Options:** No command-specific options.
- **Example:** `teamscli --account assistant --read-only auth status`
- **Changes:** Normally checks the saved login and may save a rotated refresh token. With `--read-only`, rotation is not saved.
- **Exit codes:** `0`, `1`, `2`, `3`, `7`, `8`.

### `auth logout`

- **Usage:** `teamscli auth logout`
- **Options:** No command-specific options.
- **Example:** `teamscli --account assistant auth logout`
- **Changes:** Yes. Removes the saved login; local message data is kept. Blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `4`.

### `doctor`

- **Usage:** `teamscli doctor`
- **Options:** No command-specific options.
- **Example:** `teamscli --account assistant --read-only --json doctor`
- **Changes:** Checks the local installation and probes login and Teams service access. It may save a rotated refresh token unless `--read-only` is set; read-only mode still performs remote checks and does not save token rotation.
- **Exit codes:** `0`, `1`, `2`, `3`, `5`, `7`, `8`.

## Local conversations and messages

### `chats list`

- **Usage:** `teamscli chats list [--type chat|group|meeting|channel|system] [--unread] [--limit N]`
- **Options:** `--type` filters by conversation kind; `--unread` keeps only conversations with unread messages; `--limit` caps the number of rows. System conversations are omitted unless requested. Each item carries `readUpTo` and `unreadCount` (see [Unread messages](#unread-messages)).
- **Example:** `teamscli chats list --unread`
- **Changes:** No. Reads local data and works offline.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `teams list`

- **Usage:** `teamscli teams list`
- **Options:** No command-specific options.
- **Example:** `teamscli --json teams list`
- **Changes:** No. Reads locally stored teams.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `channels list`

- **Usage:** `teamscli channels list [--team ID]`
- **Options:** `--team ID` filters by the locally stored team ID.
- **Example:** `teamscli channels list --team '19:example-team@thread.tacv2'`
- **Changes:** No. Reads locally stored channels.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `messages list`

- **Usage:** `teamscli messages list (--chat ID | --unread [--chat ID]) [--limit 50] [--before ISO] [--after ISO]`
- **Options:** `--chat` is required unless `--unread` is given; `--unread` lists only unread messages, across all chats when `--chat` is omitted (human output then prefixes each line with the chat); `--limit` must be positive; `--before` and `--after` are ISO 8601 timestamps.
- **Example:** `teamscli sync && teamscli --read-only --json messages list --unread`
- **Changes:** No. Reads local messages and works offline.
- **Exit codes:** `0`, `1`, `2` (also `--unread` with a channel), `5` if the local database does not exist yet.

### Unread messages

Teams reports, per conversation, the arrival time of the last message the user read (`readUpTo`). A message is unread when it was created after `readUpTo`, was not sent by the signed-in user, is not deleted, and is text, HTML or a card (calls, recordings and system events do not count). Only chats, group chats and meetings are covered; channels and conversations without `readUpTo` always have `unreadCount` 0. The state is the one of the last `sync`: run it first for fresh results. The CLI does not mark messages as read.

### `messages search`

- **Usage:** `teamscli messages search <query> [--chat ID] [--from X] [--after ISO] [--before ISO] [--limit 20]`
- **Options:** Optional filters restrict by chat, sender, time range and result count.
- **Example:** `teamscli --read-only --json messages search "Alice Example" --after 2026-01-01T00:00:00Z`
- **Changes:** No. Searches locally stored message text and works offline.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `messages show`

- **Usage:** `teamscli messages show <id> --chat ID`
- **Options:** Message ID is positional; `--chat` is required because IDs are unique only within a conversation.
- **Example:** `teamscli messages show 1767261600000 --chat '19:example-chat@thread.v2'`
- **Changes:** No. Reads one locally stored message.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database or message does not exist.

### `messages thread`

- **Usage:** `teamscli messages thread <rootId> --channel ID`
- **Options:** Root message ID is positional; `--channel` selects the locally stored channel.
- **Example:** `teamscli --json messages thread 1767261600000 --channel '19:example-channel@thread.tacv2'`
- **Changes:** No. Reads the locally stored channel root and replies.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `history coverage`

- **Usage:** `teamscli history coverage [--chat ID]`
- **Options:** `--chat ID` limits the report to one conversation.
- **Example:** `teamscli history coverage --chat '19:example-chat@thread.v2'`
- **Changes:** No. Reads local coverage metadata and works offline.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

### `store stats`

- **Usage:** `teamscli store stats`
- **Options:** No command-specific options.
- **Example:** `teamscli --json store stats`
- **Changes:** No. Reads local store counts and the selected database path.
- **Exit codes:** `0`, `1`, `2`, `5` if the local database does not exist yet.

## Sync and history

### `sync`

- **Usage:** `teamscli sync [--days N] [--follow] [--interval S]`
- **Options:** `--days` sets the initial recent-history window (default `90`); `--follow` polls continuously; `--interval` sets the poll interval in seconds (default `30`) and requires `--follow`.
- **Example:** `teamscli --account assistant --events sync --follow --interval 30`
- **Changes:** Yes. Reads Teams and writes conversations, messages and sync checkpoints locally; blocked by `--read-only`. `--events` streams NDJSON; it cannot be combined with `--json`.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `7`, `8`.

### `history fetch`

- **Usage:** `teamscli history fetch --chat ID [--before ISO] [--pages N]`
- **Options:** `--chat` is required; `--before` sets the cutoff timestamp; `--pages` limits pages fetched.
- **Example:** `teamscli history fetch --chat '19:example-chat@thread.v2' --before 2026-01-01T00:00:00Z --pages 3`
- **Changes:** Yes. Fetches older messages from Teams and writes them to the local store; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `5`, `7`, `8`.

## Sending and changing messages

### `send text`

- **Usage:** `teamscli send text (--to TARGET | --chat ID | --channel ID) --message TEXT [--reply-to ID] [--quote ID] [--subject TEXT] [--link-previews] [--preview-link URL] [--image PATH] [--link-icons]`
- **Options:** Choose exactly one destination. `--to` resolves a local chat or channel name; `--reply-to` is for channel threads; `--quote` is for chats. `--subject` adds the large title to a new channel post. `--link-previews` adds one card (a live post showed only the first card); it defaults to the first HTTP(S) link. `--preview-link URL` chooses another link and requires `--link-previews`. The card image is `https://<region>-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=<encoded og image>`, using the account's `regionGtms.urlpV2` host. `--image PATH` uploads a local PNG, JPEG, GIF or WebP inline after the message text; repeat it for multiple images. Markdown images (`![alt](url)` or `![alt](file.png)`) are uploaded the same way, because an external image URL does not render. `--link-icons` downloads a site's favicon (apple-touch-icon, then `<link rel=icon>`, then `/favicon.ico`) and adds it as a small inline image before each link; icons are fetched once per domain and a failed icon leaves the link intact.
- **Example:** `teamscli send text --chat '19:example-chat@thread.v2' --quote 1767261600000 --message 'Thanks, Alice Example.'`
- **Changes:** Yes. Sends remotely and updates the local store if the message is found in the first page; if accepted but not found, the CLI reports its `clientMessageId` and leaves the store unchanged. Blocked by `--read-only`. New channel posts support `--subject`. Preview lookup calls the authenticated URLP info endpoint, then public Open Graph metadata; the thumbnail URL is the account's URLP `Thumbnail` endpoint with the image URL encoded once. Icon and page requests use the public web without Teams credentials. Whether a card, its thumbnail, a favicon and an inline image render in the Teams client is not re-checked here: only a live post can confirm that. Chat quotes are currently unverified by Teams.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `5`, `6`, `7`, `8`. An ambiguous `--to` returns candidates in JSON error details.
- **Tip:** The text is markdown and arrives formatted, as if typed in Teams: `**bold**`, `*italic*`, `` `code` ``, links, `#` headings, `>` quotes, `-` and `1.` lists (nested by indentation), `|` tables, `---` rules and fenced code blocks with a language (shown with highlighting and a copy button). `![alt](https://…)` and `![alt](./file.png)` are uploaded and shown inline. Anything else, including raw HTML, is sent as literal text. Prefer one structured message over several plain ones when reporting to people.

### `send file`

- **Usage:** `teamscli send file (--to TARGET | --chat ID) --file PATH [--message TEXT]`
- **Options:** Choose exactly one destination. Images (`png`, `jpg`, `jpeg`, `gif`, `webp`) are sent inline, like a pasted image; any other file is uploaded to the sender's OneDrive (the folder Teams uses for chat files) and attached. `--message` adds markdown text above it.
- **Example:** `teamscli send file --chat '19:example-chat@thread.v2' --file ./report.pdf --message 'Monthly report'`
- **Tip:** Teams opens PDF, Markdown and HTML attachments in its own viewer for whoever receives them, so a file can carry what a chat message cannot: a full report, long Markdown documentation, or an interactive HTML page (charts, a visual diff, a prototype) generated by an agent. Send the file with a one-line `--message` saying what it is.
- **Changes:** Yes. Uploads and sends remotely and updates the local store when the message is found in the first page; blocked by `--read-only`. Chat members in the same organization get read access to an attached file without an email; members from outside do not, and the CLI warns about them. Channels are not supported yet. Sending a file needs membership in at least one team (that is how the CLI finds the OneDrive).
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `5`, `6`, `7`, `8`.

### `messages edit`

- **Usage:** `teamscli messages edit <id> --chat ID --message TEXT`
- **Options:** Message ID is positional; `--chat` and `--message` are required. Only a locally stored message sent by the signed-in account can be edited.
- **Example:** `teamscli messages edit 1767261600000 --chat '19:example-chat@thread.v2' --message 'Updated text'`
- **Changes:** Yes. Edits remotely and updates the local store; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `7`, `8`.

### `messages delete`

- **Usage:** `teamscli messages delete <id> --chat ID --yes`
- **Options:** Message ID is positional; `--chat` and the confirmation flag `--yes` are required. Only a locally stored message sent by the signed-in account can be deleted.
- **Example:** `teamscli messages delete 1767261600000 --chat '19:example-chat@thread.v2' --yes`
- **Changes:** Yes. Requests a remote soft-delete and updates the local store; blocked by `--read-only`.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `7`, `8`.

### `files download`

- **Usage:** `teamscli files download <id> --chat ID [--out DIR]`
- **Options:** Message ID is positional; `--chat` is required. Saves every inline image (`image-N.<ext>`) and attached file (its own name) of a locally stored message into `--out`, or `media/<message id>` in the account directory. Messages synced before attachments were stored need `teamscli history fetch --chat ID` first.
- **Example:** `teamscli files download 1767261600000 --chat '19:example-chat@thread.v2' --out ./downloads`
- **Changes:** Yes. Writes files locally (mode 0600) and never overwrites an existing one; blocked by `--read-only`. Only `sync` and `messages list` never download: they show `[image]` and `[file: name]` placeholders.
- **Exit codes:** `0`, `1`, `2`, `3`, `4`, `5`, `6`, `7`, `8`. Files on another organization's SharePoint or without access fail with `7`.

## Exit codes

These codes are defined in [`src/errors.ts`](../src/errors.ts) and are part of the CLI contract:

| Code | Name | Meaning |
|---:|---|---|
| `0` | — | Command succeeded. |
| `1` | `internal` | Unexpected error. Set `TEAMSCLI_DEBUG=1` to print a stack trace to stderr. |
| `2` | `usage` | Invalid command, option or argument. |
| `3` | `auth_required` | Sign-in is missing or must be renewed; run `teamscli auth`. |
| `4` | `read_only` | A data-changing command was refused by `--read-only`. |
| `5` | `not_found` | The requested local resource was not found. |
| `6` | `ambiguous` | A name matches multiple local targets; choose an ID. |
| `7` | `remote` | A remote service request failed. |
| `8` | `rate_limited` | The remote service rate-limited the request. |
