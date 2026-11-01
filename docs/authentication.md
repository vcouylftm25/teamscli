# Authentication and validated endpoints

Results of read and send spikes against the Teams web service. Times are local (BRT, UTC−3) with UTC alongside. Real ids, names and tokens are omitted.

## Device code — validated on 2026-10-26 (22:10 BRT / 01:10 UTC)

Read-only spike on a work account. Tokens were kept in memory only.

| Step | Request | Result |
|---|---|---|
| Sign-in | device code, client `1fec8e78-bce4-4aaf-ab1b-5451cc387264`, authority `organizations`, scope `https://api.spaces.skype.com/.default offline_access openid profile` | access + refresh token |
| Skype token | `POST https://teams.microsoft.com/api/authsvc/v1.0/authz` with Bearer (spaces) | 200: `tokens.skypeToken`, `region`, `regionGtms.chatService`, `regionGtms.chatServiceAggregator` |
| Chats | `GET {chatService}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=N`, header `Authentication: skypetoken=…` | 200 |
| Messages | `GET {chatService}/v1/users/ME/conversations/{id}/messages?view=msnp24Equivalent&pageSize=N` | 200; `_metadata.backwardLink` and `_metadata.syncState` |
| Chats (alternative) | same endpoint with `Authorization: Bearer` for `https://ic3.teams.office.com/.default` | 200 |
| Teams and channels | `GET https://teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me?isPrefetch=false&enableMembershipSummary=true`, Bearer for `https://chatsvcagg.teams.microsoft.com/.default` | 200: `teams[].channels`, `chats` |
| Renewal | refresh token for chatsvcagg, ic3 and Graph | ok, no interaction |
| Inline image | `GET https://{host}.asm.skype.com/v1/objects/{id}/views/imgo`, header `Authorization: skype_token …` (`Authentication: skypetoken=` and no header: 401) | 200, the image (spike 2026-10-29) |
| Attached file | refresh token for `https://{tenant}-my.sharepoint.com/.default`; `GET {site}/_api/v2.0/shares/u!{base64url(objectUrl)}/driveItem` with that Bearer, then `GET @content.downloadUrl` | 200, the file (spike 2026-10-29; another tenant's SharePoint not tested) |
| Upload image | `POST https://us-api.asm.skype.com/v1/objects` with `Authorization: skype_token …` and a User-Agent containing `Teams/<version>` (browser or Node UA: 400 "platform id"), body `{type: "pish/image", permissions: {"<conversationId>": ["read"]}, filename}`; then `PUT /v1/objects/{id}/content/imgpsh` with the bytes; the message carries `<img itemtype="http://schema.skype.com/AMSImage" …>` and `amsreferences: [id]` | 201; shown inline (spike 2026-10-29) |
| Upload file | OneDrive host from csa `teams[].teamSiteInformation.sharepointSiteUrl` + `-my`; `PUT {my}/_api/v2.0/me/drive/root:/<chat files folder>/{name}:/content`; `POST …/items/{id}/invite` `{recipients: [{objectId}], roles: ["read"], requireSignIn: true, sendInvitation: false}`; message with `content` "" and `properties.files` (shape in `src/teams/files.ts`) | 201; a colleague opened the file. `createLink` with scope `users` is refused (400, only `organization`/`anonymous`) |

`threadType` values seen in conversations: `chat`, `meeting`, `space`, `topic`, `streamofcalllogs`, `streamofnotifications`, `streamofmentions`, `streamofnotes`.

## Channels, threads and sending — validated on 2026-10-26 (22:25 BRT / 01:25 UTC)

| Step | Request | Result |
|---|---|---|
| Channel messages | `GET {chatService}/v1/users/ME/conversations/{19:…@thread.tacv2}/messages?view=msnp24Equivalent&pageSize=N` | 200; the feed has roots and replies |
| Thread | `GET {chatService}/v1/users/ME/conversations/{channel};messageid={root}/messages` | 200 with the root and its replies |
| Send | `POST {chatService}/v1/users/ME/conversations/48:notes/messages`, body `{content, messagetype: "Text", contenttype: "text", clientmessageid}` | 201 `OriginalArrivalTime`; message found in history by `clientmessageid` |

- Root vs reply: `conversationLink` ends in `;messageid=<id>`. On a root the id is its own; on a reply, the root's.
- `48:notes` is the chat with yourself (`threadType` `streamofnotes`), used as the test target.

## Replies — validated on 2026-10-27 (08:33 BRT / 11:33 UTC)

| Step | Request | Result |
|---|---|---|
| Reply in a channel thread | `POST {chatService}/v1/users/ME/conversations/{channel};messageid={root}/messages`, body `{content: "<p>…</p>", messagetype: "RichText/Html", contenttype: "text", imdisplayname, clientmessageid}` | 201; shows up in the thread |
| Quoted reply in a chat | same POST on the chat, with the `content` below | 201; Teams shows the quote with the warning "This quoted message was not verified by Teams" (in every combination tested, see "Quotes: the warning") |

```html
<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="{msgId}">
  <strong itemprop="mri" itemid="8:orgid:{oid}">{name}</strong>
  <span itemprop="time" itemid="{msgId}"></span>
  <p itemprop="preview">{quoted text}</p>
</blockquote><p>{reply}</p>
```

On messages quoted by the client, `properties.qtdMsgs` is an object (`[{messageId, sender, time, validationResult: "Valid", …}]`) and comes with `hasValidMsgReferences: true`. Sending `qtdMsgs` as a JSON string stores the property but does not validate it. *(Not reproduced in the 2026-10-27 capture: the client sent a JSON string and the server stored it as such; see "Quotes, mentions, channel posts and reply edits".)*

## Pagination, incremental sync, edit and delete — validated on 2026-10-27 (09:04 BRT / 12:04 UTC)

- **Conversation list:** `GET {chatService}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=100`. `_metadata` has `totalCount`, `forwardLink`, `backwardLink` and `syncState`. Following `backwardLink` until an empty page gave 3 pages and 197 conversations, with no repeats.
- **History:** `.../messages?view=msnp24Equivalent&pageSize=N&startTime=1`, following `_metadata.backwardLink`. `pageSize=200` is accepted. On an 86-message chat, `pageSize=20` gave 5 pages with no duplicates; the end shows up as an empty page or a missing link.
- **Order:** pages are not strictly ordered by `id`, `sequenceId`, `composetime` or `version`. Messages changed later (`version` changes on every change, not only on edit) can appear out of place. The local database must upsert by `id` and sort locally.
- **`syncState`:** the URL for incremental sync. Starting from a partial page it first delivers the rest of the history; once it reaches the present it returns only what changed: new, edited or deleted messages.
- **Edit:** `PUT {chatService}/v1/users/ME/conversations/{conv}/messages/{id}` with the same body as the send → 200. The message keeps its `id`, gets `properties.edittime` and a new `version`, and shows up in `syncState`.
- **Delete:** `DELETE .../messages/{id}?behavior=softDelete` → 200. The message keeps its `id`, with empty `content` and `properties.deletetime`, both in `syncState` and in the full listing.

## Quotes, mentions, channel posts and reply edits — validated on 2026-10-27 (17:05 BRT / 20:05 UTC)

Traffic capture of Teams web (`teams.cloud.microsoft`) on a work account, sending only to the notes chat and to a test channel; everything was deleted afterwards. Real values replaced by `<conversationId>`, `<channelId>`, `<messageId>`, `<mri>`.

### What the web client does differently from the CLI

| | Teams web | CLI today |
|---|---|---|
| Host | `https://teams.cloud.microsoft/api/chatsvc/<region>/v1/users/ME/...` (same paths as `chatService`) | `{chatService}/v1/users/ME/...` |
| Auth | `Authorization: Bearer` (audience `ic3.teams.office.com`) | `Authentication: skypetoken=…` |
| Other headers | `behavioroverride: redirectAs404`, `x-ms-migration: True`, `x-ms-request-priority`, `clientinfo` | only auth and `content-type` |
| `;` in thread id | sent as `%3B` (`<channelId>%3Bmessageid%3D<rootId>`) | literal `;` |

The client's POST body is always the same envelope (plain text, reply, quote and channel post):

```json
{
  "id": "-1", "type": "Message",
  "conversationid": "<conversationId>", "conversationLink": "blah/<conversationId>",
  "from": "<mri>", "fromUserId": "<mri>",
  "composetime": "<ISO>", "originalarrivaltime": "<ISO>",
  "content": "<p>text</p>", "messagetype": "RichText/Html", "contenttype": "Text",
  "imdisplayname": "<name>", "clientmessageid": "<digits>",
  "callId": "", "state": 0, "version": "0", "amsreferences": [],
  "properties": { "importance": "", "subject": "", "title": "", "cards": "[]", "links": "[]",
                  "mentions": "[]", "onbehalfof": null, "files": "[]", "policyViolation": null,
                  "formatVariant": "TEAMS" },
  "crossPostChannels": []
}
```

POST response: `201` with `{"OriginalArrivalTime": <ms>}`; that value is the new message's id, also in the `Location` header (`…/messages/<messageId>`).

### Results per action

| Action | Request | Body differences | Result |
|---|---|---|---|
| Quote in a chat | `POST …/conversations/<conversationId>/messages` | `content` with the blockquote from "Replies"; `properties.qtdMsgs` (below) | 201 |
| New channel post | `POST …/conversations/<channelId>/messages`, **without** `;messageid` | `"postType": "Standard"`; `properties.subject` is `""` without a subject | 201; `threadtype: topic`; the new id is the root |
| Reply to a channel post | `POST …/conversations/<channelId>;messageid=<rootId>/messages` | `conversationid` = `<channelId>`, `conversationLink` = `blah/<channelId>;messageid=<rootId>`, `"replyChainId": "<rootId>"`, no `postType` | 201 |
| Edit a channel reply | `PUT …/conversations/<channelId>/messages/<messageId>`, **without** `;messageid` | the whole envelope with the real `id`, `version` in ISO, a new `clientmessageid`, `replyChainId`, `properties.edittime` (ms) | 200, empty body |
| Delete a channel reply | `DELETE …/conversations/<channelId>/messages/<messageId>?behavior=softDelete` | no body | 200 |
| Delete a post (root) | same `DELETE`, with the root id | no body | 200 |

- Edit and delete in a channel use the **channel** id, without `;messageid=`, for both root and reply. The CLI already does this (stored `conversationId`); the shortcut in `src/teams/send.ts` is confirmed.
- The client does not ask for confirmation on delete.
- The same `DELETE` with `Authentication: skypetoken` (the CLI's client) also worked: 4 test messages deleted in the notes chat.

### Quotes: the "not verified" warning (tested on 2026-10-27, 17:40–17:55 BRT / 20:40–20:55 UTC)

The client sends `properties.qtdMsgs` as a **JSON string**, with `messageId` as a string and `time` as a number (ms):

```json
"qtdMsgs": "[{\"messageId\":\"<messageId>\",\"sender\":\"<mri>\",\"time\":<ms>}]"
```

The `content` is the same blockquote documented above (the client writes `itemscope` with no value; the CLI's `itemscope=""` is the same HTML). The server stores the string as sent: neither `hasValidMsgReferences` nor `validationResult` appear.

**Result: no tested combination removes the warning.** Only a quote sent by Teams web itself shows as verified. In a `48:notes` chat, with a test source message, each variant below quoted the source and was checked in the UI as soon as it appeared (only c5 and c6 were not re-checked after reload; the others, the native one and the one without `qtdMsgs` were also checked after reload):

| Variant (`qtdMsgs` string and client blockquote unless noted) | POST | Warning |
|---|---|---|
| control: `skypetoken`, `{chatService}`, CLI envelope | 201 | yes |
| a) `Authorization: Bearer` `ic3` (`refreshAccessToken`), `{chatService}`, CLI envelope | 201 | yes |
| b) Bearer `ic3`, client host (`teams.cloud.microsoft/api/chatsvc/<region>`), CLI envelope + client headers | 201 | yes |
| c1) full client envelope, `skypetoken`, `{chatService}` | 201 | yes |
| c2) full envelope, `skypetoken`, client host + client headers | 201 | yes |
| c3) full envelope, Bearer `ic3`, client host + client headers | 201 | yes |
| c4) full envelope, Bearer `ic3`, `{chatService}` | 201 | yes |
| c5) as c3 + browser `origin`, `referer`, `user-agent` | 201 | yes |
| c6) as c5, preceded by the `Control/Typing` POST | 201 | yes |
| current CLI envelope and body (blockquote, **without** `qtdMsgs`), `skypetoken` | 201 | yes |
| native: "Reply with quote" in the Teams web UI | 201 | **no**, still no warning after reload |

- The native POST body is identical to c3 (only `itemscope` without `=""`) and it makes no validation call: besides routine reads (sync, `consumptionhorizon`), only the `Control/Typing` POST and the message POST.
- Checked with `GET …/messages/<messageId>` (with `skypetoken`, and with Bearer on the client host, with `view=msnp24Equivalent|supportsMessageProperties`): the native quote and variant c3 have the same set of fields and `properties` (`cards`, `files`, `formatVariant`, `languageStamp`, `links`, `mentions`, `qtdMsgs`), with equal `qtdMsgs`. Nothing the server returns tells them apart.
- The only differences not reproduced: the token (the web client's is issued to the Teams web broker app, a different `appid` from the CLI's, with `scp` `Teams.AccessAsUser.All` and MFA in `amr`) and browser cookies/local state. Open hypotheses: (1) the server flags the message by the token's app without exposing it on read; (2) the client stores locally, per message, that it sent the quote itself. To tell them apart: open the chat with the native quote in another client (mobile or desktop app). If the warning shows there, it is local state; if not, it is a server flag.
- For the CLI: quoting works (the block shows with sender, time and preview) but Teams shows the warning. Switching to Bearer, the client host or the full envelope gains nothing; `qtdMsgs` does not change the warning either. Keep the current send and document the limitation.

Other useful observations:

- The `ic3` Bearer requested by the CLI (`refreshAccessToken` with `https://ic3.teams.office.com/.default`) is accepted on message POST and GET, on both `{chatService}` and the client host. `skypetoken` is also accepted on the client host (c2).
- Before each message the client sends a typing POST to the same endpoint: `{"content":"","contenttype":"Application/Message","messagetype":"Control/Typing"}` → 201 `{"OriginalArrivalTime": <ms>}`. Not required to send.
- Client POST headers: `x-ms-request-priority` (`10` for typing, `0` for the message), `behavioroverride: redirectAs404`, `x-ms-test-user: False`, `x-ms-migration: True`, `clientinfo`, `content-type: application/json`.
- 15 back-to-back `DELETE`s got 429 on 6; with 1.5 s between them and retry with backoff, all passed.

### Mentions (validated on 2026-10-27, 17:26–17:40 BRT / 20:26–20:40 UTC)

Captured in a 1:1 chat with an authorized test person (picked in the `@` picker); the test messages were deleted. Teams does not offer the signed-in user in the `@` picker (neither in the notes chat nor in a channel), so another person is needed.

Client POST (full envelope from "What the web client does differently from the CLI", changing only `content` and `properties.mentions`):

```json
"content": "<p>mention test <readonly class=\"skipProofing\" itemtype=\"http://schema.skype.com/Mention\" spellcheck=\"false\"><span itemtype=\"http://schema.skype.com/Mention\" itemscope itemid=\"0\"><name></span></readonly></p>",
"properties": {
  "mentions": "[{\"@type\":\"http://schema.skype.com/Mention\",\"itemid\":\"0\",\"mri\":\"<mri>\",\"mentionType\":\"person\",\"displayName\":\"<name>\"}]",
  …
}
```

- `properties.mentions` is a **JSON string** (like `qtdMsgs`), one object per mention: fixed `@type`, `itemid` as a string, `mri`, `mentionType: "person"`, `displayName`. The object's `itemid` matches the `<span>`'s (`"0"` for the first mention; several mentions with increasing `itemid` were not tested).
- In `content` the span holds only the name, with no `@`; the user's text stays outside `<readonly>`. `messagetype` is `RichText/Html`.
- Response: 201 `{"OriginalArrivalTime": <ms>}`; the message showed in the UI with the name highlighted (`span data-mention-type="person"`, a "<name> mentioned" button).
- The CLI's minimal body works: `{content, messagetype: "RichText/Html", contenttype: "text", clientmessageid, properties: {mentions}}` with `skypetoken` and the same `content` → 201, and the UI renders it exactly like the native message (same `span`, same color, same button). `imdisplayname` and the rest of the envelope are not needed.
- Not verified: whether the recipient gets the mention notification (the message was deleted right after), and mentions of channels, teams or groups.

## Session persistence

Refresh token encrypted with DPAPI (`CurrentUser`) through `powershell.exe`. The secret travels only over stdin and stdout, never as an argument. The blob lives in `~/.teamscli/accounts/<account>/refresh_token.dpapi`, mode 600. WSL has no keyring.

Validated: a new process read the blob, renewed the token and saved the rotated refresh token. Each DPAPI operation takes about 0.5 s.

## Pending

- Real session lifetime: the response has no `refresh_token_expires_in`. Re-run with the saved token after days without use.
- Verified quotes: no solution. 10 combinations of auth, host, envelope, headers and `qtdMsgs` tested (see "Quotes: the warning"); none removes the warning, and the server stores the same content as for the native quote. Still to separate "server flag per app" from "client local state": view the native quote in another client (mobile or desktop). If it is a per-app flag, only the Teams web app verifies and the CLI documents the limitation.
- Mentions: send format confirmed and reproduced by the CLI; `send text --mention` implemented. Still to confirm the recipient notification, several mentions in one message (increasing `itemid`?) and mentions of channels, teams or groups.
- Channel post subject: the CLI sends `properties.subject` on a new channel post (the field is present, empty, in the client envelope). Not checked in the Teams UI.
- Link preview cards, favicon icons and uploaded markdown images: the send body matches captured message shape (`properties.links`, AMS images). A live post showed favicons drawn inside a large grey placeholder box, so `--radar` uses text emojis (`--link-emojis`) instead.
- Real time: Teams web receives events over a websocket. Not investigated yet; polling `syncState` works as an alternative.
- Behavior under rate limiting during an initial sync with real volume.

## Known limitations

- Some embedded browsers do not support WebAuthn: passkey sign-in fails with `AADSTS10040141`. Complete the device code in a regular browser, including a phone's.
- Device code with the Teams client can be blocked by Conditional Access or raise a security alert, depending on tenant policy.
