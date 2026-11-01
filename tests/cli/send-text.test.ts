import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { AAD_TOKEN_URL, AUTHZ_ROUTE, fakeJwt, TEST_OID, TEST_TID, tokenResponse } from '../helpers/fake-auth.ts';
import { runTestCli } from '../helpers/cli.ts';
import { fakeFetch, type FakeRoute } from '../helpers/fake-fetch.ts';
import { aConversation, aMessage } from '../helpers/builders.ts';
import { databasePath } from '../../src/runtime/paths.ts';
import { openDatabase } from '../../src/store/database.ts';
import { upsertConversations } from '../../src/store/conversations.ts';
import { upsertPeople } from '../../src/store/people.ts';
import { shortId } from '../../src/store/resolve.ts';
import { getMessage, upsertMessages } from '../../src/store/messages.ts';

const env = { TEAMSCLI_SECRET_STORE: 'file' };
const chatService = 'https://chat.example.test';

test('send text help documents preview and inline image options', async () => {
  const run = await runTestCli(['send', 'text', '--help']);
  assert.equal(run.code, 0);
  assert.match(run.stdout, /--link-previews\s+Attach one link preview card\./);
  assert.match(run.stdout, /--link-icons\s+Add each site favicon as a small inline image before its links/);
  assert.match(run.stdout, /--link-emojis\s+Put a small emoji before each link/);
  assert.match(run.stdout, /--link-sources\s+Show the site/);
  assert.match(run.stdout, /--compact\s+Tighter spacing/);
  assert.match(run.stdout, /--radar\s+Newsletter preset: --compact, --link-emojis and --link-sources/);
  assert.match(run.stdout, /--preview-link URL\s+Choose which link gets the single preview card/);
  assert.match(run.stdout, /--image PATH\s+Upload an inline image; may be repeated\./);
});

test('--radar puts an emoji and the grey source on each link and never fetches or uploads a favicon', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-radar';
  seed(home, [aConversation({ id: chatId, title: 'Radar group' })]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600030', clientMessageId));

  const text = '[Alerta](https://www.cisa.gov/news-events/alerts/x) e [repo](https://github.com/acme/tool) e [post](https://example.com/page)';
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId, '--radar', '--message', text], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.equal(fake.calls.some(({ url }) => /favicon|asm\.skype\.com|cisa\.gov|github\.com\/acme|example\.com/.test(url)), false, 'no icon download or AMS upload');
  const body = chatCalls(fake.calls)[0]?.body as Record<string, unknown>;
  assert.equal('amsreferences' in body, false);
  assert.equal(String(body.content).includes('<img'), false);
  assert.equal(body.content, '<p>🛡️&nbsp;<a href="https://www.cisa.gov/news-events/alerts/x">Alerta</a> <span style="font-size:12px;color:#616161">· cisa.gov</span> e '
    + '💻&nbsp;<a href="https://github.com/acme/tool">repo</a> <span style="font-size:12px;color:#616161">· github.com</span> e '
    + '🔗&nbsp;<a href="https://example.com/page">post</a> <span style="font-size:12px;color:#616161">· example.com</span></p>');
});

test('send text resolves --to, escapes user HTML, and stores the returned message', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-send';
  seed(home, [aConversation({ id: chatId, title: 'Example group' })]);
  const base = `${chatService}/v1/users/ME/conversations/${chatId}`;
  const fake = authenticatedFetch([
    { method: 'POST', url: `${base}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600000', clientMessageId));

  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--to', 'Example group', '<img src=x onerror=alert(1)>'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  const sent = JSON.parse(run.stdout).data;
  assert.equal(sent.id, '1767261600000');
  assert.equal(sent.text, 'Sent safely');
  assert.equal(sent.senderName, 'Alex Example');
  const remoteCalls = chatCalls(fake.calls);
  assert.match(remoteCalls[0]?.url ?? '', /\/conversations\/chat-send\/messages$/);
  const body = remoteCalls[0]?.body as Record<string, unknown>;
  assert.equal(body.content, '<p>&lt;img src=x onerror=alert(1)&gt;</p>');
  assert.doesNotMatch(String(body.content), /<img/);
  assert.match(String(body.clientmessageid), /^\d+$/);
  assert.equal('imdisplayname' in body, false);
  const db = openDatabase(databasePath(home, 'work'));
  try {
    assert.equal(getMessage(db, chatId, sent.id)?.clientMessageId, body.clientmessageid);
  } finally {
    db.close();
  }
});

test('send text renders a plain-text detail, accepts positional text, and reports missing inputs', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-human-send';
  seed(home, [aConversation({ id: chatId, title: 'Example group' })]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600020', clientMessageId));

  const sent = await runTestCli(['--store', home, '--account', 'work', 'send', 'text', '--to', 'Example group', 'Hello <there>'], { env, fetch: fake.fetch });
  assert.equal(sent.code, 0, sent.stderr);
  assert.match(sent.stdout, /id\s+1767261600020/);
  assert.match(sent.stdout, /text\s+Sent safely/);
  assert.doesNotMatch(sent.stdout, /<p>|<\/p>/);

  const missingMessageFetch = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ]);
  const notFound = await runTestCli(['--store', home, '--account', 'work', 'send', 'text', '--to', 'Example group', 'Still sent'], { env, fetch: missingMessageFetch.fetch });
  assert.equal(notFound.code, 0, notFound.stderr);
  assert.match(notFound.stdout, /sent\s+true/);
  assert.match(notFound.stdout, /clientMessageId\s+\d+/);
  assert.doesNotMatch(notFound.stdout, /text\s+/);

  for (const [args, expected] of [
    [['send', 'text'], 'Missing the destination: pass --to TARGET, --chat ID or --channel ID (id, short id or name).'],
    [['send', 'text', '--chat', 'Alice Example'], 'Missing the message text: pass it after the options or with --message.'],
  ] as const) {
    const run = await runTestCli(['--store', home, ...args]);
    assert.equal(run.code, 2, args.join(' '));
    assert.equal(run.stderr, 'teamscli: ' + expected + '\n');
  }

  const both = await runTestCli(['--store', home, 'send', 'text', '--chat', 'Alice Example', 'positional', '--message', 'option']);
  assert.equal(both.code, 2);
  assert.equal(both.stderr, 'teamscli: Pass the message text either after the options or with --message, not both.\n');
});

test('--to preserves ambiguous errors and returns candidate ids', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  seed(home, [
    aConversation({ id: 'chat-one', title: 'Example group' }),
    aConversation({ id: 'chat-two', title: 'Example group' }),
  ]);
  const fake = authenticatedFetch([]);
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--to', 'Example group', '--message', 'Synthetic'], { env, fetch: fake.fetch });

  assert.equal(run.code, 6);
  assert.equal(JSON.parse(run.stdout).error.code, 'ambiguous');
  assert.deepEqual(JSON.parse(run.stdout).error.details.candidates.map((item: { id: string }) => item.id), ['chat-one', 'chat-two']);
  assert.equal(chatCalls(fake.calls).some(({ url }) => url.includes('/conversations/')), false);
});

test('--mention replaces a matching name and sends stringified mention properties', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-mention-middle';
  seed(home, [aConversation({ id: chatId })]);
  seedPeople(home, [{ id: '8:orgid:00000000-0000-0000-0000-000000000021', displayName: 'João Example', email: null }]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600021', clientMessageId));

  const run = await runTestCli([
    '--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId,
    '--mention', 'joao example', '--message', 'Oi @JOÃO EXAMPLE, tudo bem?',
  ], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.text, 'Sent safely');
  const body = chatCalls(fake.calls)[0]?.body as Record<string, unknown>;
  assert.equal(body.content,
    '<p>Oi <readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" spellcheck="false">'
    + '<span itemtype="http://schema.skype.com/Mention" itemscope itemid="0">João Example</span></readonly>, tudo bem?</p>');
  assert.equal((body.properties as Record<string, unknown>).mentions, JSON.stringify([{
    '@type': 'http://schema.skype.com/Mention', itemid: '0', mri: '8:orgid:00000000-0000-0000-0000-000000000021',
    mentionType: 'person', displayName: 'João Example',
  }]));
});

test('--mention inserts a mention at the start when no name occurs in the text', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-mention-prefix';
  seed(home, [aConversation({ id: chatId })]);
  seedPeople(home, [{ id: '8:orgid:00000000-0000-0000-0000-000000000022', displayName: 'Bob Example', email: null }]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ]);

  const run = await runTestCli([
    '--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId,
    '--mention', 'Bob Example', '--message', 'Hello team',
  ], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.match(String((chatCalls(fake.calls)[0]?.body as Record<string, unknown>).content),
    /^<p><readonly class="skipProofing" itemtype="http:\/\/schema\.skype\.com\/Mention" spellcheck="false"><span itemtype="http:\/\/schema\.skype\.com\/Mention" itemscope itemid="0">Bob Example<\/span><\/readonly> Hello team<\/p>$/);
});

test('--mention uses increasing itemids in the order mentions appear in the content', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-mention-order';
  seed(home, [aConversation({ id: chatId })]);
  seedPeople(home, [
    { id: '8:orgid:00000000-0000-0000-0000-000000000023', displayName: 'Alice Example', email: null },
    { id: '8:orgid:00000000-0000-0000-0000-000000000024', displayName: 'Bob Example', email: null },
  ]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ]);

  const run = await runTestCli([
    '--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId,
    '--mention', 'Bob Example', '--mention', 'Alice Example', '--message', 'Hi @Alice Example, then @Bob Example.',
  ], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  const body = chatCalls(fake.calls)[0]?.body as Record<string, unknown>;
  assert.match(String(body.content), /Alice Example<\/span><\/readonly>, then <readonly[^>]*><span[^>]*itemid="1">Bob Example/);
  assert.equal((body.properties as Record<string, unknown>).mentions, JSON.stringify([
    { '@type': 'http://schema.skype.com/Mention', itemid: '0', mri: '8:orgid:00000000-0000-0000-0000-000000000023', mentionType: 'person', displayName: 'Alice Example' },
    { '@type': 'http://schema.skype.com/Mention', itemid: '1', mri: '8:orgid:00000000-0000-0000-0000-000000000024', mentionType: 'person', displayName: 'Bob Example' },
  ]));
});

test('--mention escapes HTML in the resolved display name', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-mention-escape';
  seed(home, [aConversation({ id: chatId })]);
  seedPeople(home, [{ id: '8:orgid:00000000-0000-0000-0000-000000000025', displayName: '**Alice** <Example> & "Co"', email: null }]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ]);

  const run = await runTestCli([
    '--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId,
    '--mention', 'Alice', '--message', 'Hi @Alice',
  ], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  const body = chatCalls(fake.calls)[0]?.body as Record<string, unknown>;
  assert.match(String(body.content), /\*\*Alice\*\* &lt;Example&gt; &amp; &quot;Co&quot;<\/span>/);
  assert.doesNotMatch(String(body.content), /<Example>/);
});

test('--mention reports ambiguous and missing people from the local store', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-mention-resolution';
  seed(home, [aConversation({ id: chatId })]);
  seedPeople(home, [
    { id: '8:orgid:00000000-0000-0000-0000-000000000026', displayName: 'Alice One', email: null },
    { id: '8:orgid:00000000-0000-0000-0000-000000000027', displayName: 'Alice Two', email: null },
  ]);
  const fake = authenticatedFetch([]);
  const args = ['--store', home, '--account', 'work', '--json', 'send', 'text', '--to', chatId, '--message', 'Hello'];

  const ambiguous = await runTestCli([...args, '--mention', 'Alice'], { env, fetch: fake.fetch });
  assert.equal(JSON.parse(ambiguous.stdout).error.code, 'ambiguous');
  assert.equal(chatCalls(fake.calls).some(({ method }) => method === 'POST'), false);

  const missing = await runTestCli([...args, '--mention', 'Nobody'], { env, fetch: fake.fetch });
  assert.equal(JSON.parse(missing.stdout).error.code, 'not_found');
  assert.equal(chatCalls(fake.calls).some(({ method }) => method === 'POST'), false);
});

test('send text accepts a chat name with --chat', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  seed(home, [aConversation({ id: 'chat-direct' })]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/chat-direct/messages`, status: 201, json: {} },
    { url: messagePage('chat-direct'), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage('chat-direct', '1767261600001', clientMessageId));
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--chat', 'Example group', '--message', 'Hello'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.conversationId, 'chat-direct');
  assert.equal(chatCalls(fake.calls)[0]?.method, 'POST');
});

test('send text replies to a channel root even when --reply-to names a reply', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const channelId = '19:channel@example.test';
  seed(home, [aConversation({ id: channelId, kind: 'channel' })], [
    aMessage({ id: 'root-id', conversationId: channelId, senderId: '8:orgid:someone', text: 'Root', content: 'Root' }),
    aMessage({ id: 'reply-id', conversationId: channelId, parentId: 'root-id', senderId: '8:orgid:someone', text: 'Reply', content: 'Reply' }),
  ]);
  const threadId = `${channelId};messageid=root-id`;
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)};messageid=root-id/messages`, status: 201, json: {} },
    { url: messagePage(threadId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(channelId, '1767261600002', clientMessageId, 'root-id'));
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--channel', shortId(channelId), '--reply-to', 'reply-id', '--message', 'Thread reply'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.parentId, 'root-id');
  const remoteCalls = chatCalls(fake.calls);
  assert.match(remoteCalls[0]?.url ?? '', /;messageid=root-id\/messages$/);
  assert.equal(remoteCalls[1]?.url, messagePage(threadId));
});

test('--quote prepends an escaped blockquote and warns that Teams did not verify it', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-quote';
  seed(home, [aConversation({ id: chatId })], [aMessage({
    id: 'quote-id', conversationId: chatId, senderName: '<Bob & Co>', text: 'quoted <text>', content: 'quoted <text>',
  })]);
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ]);
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--chat', chatId, '--quote', 'quote-id', '--message', 'Answer'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  const content = (chatCalls(fake.calls)[0]?.body as Record<string, unknown>).content as string;
  assert.ok(content.startsWith('<blockquote '));
  assert.ok(content.indexOf('</blockquote>') < content.indexOf('<p>Answer</p>'));
  assert.match(content, /&lt;Bob &amp; Co&gt;/);
  assert.match(content, /quoted &lt;text&gt;/);
  assert.doesNotMatch(content, /<Bob|<text>/);
  assert.match(run.stderr, /not verified by Teams/);
});

test('send text rejects a preview link without previews and sends a channel subject', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const channelId = '19:channel-subject@example.test';
  seed(home, [aConversation({ id: channelId, kind: 'channel', title: 'Example channel' })]);
  const missing = await runTestCli(['--store', home, '--account', 'work', 'send', 'text', '--channel', channelId, '--preview-link', 'https://example.com/a', '--message', 'Hi'], { env });
  assert.equal(missing.code, 2);
  assert.match(missing.stderr, /--preview-link requires --link-previews/);

  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)}/messages`, status: 201, json: {} },
    { url: messagePage(channelId), json: { messages: [], _metadata: {} } },
  ]);
  const run = await runTestCli(['--store', home, '--account', 'work', 'send', 'text', '--channel', channelId, '--subject', 'Weekly note', '--message', 'Hello'], { env, fetch: fake.fetch });
  assert.equal(run.code, 0, run.stderr);
  const body = chatCalls(fake.calls)[0]?.body as Record<string, unknown>;
  assert.equal((body.properties as Record<string, unknown>).subject, 'Weekly note');
  assert.equal(body.postType, 'Standard');
});

test('send text uploads a local --image through AMS', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const channelId = '19:channel-image@example.test';
  seed(home, [aConversation({ id: channelId, kind: 'channel' })]);
  const image = join(home, 'banner.png');
  writeFileSync(image, Buffer.from([0x89, 0x50, 0x4e, 0x47, 1]));
  const fake = authenticatedFetch([
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', status: 201, json: { id: '0-cli-image' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-cli-image/content/imgpsh', json: {} },
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)}/messages`, status: 201, json: {} },
    { url: messagePage(channelId), json: { messages: [], _metadata: {} } },
  ]);
  const run = await runTestCli(['--store', home, '--account', 'work', 'send', 'text', '--channel', channelId, '--image', image, '--message', 'Chart'], { env, fetch: fake.fetch });
  assert.equal(run.code, 0, run.stderr);
  const body = chatCalls(fake.calls).find((call) => call.method === 'POST')?.body as Record<string, unknown>;
  assert.deepEqual(body.amsreferences, ['0-cli-image']);
  assert.match(String(body.content), /<p>Chart<\/p><p><img [^>]*itemid="0-cli-image"/);
});

test('send text creates and stores a root channel post', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const channelId = '19:channel-new@example.test';
  seed(home, [aConversation({ id: channelId, kind: 'channel' })]);
  const id = '1767261600012';
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)}/messages`, status: 201, json: { OriginalArrivalTime: id } },
    { url: messagePage(channelId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(channelId, id, clientMessageId));
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--channel', shortId(channelId), '--message', 'Synthetic'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  const sent = JSON.parse(run.stdout).data;
  assert.equal(sent.id, id);
  assert.equal(sent.parentId, null);
  const remoteCalls = chatCalls(fake.calls);
  assert.equal(remoteCalls[0]?.url, `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)}/messages`);
  assert.equal(remoteCalls[1]?.url, messagePage(channelId));
  assert.equal((remoteCalls[0]?.body as Record<string, unknown>).postType, 'Standard');
  const db = openDatabase(databasePath(home, 'work'));
  try {
    assert.equal(getMessage(db, channelId, id)?.parentId, null);
    assert.equal(getMessage(db, channelId, id)?.clientMessageId, (remoteCalls[0]?.body as Record<string, unknown>).clientmessageid);
  } finally {
    db.close();
  }
});

test('--to can resolve a channel for a new root post', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const channelId = '19:channel-to@example.test';
  seed(home, [aConversation({ id: channelId, kind: 'channel', title: 'Example channel' })]);
  const id = '1767261600013';
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)}/messages`, status: 201, json: { OriginalArrivalTime: id } },
    { url: messagePage(channelId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(channelId, id, clientMessageId));
  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--to', 'Example channel', '--message', 'Synthetic'], { env, fetch: fake.fetch });

  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.conversationId, channelId);
  assert.equal((chatCalls(fake.calls)[0]?.body as Record<string, unknown>).postType, 'Standard');
});

test('send text fills a missing senderName from the local identity', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-sender-local';
  seed(home, [aConversation({ id: chatId, title: 'Example group' })]);
  writeFileSync(join(home, 'accounts', 'work', 'account.json'), JSON.stringify({
    userId: `8:orgid:${TEST_OID}`, tenantId: TEST_TID, region: 'test', displayName: 'Local Self',
  }), { mode: 0o600 });
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600010', clientMessageId, undefined, null));

  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--chat', 'Example group', '--message', 'Hello'], { env, fetch: fake.fetch });
  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.senderName, 'Local Self');
});

test('send text fills a missing senderName from the local self person record', async (t) => {
  const home = signedInHome();
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const chatId = 'chat-sender-person';
  seed(home, [aConversation({ id: chatId })]);
  const db = openDatabase(databasePath(home, 'work'));
  upsertPeople(db, [{ id: `8:orgid:${TEST_OID}`, displayName: 'Stored Self', email: null }]);
  db.close();
  const fake = authenticatedFetch([
    { method: 'POST', url: `${chatService}/v1/users/ME/conversations/${chatId}/messages`, status: 201, json: {} },
    { url: messagePage(chatId), json: { messages: [], _metadata: {} } },
  ], (clientMessageId) => rawMessage(chatId, '1767261600011', clientMessageId, undefined, null));

  const run = await runTestCli(['--store', home, '--account', 'work', '--json', 'send', 'text', '--chat', chatId, '--message', 'Hello'], { env, fetch: fake.fetch });
  assert.equal(run.code, 0, run.stderr);
  assert.equal(JSON.parse(run.stdout).data.senderName, 'Stored Self');
});

function signedInHome(): string {
  const home = mkdtempSync(join(tmpdir(), 'teamscli-send-'));
  mkdirSync(join(home, 'accounts', 'work'), { recursive: true });
  writeFileSync(join(home, 'accounts', 'work', 'refresh_token'), 'synthetic-refresh-value', { mode: 0o600 });
  return home;
}

function seed(home: string, conversations: ReturnType<typeof aConversation>[], messages: ReturnType<typeof aMessage>[] = []): void {
  const db = openDatabase(databasePath(home, 'work'));
  try {
    upsertConversations(db, conversations);
    upsertMessages(db, messages);
  } finally {
    db.close();
  }
}

function seedPeople(home: string, people: Array<{ id: string; displayName: string | null; email: string | null }>): void {
  const db = openDatabase(databasePath(home, 'work'));
  try {
    upsertPeople(db, people);
  } finally {
    db.close();
  }
}

function messagePage(conversationId: string): string {
  const marker = ';messageid=';
  const separator = conversationId.lastIndexOf(marker);
  const path = separator < 0
    ? encodeURIComponent(conversationId)
    : `${encodeURIComponent(conversationId.slice(0, separator))}${marker}${encodeURIComponent(conversationId.slice(separator + marker.length))}`;
  return `${chatService}/v1/users/ME/conversations/${path}/messages?view=msnp24Equivalent&pageSize=200`;
}

function rawMessage(conversationId: string, id: string, clientMessageId: string | null, parentId?: string, senderName: string | null = 'Alex Example') {
  return {
    id,
    messagetype: 'RichText/Html',
    content: '<p>Sent safely</p>',
    conversationid: conversationId,
    ...(parentId === undefined ? {} : { conversationLink: `https://chat.example.test/v1/users/ME/conversations/${conversationId};messageid=${parentId}` }),
    from: `https://chat.example.test/v1/users/ME/contacts/8:orgid:${TEST_OID}`,
    ...(senderName === null ? {} : { imdisplayname: senderName }),
    composetime: '2026-01-01T10:00:00.000Z',
    version: '1767261600000',
    ...(clientMessageId === null ? {} : { clientmessageid: clientMessageId }),
    properties: {},
  };
}

function authenticatedFetch(routes: FakeRoute[], resultMessage?: (clientMessageId: string) => unknown) {
  const fake = fakeFetch([
    { method: 'POST', url: AAD_TOKEN_URL, json: tokenResponse(fakeJwt({ oid: TEST_OID, tid: TEST_TID }), 'synthetic-rotated-value') },
    AUTHZ_ROUTE,
    ...routes,
  ]);
  let clientMessageId: string | null = null;
  const fetch: typeof globalThis.fetch = async (input, init) => {
    const url = String(input);
    const method = (init?.method ?? 'GET').toUpperCase();
    if (method === 'POST' && /\/conversations\/.*\/messages$/.test(url)) {
      clientMessageId = (JSON.parse(String(init?.body)) as { clientmessageid: string }).clientmessageid;
    }
    const response = await fake.fetch(input, init);
    if (method === 'GET' && resultMessage && clientMessageId !== null) {
      return new Response(JSON.stringify({ messages: [resultMessage(clientMessageId)], _metadata: {} }), { status: response.status });
    }
    return response;
  };
  return { fetch, calls: fake.calls };
}

function chatCalls(calls: ReturnType<typeof fakeFetch>['calls']) {
  return calls.filter(({ url }) => url.startsWith(chatService));
}
